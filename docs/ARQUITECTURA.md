# Arquitectura del portal de Gastos de viaje

## Origen

La aplicación **Gastos** (VB6, base de datos Informix `crislaym`, DSN GALILEO) gestionaba los gastos de viaje de la red
comercial y de los empleados internos de las sociedades del grupo. De ella se conservan **solo dos utilidades**, ambas
del bloque *Gastos internos*:

| VB6 | Portal |
|---|---|
| `frmCargarFicheroORBIS.frm` — "Cargar fichero LayTours" (CSV en `GeneracionCSV\LayTours.csv`) | Carga **LayTours**: el administrador sube el CSV y elige la sociedad. |
| `frmCargaFicheroSolred2.frm` — "Carga fichero Solred" (CSV en `GeneracionCSV\Solred.csv`) | Carga **Repsol** por **API** (misma clasificación de movimientos) y, nueva, carga **Moeve** por API. |

El resto (gastos de red comercial, informes Data Report/Crystal, validación, pagos, contabilización en SAP...) queda fuera.

## Capas (mismo planteamiento que MasterPlan v2)

```
Gastos.Web  (Blazor Server + MudBlazor + Entra ID)
     │
Gastos.Business  (servicios: usuarios, permisos, catálogos, gastos, cargas; clientes API Repsol/Moeve)
     │
Gastos.Data  (GastosDbContext mínimo + SqlExecutor: Dapper con SQL parametrizado y transacciones)
     │
Gastos.Domain  (entidades GT_*, modelos, enums, contratos)  ← sin dependencias
```

- **Domain** define `ISqlExecutor` (consultas/comandos parametrizados y `EnTransaccionAsync`), los servicios
  `IUsuarioService`, `IPermisosService`, `ICatalogoService`, `IGastoService`, `ICargaService` y el contrato de proveedor
  de tarjetas `IProveedorTarjetaCarburante`.
- **Business** registra los servicios por convención (`IXxxService -> XxxService`) y los dos clientes de API como
  `IProveedorTarjetaCarburante` (Repsol, Moeve).
- **Web** aporta `SesionUsuario` (identidad + rol + sociedades por circuito), el layout con el *gate* de acceso, el grid
  reutilizable `GtGrid` y las páginas.

## Modelo de datos (SQL Server, prefijo `GT_`)

| Tabla | Origen Informix | Uso |
|---|---|---|
| `GT_Sociedades` | `paises` | Sociedades (el VB6 llamaba "país" a cada empresa). Sembrada desde `Constantes.bas`. |
| `GT_GastosSociedades` | `gastos_empresas` | Empresas de gasto; su `Id` es el "Tipo" de tipos de gasto para sociedades no estándar. |
| `GT_Departamentos` | `departamentoscl` | Departamentos por sociedad. |
| `GT_UsuariosPersonal` | `usuarios_personal` | Empleados. `Id_Epsilon` es la clave que usa LayTours. |
| `GT_TiposGasto` | `tiposgasto` | Catálogo de tipos de gasto (LayTours busca "Billetes"/"Hoteles"/"Otros gastos"; las tarjetas usan 1/9/10/11). |
| `GT_GastosInternos` | `gastos_internos` | **Tabla principal.** `Origen` (LayTours/Repsol/Moeve) e `IdCarga` son nuevas. `CodPais < 0` = borrado lógico. |
| `GT_GastosSolred` / `GT_GastosMoeve` | `gastos_solred` | Detalle de cada movimiento de tarjeta, enlazado al gasto por `Cont_GastoInterno`. |
| `GT_TarjetasSolred` | `tarjeta_solred` | Tarjetas de carburante → empleado. `Proveedor` (Repsol/Moeve) es nueva. |
| `GT_UsuariosRoles` | — | Rol del usuario de Azure (`Email` = `MP_Usuarios.usId`): Administrador o Consultor. |
| `GT_UsuariosSociedades` | — | Sociedades a las que accede un Consultor. |
| `GT_Cargas` | — | Auditoría de cada carga (origen, usuario, periodo/fichero, filas, insertados, omitidos, estado, mensaje). |

## Seguridad y acceso

1. **Autenticación**: OpenID Connect contra Entra ID (`Microsoft.Identity.Web`). Toda la aplicación exige usuario
   autenticado (`FallbackPolicy`). Cookie con caducidad deslizante de 1 hora.
2. **Identidad**: el correo del token (`preferred_username`) o el `oid` se resuelven contra `MP_Usuarios`
   (por `usId`, `usUserMail` u `usObjectId`, porque el correo del token varía entre entornos).
3. **Autorización**: `GT_UsuariosRoles`. Sin fila → pantalla "Acceso no autorizado" y solo puede cerrar sesión.
   - *Administrador*: todas las sociedades, botón "Cargar datos", historial y Configuración.
   - *Consultor*: solo Gastos y solo de sus sociedades (`GT_UsuariosSociedades`). El filtro se aplica **en la consulta
     SQL** (la UI pasa siempre la intersección entre lo seleccionado y lo permitido), y las rutas `/config/*` se bloquean
     también por URL en `MainLayout`.
4. **Secretos** fuera del repositorio: `ConnectionStrings:Gastos`, `AzureAd:ClientSecret`, `Proveedores:*:ApiKey`
   (user-secrets en desarrollo, variables de entorno en el servidor).

## Cargas

Todas son **"todo o nada"** (transacción; cualquier error → rollback) y dejan rastro en `GT_Cargas` (también las fallidas).

### LayTours (`LayToursParser` + `CargaService.CargarLayToursAsync`)
Port fiel de `cmdCargar_Click` del VB6: 9 campos separados por `;` (factura; concepto; importe; usuario Epsilon; albarán;
detalle; fecha `ddMMyyyy`; nº días; destino), validación campo a campo con el número de fila, concepto → familia
(Billetes / Hoteles / Otros gastos), empleado por `Id_Epsilon` en la sociedad elegida, tipo de gasto por
(`Tipo`, `NombreGasto`), e inserción en `GT_GastosInternos` con `VISA='F'`, `Validacion='Carga Orbis'`, `FamiliaGasto='O'`.

### Repsol / Moeve (`IProveedorTarjetaCarburante` + `ClasificadorCarburante` + `CargaService.CargarMovimientosAsync`)
1. Se piden los movimientos del periodo a la API (cliente `RepsolApiClient` / `MoeveApiClient`; contrato **pendiente**).
2. Comprobación previa de tarjetas: si alguna no está en `GT_TarjetasSolred`, se aborta listándolas (equivale al botón
   "Comprobar tarjetas" del VB6).
3. Idempotencia: un movimiento ya cargado (tarjeta + factura + fecha + hora + importe) se omite.
4. Si la tarjeta es interna (`Departamento <> -1`), tiene `Ins_Gasto_Int='S'` y el importe es positivo, se genera el
   gasto interno con la clasificación del VB6 (Gasolina por defecto; Parking / Peaje / Otros gastos según la descripción
   del producto; importe sin IVA redondeado a 2 decimales; `VISA='R'`, `Validacion='Carga SolRed'|'Carga Moeve'`).
5. Siempre se guarda el detalle del movimiento en `GT_GastosSolred` o `GT_GastosMoeve`.

`SolredCsvParser` conserva el lector del CSV antiguo: sirve para el **modo simulado** de las APIs en desarrollo
(`Proveedores:Repsol:Simulado=true` + `FicheroSimulacion`) y para cargar históricos si hiciera falta.

## Grid (`GtGrid<T>`)
Búsqueda global con resaltado amarillo (`MudHighlighter`), filtro por columna (texto "contiene" o comparación numérica
`>`, `<`, `=`) más embudo con los valores existentes, orden múltiple, paginación y exportación a Excel (OOXML sin
dependencias) de las filas filtradas, con los importes como número.
