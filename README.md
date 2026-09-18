# Gastos de viaje — portal web

Portal web en **.NET 8 · Blazor Server · MudBlazor** que sustituye a las utilidades que se conservan de la
aplicación antigua **Gastos (Visual Basic 6 sobre Informix)**, cuyo código fuente está en `legacy/vb6/` como referencia.

- Inicio de sesión con **Azure AD / Entra ID** (Microsoft.Identity.Web). El acceso real lo da el **rol** del usuario
  en el portal: **Administrador** (ve todo, carga datos, configura usuarios) o **Consultor** (solo consulta los gastos de
  las sociedades a las que tiene acceso). Sin rol, el usuario ve una pantalla de "acceso no autorizado".
- Datos en **SQL Server** (`MasterPlan_DES` en desarrollo) en tablas con prefijo **`GT_`**; el directorio de usuarios es
  la tabla `MP_Usuarios` (compartida con MasterPlan, solo lectura).
- Misma arquitectura en capas que **MasterPlan v2** (Domain · Data · Business · Web), que se ha usado como referencia
  de estilo sin modificarlo.

## Pantallas

| Pantalla | Quién | Qué hace |
|---|---|---|
| **Gastos** (`/gastos`) | Administrador y Consultor | Grid de gastos internos con buscador general (lupa, resalta en amarillo), filtro bajo cada columna, ordenación, paginación y **Exportar a Excel** de lo que se ve. Filtro por periodo y sociedades. El Consultor solo ve sus sociedades. |
| **Cargar datos** (botón en Gastos) | Administrador | Origen **Repsol** o **Moeve** (API, entre fechas) o **LayTours** (subida del CSV para una sociedad). Carga "todo o nada" con resumen de insertados/omitidos. |
| **Historial de cargas** (`/cargas`) | Administrador | Registro de todas las cargas (correctas y fallidas). |
| **Configuración → Usuarios y permisos** (`/config/usuarios`) | Administrador | Asigna rol a cada usuario de Azure y, a los Consultores, las sociedades a las que acceden. |
| **Configuración → Departamentos / Empleados / Tarjetas de carburante** | Administrador | Mantenimiento de los catálogos que usan las cargas (alta y edición; bajas lógicas). |

Además: al pulsar una fila de Gastos se abre el **detalle** (gasto, movimiento de tarjeta y carga que lo generó); el diálogo
"Cargar datos" tiene un botón **Comprobar** que lista las tarjetas o usuarios Epsilon sin dar de alta antes de cargar; y existe
una **carga programada** diaria por API (sección `CargasProgramadas` de `appsettings.json`, deshabilitada por defecto).

## Estructura

```
Gastos.sln
src/
  Gastos.Domain/     Entidades GT_*, modelos, enumeraciones y contratos (servicios, ISqlExecutor, IProveedorTarjetaCarburante)
  Gastos.Data/       EF Core (DbContext mínimo) + Dapper con SQL parametrizado y transacciones
  Gastos.Business/   Servicios de negocio, cargas (parser LayTours, clasificador de tarjetas) y clientes de API Repsol/Moeve
  Gastos.Web/        Blazor Server + MudBlazor + Azure AD
tests/Gastos.Tests/  Tests xUnit de la lógica portada del VB6
db/                  Scripts SQL: creación de tablas GT_ y migración desde Informix
docs/                ARQUITECTURA.md · PENDIENTES.md · DESPLIEGUE.md · PLAN_PRUEBAS.md
legacy/vb6/          Código fuente original en Visual Basic 6 (solo referencia)
```

## Puesta en marcha (desarrollo)

1. Crear las tablas: ejecutar `db/01_esquema.sql y db/02_datos.sql` en `MasterPlan_DES` y dar de alta al primer Administrador
   (bloque comentado al final del script).
2. Configurar la cadena de conexión (no se commitea):
   ```bash
   cd src/Gastos.Web
   dotnet user-secrets set "ConnectionStrings:Gastos" "Server=CLAPP107v.clay.local\DESARROLLO,1444;Database=MasterPlan_DES;User Id=bd_masterplan_o;Password=****;TrustServerCertificate=True"
   ```
3. Ejecutar:
   ```bash
   dotnet run --project src/Gastos.Web
   ```
   y abrir `https://localhost:7444`.

### Autenticación en desarrollo (mientras no exista el registro de Azure)

`appsettings.Development.json` trae `Autenticacion:Simulada = true`: en lugar de Azure se abre la página local
`/dev-login`, donde eliges el correo (`usId` de `MP_Usuarios`) con el que entrar, sin contraseña. A partir de ahí todo
es idéntico al modo real: el rol y las sociedades se leen de `GT_UsuariosRoles` / `GT_UsuariosSociedades`, y un correo
sin rol ve la pantalla de "acceso no autorizado". El modo simulado **solo se admite con `ASPNETCORE_ENVIRONMENT=Development`**;
en cualquier otro entorno la aplicación no arranca.

Para probar el login real de Azure cuando infra entregue el registro: poner `Simulada` a `false` y configurar
```bash
dotnet user-secrets set "AzureAd:ClientSecret" "<secreto>"
```
(TenantId y ClientId van en `appsettings.json`, sección `AzureAd`; el secreto y la contraseña de BD nunca se commitean.)

Tests: `dotnet test`.

Ver `docs/PLAN_PRUEBAS.md` para las pruebas a ejecutar con datos reales y `docs/PENDIENTES.md` para lo que falta por cerrar (registro de aplicación en Azure, APIs de Repsol/Moeve, migración del histórico).
