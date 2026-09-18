# Plan de pruebas — portal de Gastos de viaje

Lista de verificación para ejecutar **cuando tengamos acceso a Informix y datos reales**. Cada prueba indica pasos y
resultado esperado; marcar `[x]` al pasar y anotar incidencias al final. Orden recomendado: bloques 0 → 9.

Entorno: desarrollo (`https://localhost:7444`, autenticación simulada) contra `MasterPlan_DES`. Repetir los bloques 1, 2
y 8 en el servidor de test (`gastos-des`) con Azure real cuando infra entregue el registro de aplicación.

---

## 0. Validación de datos de Informix (antes de migrar)

- [ ] **0.1 Sociedades.** Comparar `SELECT codpais, nombre FROM paises` (Informix) con el bloque A de `db/02_datos.sql`.
  Esperado: mismos códigos y nombres; anotar sociedades que existan en Informix y no en la semilla (o al revés) y cuáles
  están inactivas. Corregir `02_datos.sql` antes de ejecutarlo.
- [ ] **0.2 Columnas reales de las 7 tablas.** Obtener el esquema (`INFO COLUMNS FOR tabla` o desde DBeaver) de `paises`,
  `gastos_empresas`, `departamentoscl`, `usuarios_personal`, `gastos_internos`, `gastos_solred`, `tarjeta_solred`, y también
  `tiposgasto`. Esperado: todas las columnas que usa `db/03_migracion_informix.sql` existen con ese nombre; anotar tipos y
  longitudes que superen las del esquema SQL Server (`NVARCHAR(50)` de `Donde`, `NVARCHAR(30)` de `Factura`, etc.).
- [ ] **0.3 Volumen.** `SELECT COUNT(*)` de cada tabla y rango `MIN/MAX(fecha)` de `gastos_internos` y `gastos_solred`.
  Esperado: cifras anotadas para contrastar tras la migración.
- [ ] **0.4 Borrados lógicos.** `SELECT COUNT(*) FROM gastos_internos WHERE codpais < 0`. Esperado: saber cuántos hay; el
  portal no los mostrará.
- [ ] **0.5 Tipos de gasto de LayTours.** `SELECT * FROM tiposgasto WHERE nombregasto IN ('Billetes','Hoteles','Otros gastos')`.
  Esperado: existen para `tipo = 'I'` y para cada `gastos_empresas.id` que use LayTours; anotar los códigos.
- [ ] **0.6 Tarjetas.** `SELECT departamento, ins_gasto_int, COUNT(*) FROM tarjeta_solred GROUP BY 1,2`. Esperado: saber
  cuántas son de red comercial (`-1`) y cuántas no generan gasto, para interpretar los "omitidos" de las cargas.
- [ ] **0.7 Enlace gasto ↔ movimiento.** `SELECT COUNT(*) FROM gastos_solred s WHERE NOT EXISTS (SELECT 1 FROM gastos_internos g
  WHERE g.contador = s.cont_gastored)`. Esperado: cifra de movimientos que quedarán con `Cont_GastoInterno` nulo (red comercial).

## 1. Migración (`db/03_migracion_informix.sql`)

- [ ] **1.1 Ejecución completa** en el orden del script, sin errores. Esperado: la consulta final de recuento coincide con 0.3
  (gastos internos: total menos ninguno, ya que se migran también los borrados lógicos con `CodPais < 0`).
- [ ] **1.2 Idempotencia.** Relanzar el script. Esperado: 0 filas insertadas en todos los bloques, mismos recuentos.
- [ ] **1.3 Contadores conservados.** `SELECT MAX(Contador) FROM GT_GastosInternos` = `MAX(contador)` de Informix y
  `DBCC CHECKIDENT` deja la identidad por encima. Insertar un gasto de prueba y borrarlo: recibe un contador nuevo mayor.
- [ ] **1.4 Muestreo de 20 gastos** elegidos al azar (`ORDER BY NEWID()`): comparar campo a campo con Informix (importe, fecha,
  donde, tipogasto, validado, pagado, factura). Esperado: idénticos; fechas sin desplazamiento de zona horaria.
- [ ] **1.5 Origen deducido.** `SELECT Origen, COUNT(*) FROM GT_GastosInternos GROUP BY Origen`. Esperado: `LayTours` para
  `Validacion = 'Carga Orbis'`, `Repsol` para `'Carga SolRed'`, `NULL` para el resto (manuales).
- [ ] **1.6 Empleados con Id_Epsilon.** Recuento de `GT_UsuariosPersonal` con `Id_Epsilon` no nulo igual al de Informix.
- [ ] **1.7 Caracteres especiales.** Buscar en `GT_UsuariosPersonal.Nombre` y `GT_GastosInternos.Donde` acentos y `ñ`.
  Esperado: se ven correctamente (no `Ã±`); si no, ajustar `CODEPAGE` / collation en la carga.
- [ ] **1.8 Truncamientos.** Ninguna fila rechazada por longitud; si el script se hizo por CSV, comparar longitudes máximas.

## 2. Acceso y roles

- [ ] **2.1 Sin rol.** Entrar por `/dev-login` con un correo de `MP_Usuarios` sin fila en `GT_UsuariosRoles`. Esperado:
  pantalla "Acceso no autorizado" con el correo, sin menú lateral, solo botón Cerrar sesión.
- [ ] **2.2 Administrador.** Entrar con el administrador. Esperado: chip "Administrador" junto al nombre; menú con Gastos,
  Historial de cargas y Configuración (Usuarios y permisos, Departamentos, Empleados, Tarjetas).
- [ ] **2.3 Consultor.** Desde Usuarios y permisos, asignar rol Consultor a un segundo usuario y 2 sociedades. Entrar con él.
  Esperado: sin chip, menú solo con Gastos, el grid muestra únicamente gastos de esas 2 sociedades y el recuento cuadra con
  `SELECT COUNT(*) ... WHERE CodPais IN (…) AND CodPais > 0 AND Fecha BETWEEN …`.
- [ ] **2.4 Consultor por URL.** Con el Consultor, teclear `/config/usuarios`, `/config/tarjetas` y `/cargas`. Esperado:
  aviso "No tienes acceso a esa pantalla" y vuelta a Gastos.
- [ ] **2.5 Consultor sin sociedades.** Quitarle todas las sociedades. Esperado: entra pero el grid está vacío.
- [ ] **2.6 Retirar acceso.** Poner "Sin acceso" al Consultor. Esperado: desaparecen sus sociedades (`GT_UsuariosSociedades`) y
  al volver a entrar ve "Acceso no autorizado".
- [ ] **2.7 Autoprotección.** Como administrador, intentar cambiarse a sí mismo a Consultor o Sin acceso. Esperado: aviso y no
  se aplica.
- [ ] **2.8 Cierre de sesión.** Menú de usuario → Cerrar sesión. Esperado: vuelve a `/dev-login`; al pulsar atrás no se entra.
- [ ] **2.9 Modo simulado bloqueado fuera de Development.** Arrancar con `ASPNETCORE_ENVIRONMENT=Staging` y `Simulada=true`.
  Esperado: la aplicación no arranca y el mensaje indica el motivo.

## 3. Pantalla Gastos (con el histórico migrado)

- [ ] **3.1 Carga inicial.** Periodo por defecto = año en curso. Esperado: recuento y total coinciden con
  `SELECT COUNT(*), SUM(Importe) FROM GT_GastosInternos WHERE CodPais > 0 AND Fecha BETWEEN '01/01' AND hoy`.
- [ ] **3.2 Periodo amplio.** Poner desde 2015. Esperado: carga en menos de ~5 s con decenas de miles de filas; el grid pagina
  sin bloquearse; anotar tiempo y nº de filas.
- [ ] **3.3 Borrados lógicos.** Ningún gasto con `CodPais < 0` aparece.
- [ ] **3.4 Nombres resueltos.** Sociedad, Departamento, Empleado y Tipo de gasto muestran nombres, no códigos. Anotar los
  casos que muestren código (empleado dado de baja, tipo desconocido) y decidir si es aceptable.
- [ ] **3.5 Buscador.** Escribir parte del nombre de un empleado. Esperado: solo sus filas y el texto resaltado en amarillo en
  todas las columnas donde aparezca; borrar con la X restaura.
- [ ] **3.6 Filtro de texto por columna** (Concepto). Esperado: filtra "contiene" sin distinguir mayúsculas.
- [ ] **3.7 Embudo de valores** (Sociedad): marcar 2 valores. Esperado: solo esas sociedades; el icono cambia de color; la
  caja muestra los valores marcados en solo lectura; Limpiar del embudo la vacía.
- [ ] **3.8 Filtro numérico** (Importe): `>100`, `<=50`, `=0`. Esperado: compara sobre el importe, no sobre el texto.
- [ ] **3.9 Filtro de fecha** en columna y ordenación por Fecha e Importe (asc/desc, y orden múltiple con Ctrl).
- [ ] **3.10 Limpiar filtros.** Vacía buscador y todos los filtros de columna de golpe.
- [ ] **3.11 Refrescar** tras insertar un gasto por SQL. Esperado: aparece sin recargar la página.
- [ ] **3.12 Exportar a Excel** con filtros activos. Esperado: el fichero tiene exactamente las filas visibles, mismas
  columnas que el grid, Importe como número (se puede sumar) y fechas legibles; abre sin avisos en Excel.
- [ ] **3.13 Scroll horizontal** con todas las columnas y cabecera fija al bajar.
- [ ] **3.14 Modo oscuro.** Grid, chips y resaltado legibles.

## 4. Detalle del gasto

- [ ] **4.1 Gasto de LayTours** (Origen LayTours). Esperado: ficha sin bloque de tarjeta, con factura y nº de días si es hotel.
- [ ] **4.2 Gasto de Solred migrado.** Esperado: bloque "Movimiento de tarjeta Repsol" con tarjeta, conductor, litros, importe
  con/sin IVA coherentes con `GT_GastosSolred`.
- [ ] **4.3 Gasto manual** (Origen nulo). Esperado: pie "origen desconocido (migrado)", sin carga asociada.
- [ ] **4.4 Consultor** abre un detalle de su sociedad (funciona) y, por URL/manipulación, uno de otra sociedad no es posible
  (el servicio devuelve nulo → aviso "no tienes acceso").

## 5. Carga LayTours (con un CSV real de LayTours)

- [ ] **5.1 Comprobar con fichero correcto.** Esperado: "Fichero correcto: N filas…", botón Cargar habilitado.
- [ ] **5.2 Comprobar con usuario Epsilon desconocido.** Esperado: tabla con el id, nº de filas e importe; Cargar deshabilitado.
  Darlo de alta en Configuración → Empleados (campo Id Epsilon), volver a Comprobar: desaparece.
- [ ] **5.3 Errores de formato** (probar cada uno editando el CSV): 8 columnas, concepto no admitido, importe con letras, fecha
  31/02, línea vacía intermedia. Esperado: mensaje con el campo y el nº de fila, nada guardado, fila en Historial con Error.
- [ ] **5.4 Carga correcta.** Esperado: resumen con insertados = filas; en Gastos aparecen con Origen LayTours,
  `Validacion='Carga Orbis'`, `VISA='F'`, `FamiliaGasto='O'`, Moneda EUR; hotel con Numero = días y Numero2 = 1; billete con
  Donde = albarán/destino; otros con Donde = albarán/detalle, truncado a 50.
- [ ] **5.5 Comparación con el VB6.** Cargar el mismo CSV en la aplicación antigua (entorno de pruebas) y comparar los
  registros generados uno a uno. Esperado: idénticos salvo `Origen`/`IdCarga`.
- [ ] **5.6 Doble carga.** Cargar el mismo fichero dos veces. Esperado: se duplica (LayTours no deduplica, igual que el VB6);
  decidir si queremos control de duplicados por factura.
- [ ] **5.7 Codificación.** CSV con acentos en detalle/destino. Esperado: se guardan bien.
- [ ] **5.8 Sociedad no estándar** (`Tipo` de `GT_GastosSociedades`). Esperado: busca los tipos de gasto con ese `Tipo`.

## 6. Carga Repsol / Moeve

### 6.a Modo simulado (CSV antiguo de Solred), antes de tener API
- [ ] **6.1 Configurar** `Proveedores:Repsol:Simulado=true` y `FicheroSimulacion` con un Solred.csv real de un mes.
  Esperado: en "Cargar datos" la opción Repsol aparece habilitada con "Simulado (fichero …)".
- [ ] **6.2 Comprobar.** Esperado: nº de movimientos del periodo; lista de tarjetas sin dar de alta (si las hay) con conductor,
  matrícula e importe; avisos de tarjetas de red comercial y de tarjetas sin "generar gasto".
- [ ] **6.3 Alta de tarjeta faltante** en Configuración → Tarjetas (Repsol, sociedad, departamento, empleado). Volver a
  Comprobar: desaparece de la lista.
- [ ] **6.4 Carga.** Esperado: insertados = movimientos de tarjetas internas con `Ins_Gasto_Int='S'` e importe > 0; omitidos =
  el resto; en `GT_GastosSolred` hay una fila por movimiento con `Cont_GastoInterno` relleno solo en los que generaron gasto.
- [ ] **6.5 Clasificación.** Revisar en Gastos: producto con "aparcamiento" → Parking (11) y Donde "Parking Población (info)";
  "autopista"/"estradas"/"interoperable" → Peaje (9) con Numero 1; "lavados" → Otros gastos (10); resto → Gasolina (1).
  Importe = total / 1,21 redondeado a 2 decimales (comprobar 5 casos a mano).
- [ ] **6.6 Comparación con el VB6** para el mismo CSV. Esperado: mismos gastos internos (los de red comercial no, por alcance).
- [ ] **6.7 Idempotencia.** Repetir la carga del mismo periodo. Esperado: insertados 0, omitidos = todos, sin duplicados.
- [ ] **6.8 Km > 9999** y descuentos (importe negativo). Esperado: Km guardado como 0; el negativo se guarda en detalle sin gasto.

### 6.b API real (cuando llegue la documentación)
- [ ] **6.9 Configuración** de URL y credenciales por user-secrets; la opción aparece "Configurado".
- [ ] **6.10 Comprobar** un periodo de 7 días. Esperado: nº de movimientos coherente con el portal del proveedor.
- [ ] **6.11 Mapeo de campos.** Cotejar 5 movimientos con el extracto del proveedor: tarjeta, fecha/hora, población, producto,
  litros, IVA, total, precio litro, matrícula, conductor. Ajustar `TraducirAsync` si algún campo no cuadra.
- [ ] **6.12 Carga real** de un mes ya cargado antes en el VB6 vía CSV. Esperado: todo omitido si la clave de duplicado
  (tarjeta+factura+fecha+hora+importe) coincide; si la API usa otra numeración de factura, revisar la clave.
- [ ] **6.13 Errores de API** (credencial incorrecta, timeout). Esperado: mensaje claro, fila Error en Historial, nada guardado.
- [ ] **6.14 Moeve**: repetir 6.9 a 6.13; tarjetas con `Proveedor='Moeve'`; detalle en `GT_GastosMoeve`;
  `Validacion='Carga Moeve'`.

## 7. Historial de cargas

- [ ] **7.1** Cada carga (OK y Error) de los bloques 5 y 6 aparece con origen, usuario, periodo o fichero, filas, insertados,
  omitidos, estado y mensaje. Filtro y Excel funcionan.
- [ ] **7.2** Una carga fallida no deja filas en `GT_GastosInternos` ni en el detalle con su `IdCarga`.

## 8. Configuración (mantenimiento)

- [ ] **8.1 Departamentos.** Alta (código duplicado → error), edición, desactivar y ver que desaparece salvo con "Mostrar
  inactivos". El departamento nuevo aparece en el desplegable de Empleados y Tarjetas.
- [ ] **8.2 Empleados.** Alta con Id Epsilon, edición, fecha de baja (desaparece de la lista y del selector de tarjetas).
- [ ] **8.3 Tarjetas.** Alta Repsol y Moeve; al elegir empleado se rellenan jerarquía y titular; número con ceros a la izquierda
  o decimales se normaliza; duplicada → error; "Red comercial" fuerza "no genera gasto"; desactivar.
- [ ] **8.4 Usuarios y permisos.** Interruptores "Solo con acceso" y "Mostrar bajas"; Excel del directorio; contador de
  sociedades se actualiza al cerrar el diálogo.

## 9. Carga programada y despliegue

- [ ] **9.1 Programada.** `CargasProgramadas:Habilitado=true`, `Hora` = dentro de 2 minutos, `DiasAtras=7`. Esperado: en el log
  "próxima ejecución en …"; a la hora, fila en Historial con usuario "Carga programada"; con API sin configurar, aviso en
  log y sin fila.
- [ ] **9.2 Servidor de test (gastos-des).** Publicación con `ASPNETCORE_ENVIRONMENT=Staging`, variables
  `ConnectionStrings__Gastos` y `AzureAd__ClientSecret`; login real con Azure; usuario cuyo token trae el UPN
  `@clgrupoindustrial.com` se resuelve igualmente (por `usUserMail`/`oid`).
- [ ] **9.3 Caducidad de sesión** (1 h de inactividad en Azure). Esperado: vuelve al login.
- [ ] **9.4 Rotación de secretos.** Contraseña de `bd_masterplan_o` cambiada tras haber circulado por chat; credencial de
  Informix de `Inicio.bas` cambiada.

---

## Incidencias detectadas

| # | Prueba | Descripción | Estado |
|---|--------|-------------|--------|
|   |        |             |        |
