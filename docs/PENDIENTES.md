# Pendientes para poner el portal en marcha

## 🔴 Bloquean el uso real

> Mientras no llegue el punto 1, en desarrollo se entra con la **autenticación simulada** (`/dev-login`, ver README).
> Quitar `Autenticacion:Simulada` de `appsettings.Development.json` cuando el registro de Azure esté disponible.

1. **Registro de aplicación en Entra ID** (Infra, ticket solicitado). Hace falta `TenantId`, `ClientId` y `ClientSecret`.
   URIs de redirección y de cierre de sesión pedidas:

   | Entorno | Redirección | Cierre de sesión |
   |---|---|---|
   | Desarrollo | `https://localhost:7444/signin-oidc` | `https://localhost:7444/signout-callback-oidc` |
   | Test | `https://gastos-des.clgrupoindustrial.com/signin-oidc` | `https://gastos-des.clgrupoindustrial.com/signout-callback-oidc` |
   | Producción | `https://gastos.clgrupoindustrial.com/signin-oidc` | `https://gastos.clgrupoindustrial.com/signout-callback-oidc` |

   TenantId/ClientId van en `appsettings.json` (sección `AzureAd`); el secreto en user-secrets (`AzureAd:ClientSecret`)
   o variable de entorno `AzureAd__ClientSecret`.
2. **Cadena de conexión** con contraseña: `ConnectionStrings:Gastos` en user-secrets / `ConnectionStrings__Gastos`.
   Recomendado rotar la contraseña de `bd_masterplan_o`, que ha circulado por chat.
3. **Crear las tablas** con `db/01_esquema.sql y db/02_datos.sql` y dar de alta al **primer Administrador** (bloque final del script).

## 🟠 Necesarios para las cargas

4. **APIs de Repsol y Moeve**: documentación (URL, autenticación, formato de respuesta). El cliente genérico
   (`ProveedorApiBase`) hace un GET con Bearer/API key y espera un array JSON de movimientos; adaptar
   `TraducirAsync` en `RepsolApiClient`/`MoeveApiClient` al contrato real. Mientras tanto, las opciones aparecen
   deshabilitadas en "Cargar datos" con el motivo.
5. **Migración del histórico desde Informix** (`db/03_migracion_informix.sql`): validar nombres/tipos de columna reales
   cuando se tenga acceso, decidir servidor vinculado u exportación a CSV, y ejecutar en el orden indicado. Sin los
   `GT_TiposGasto` de LayTours ("Billetes", "Hoteles", "Otros gastos" por `Tipo`) la carga LayTours falla con
   "No se ha encontrado el tipo de gasto".
6. **Tarjetas Moeve**: dar de alta en `GT_TarjetasSolred` con `Proveedor='Moeve'` (no existían en el VB6).

## 🟡 Decisiones / mejoras

7. **Alcance**: solo gastos internos. Los movimientos de tarjetas de la red comercial (`Departamento = -1`) se guardan
   como detalle sin generar gasto (el VB6 los insertaba en `gastos`, con conversión a pesetas).
8. ~~Mantenimiento de catálogos desde el portal~~ Hecho: Configuración → Departamentos / Empleados / Tarjetas.
9. ~~Programar las cargas por API~~ Hecho: `CargasProgramadas` en `appsettings.json` (activar con `Habilitado=true`
   cuando las APIs estén configuradas; hora y días hacia atrás configurables).
10. **Credenciales en el código VB6**: `legacy/vb6/GastosVB6/Inicio.bas` contiene usuario y contraseña de Informix y ya
    está en el historial de git. Cambiar esa contraseña en Informix y valorar limpiar el historial.
