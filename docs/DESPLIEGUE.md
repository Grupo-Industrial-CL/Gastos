# Despliegue en IIS (Windows)

Mismo esquema que MasterPlan v2 (ver su `docs/DESPLIEGUE_IIS.md` para el paso a paso detallado).

## Entornos

| Entorno | `ASPNETCORE_ENVIRONMENT` | Host | BD |
|---|---|---|---|
| Desarrollo | `Development` | `https://localhost:7444` | `MasterPlan_DES` (user-secrets) |
| Test | `Staging` | `https://gastos-des.clgrupoindustrial.com` | `MasterPlan_DES` (variable de entorno) |
| Producción | `Production` | `https://gastos.clgrupoindustrial.com` | BD de producción (variable de entorno) |

Las tres URIs (`/signin-oidc` y `/signout-callback-oidc` de cada host) están solicitadas en el registro de aplicación de Entra ID.

## Requisitos del servidor
- **.NET 8 Hosting Bundle** (módulo ASP.NET Core). Tras instalarlo: `iisreset`.
- Característica **WebSocket Protocol** (Blazor Server usa SignalR).
- App Pool en modo *No Managed Code*.
- Acceso de red a SQL Server (`CLAPP107v.clay.local\DESARROLLO,1444` en pruebas) y salida HTTPS a las APIs de Repsol/Moeve.

## Publicar
```powershell
dotnet publish src/Gastos.Web/Gastos.Web.csproj -c Release -o C:\publish\gastos
```
Copiar el contenido a la carpeta física del sitio en IIS.

## Variables de entorno (en el servidor, variables del sistema; dos guiones bajos `__`)
| Variable | Valor |
|---|---|
| `ASPNETCORE_ENVIRONMENT` | `Staging` (pruebas) o `Production` |
| `ConnectionStrings__Gastos` | Cadena de conexión completa con contraseña |
| `AzureAd__ClientSecret` | Secreto del registro de aplicación |
| `Proveedores__Repsol__ApiKey` / `Proveedores__Moeve__ApiKey` | Credenciales de las APIs (cuando existan) |

Tras fijarlas: `iisreset`.
