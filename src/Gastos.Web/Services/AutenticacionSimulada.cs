using System.Net;
using System.Security.Claims;
using System.Text;
using Gastos.Domain.Abstractions;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;

namespace Gastos.Web.Services;

/// <summary>Indica al resto de la aplicación si la autenticación es la simulada de desarrollo (para redirigir al login local y mostrar el aviso).</summary>
public sealed record OpcionesAutenticacion(bool Simulada);

/// <summary>
/// AUTENTICACIÓN SIMULADA — SOLO DESARROLLO. Sustituye a Entra ID mientras no exista el registro de aplicación:
/// una página local (/dev-login) donde se elige el correo (usId de MP_Usuarios) con el que entrar. Emite la misma
/// cookie y los mismos claims que dejaría Azure (preferred_username, name), así que roles, sociedades y gate de
/// acceso funcionan exactamente igual. Se activa con Autenticacion:Simulada=true y SOLO en el entorno Development:
/// en cualquier otro entorno el arranque falla a propósito.
/// </summary>
public static class AutenticacionSimulada
{
    public const string RutaLogin = "/dev-login";

    public static void MapearEndpoints(WebApplication app)
    {
        app.MapGet(RutaLogin, async (HttpContext http, ISqlExecutor sql, string? returnUrl, string? error) =>
        {
            // Lista de ayuda: usuarios que ya tienen rol en el portal (para no tener que recordar el correo).
            var conRol = new List<(string Email, string Rol, string Nombre)>();
            string? aviso = null;
            try
            {
                var filas = await sql.QueryAsync<UsuarioRolFila>(@"
                    SELECT TOP 100 r.Email, r.Rol, ISNULL(u.usDisplayName, '') AS Nombre
                    FROM dbo.GT_UsuariosRoles r LEFT JOIN dbo.MP_Usuarios u ON u.usId = r.Email
                    ORDER BY r.Rol, r.Email");
                conRol = filas.Select(f => (f.Email, f.Rol, f.Nombre)).ToList();
                if (conRol.Count == 0) aviso = "No hay ningún usuario con rol en GT_UsuariosRoles: ejecuta el bloque C (primer Administrador) de db/02_datos.sql.";
            }
            catch (Exception ex) { aviso = "No se ha podido consultar la base de datos: " + ex.Message; }

            var sb = new StringBuilder();
            sb.Append("<!DOCTYPE html><html lang=\"es\"><head><meta charset=\"utf-8\"><title>Acceso de desarrollo</title>")
              .Append("<style>body{font-family:Roboto,Segoe UI,sans-serif;background:#f5f6f8;margin:0;display:flex;justify-content:center;padding-top:8vh}")
              .Append(".card{background:#fff;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.12);padding:28px 32px;width:520px}")
              .Append("h1{font-size:1.25rem;margin:0 0 4px;color:#152238}.sub{color:#b45309;background:#fff7ed;border:1px solid #fed7aa;border-radius:6px;padding:8px 10px;font-size:.85rem;margin:12px 0 18px}")
              .Append("label{display:block;font-size:.85rem;color:#555;margin-top:12px}input,select{width:100%;padding:9px;border:1px solid #ccc;border-radius:6px;font-size:.95rem;box-sizing:border-box;margin-top:4px}")
              .Append("button{margin-top:18px;width:100%;padding:10px;border:0;border-radius:6px;background:#00695C;color:#fff;font-size:1rem;cursor:pointer}")
              .Append(".err{color:#b91c1c;font-size:.85rem;margin-top:10px}.hint{color:#777;font-size:.8rem;margin-top:14px}</style></head><body><div class=\"card\">")
              .Append("<h1>Gastos de viaje · acceso de desarrollo</h1>")
              .Append("<div class=\"sub\">Autenticación <b>simulada</b>: no se comprueba ninguna contraseña. Solo está disponible en el entorno Development; en test y producción se entra con la cuenta corporativa de Microsoft.</div>")
              .Append("<form method=\"post\" action=\"").Append(RutaLogin).Append("\">")
              .Append("<input type=\"hidden\" name=\"returnUrl\" value=\"").Append(WebUtility.HtmlEncode(returnUrl ?? "/")).Append("\">");
            if (conRol.Count > 0)
            {
                sb.Append("<label>Usuarios con rol en el portal<select name=\"emailLista\"><option value=\"\">— elegir —</option>");
                foreach (var u in conRol)
                    sb.Append("<option value=\"").Append(WebUtility.HtmlEncode(u.Email)).Append("\">")
                      .Append(WebUtility.HtmlEncode($"{u.Email}  ·  {u.Rol}{(u.Nombre.Length > 0 ? "  ·  " + u.Nombre : string.Empty)}")).Append("</option>");
                sb.Append("</select></label>");
            }
            sb.Append("<label>O escribe el correo (usId de MP_Usuarios)<input type=\"email\" name=\"email\" placeholder=\"usuario@gicl.es\"></label>");
            if (!string.IsNullOrEmpty(error)) sb.Append("<div class=\"err\">").Append(WebUtility.HtmlEncode(error)).Append("</div>");
            if (aviso is not null) sb.Append("<div class=\"err\">").Append(WebUtility.HtmlEncode(aviso)).Append("</div>");
            sb.Append("<button type=\"submit\">Entrar</button>")
              .Append("<div class=\"hint\">Un correo sin rol en GT_UsuariosRoles verá la pantalla de \"acceso no autorizado\", igual que ocurriría con Azure.</div>")
              .Append("</form></div></body></html>");
            return Results.Content(sb.ToString(), "text/html; charset=utf-8");
        }).AllowAnonymous();

        app.MapPost(RutaLogin, async (HttpContext http) =>
        {
            var form = await http.Request.ReadFormAsync();
            var email = (form["email"].ToString().Trim().Length > 0 ? form["email"].ToString() : form["emailLista"].ToString()).Trim();
            var returnUrl = form["returnUrl"].ToString();
            if (string.IsNullOrWhiteSpace(returnUrl) || !returnUrl.StartsWith('/') || returnUrl.StartsWith("//")) returnUrl = "/";
            if (string.IsNullOrWhiteSpace(email) || !email.Contains('@'))
                return Results.Redirect($"{RutaLogin}?returnUrl={Uri.EscapeDataString(returnUrl)}&error={Uri.EscapeDataString("Indica un correo válido.")}");

            // Mismos claims que dejaría Entra ID: preferred_username (correo) y name. Sin oid (no hay token de Azure).
            var identity = new ClaimsIdentity(CookieAuthenticationDefaults.AuthenticationScheme, ClaimTypes.Name, ClaimTypes.Role);
            identity.AddClaim(new Claim("preferred_username", email));
            identity.AddClaim(new Claim(ClaimTypes.Name, email));
            identity.AddClaim(new Claim(ClaimTypes.Email, email));
            identity.AddClaim(new Claim("auth_simulada", "1"));
            await http.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(identity),
                new AuthenticationProperties { IsPersistent = false });
            return Results.Redirect(returnUrl);
        }).AllowAnonymous().DisableAntiforgery();
    }

    private sealed class UsuarioRolFila
    {
        public string Email { get; set; } = string.Empty;
        public string Rol { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
    }
}
