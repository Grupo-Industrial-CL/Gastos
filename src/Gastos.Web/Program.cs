using System.Globalization;
using Gastos.Business;
using Gastos.Data;
using Gastos.Web.Components;
using Gastos.Web.Services;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Server;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;
using MudBlazor.Services;

var builder = WebApplication.CreateBuilder(args);

// Cultura es-ES por defecto (fechas dd/MM/yyyy, decimales con coma).
var culturaEs = new CultureInfo("es-ES");
CultureInfo.DefaultThreadCurrentCulture = culturaEs;
CultureInfo.DefaultThreadCurrentUICulture = culturaEs;

// ─────────────────────────────────────────────────────────────────────────────
// Autenticación.
//  - Normal: Azure AD / Entra ID (Microsoft.Identity.Web + OpenID Connect). Sección "AzureAd".
//  - SIMULADA (solo Development, Autenticacion:Simulada=true): página local /dev-login que emite la misma cookie
//    con los mismos claims, para poder avanzar mientras no exista el registro de aplicación en Azure.
// En ambos casos el ACCESO lo decide GT_UsuariosRoles (ver SesionUsuario y MainLayout).
// ─────────────────────────────────────────────────────────────────────────────
var autenticacionSimulada = builder.Configuration.GetValue<bool>("Autenticacion:Simulada");
if (autenticacionSimulada && !builder.Environment.IsDevelopment())
    throw new InvalidOperationException("Autenticacion:Simulada solo se admite en el entorno Development. Quita el flag o corrige ASPNETCORE_ENVIRONMENT.");
builder.Services.AddSingleton(new OpcionesAutenticacion(autenticacionSimulada));

if (autenticacionSimulada)
{
    builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
        .AddCookie(o =>
        {
            o.LoginPath = AutenticacionSimulada.RutaLogin;
            o.ExpireTimeSpan = TimeSpan.FromHours(8);
            o.SlidingExpiration = true;
        });
}
else
{
    builder.Services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
        .AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAd"));

    // Sesión con caducidad deslizante de 1 hora por inactividad (mismo criterio que MasterPlan v2).
    builder.Services.Configure<CookieAuthenticationOptions>(CookieAuthenticationDefaults.AuthenticationScheme, o =>
    {
        o.ExpireTimeSpan = TimeSpan.FromHours(1);
        o.SlidingExpiration = true;
    });
    builder.Services.Configure<OpenIdConnectOptions>(OpenIdConnectDefaults.AuthenticationScheme, o =>
    {
        o.SignedOutRedirectUri = "/";
        // Permite forzar el selector de cuenta al cerrar sesión (ver endpoint /account/logout).
        var anterior = o.Events.OnRedirectToIdentityProvider;
        o.Events.OnRedirectToIdentityProvider = async ctx =>
        {
            if (anterior is not null) await anterior(ctx);
            if (ctx.Properties.Items.TryGetValue("prompt", out var prompt) && !string.IsNullOrEmpty(prompt))
                ctx.ProtocolMessage.Prompt = prompt;
        };
    });
}

// Toda la aplicación exige usuario autenticado (no hay páginas públicas).
builder.Services.AddAuthorization(o => o.FallbackPolicy = o.DefaultPolicy);
builder.Services.AddCascadingAuthenticationState();
builder.Services.AddScoped<AuthenticationStateProvider, ServerAuthenticationStateProvider>();
var mvc = builder.Services.AddControllersWithViews();
if (!autenticacionSimulada) mvc.AddMicrosoftIdentityUI();

// UI: Blazor Server + MudBlazor
builder.Services.AddRazorComponents().AddInteractiveServerComponents();
builder.Services.AddMudServices(c =>
{
    c.SnackbarConfiguration.PositionClass = MudBlazor.Defaults.Classes.Position.BottomEnd;
    c.SnackbarConfiguration.PreventDuplicates = true;
    c.SnackbarConfiguration.ShowCloseIcon = true;
    c.SnackbarConfiguration.VisibleStateDuration = 5000;
    c.SnackbarConfiguration.SnackbarVariant = MudBlazor.Variant.Filled;
});
builder.Services.AddHttpContextAccessor();

// Capas: datos (ConnectionStrings:Gastos) y negocio (servicios + APIs de proveedores).
builder.Services.AddGastosData(builder.Configuration);
builder.Services.AddGastosBusiness(builder.Configuration);
builder.Services.AddScoped<SesionUsuario>();

// Carga programada diaria por API (Repsol/Moeve). Sección CargasProgramadas; deshabilitada por defecto.
builder.Services.Configure<CargasProgramadasOptions>(builder.Configuration.GetSection("CargasProgramadas"));
builder.Services.AddHostedService<CargaProgramadaBackgroundService>();

var app = builder.Build();

app.UseRequestLocalization(new RequestLocalizationOptions
{
    DefaultRequestCulture = new Microsoft.AspNetCore.Localization.RequestCulture(culturaEs),
    SupportedCultures = new List<CultureInfo> { culturaEs },
    SupportedUICultures = new List<CultureInfo> { culturaEs }
});

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseAntiforgery();
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();

// Cierre de sesión: limpia la cookie y vuelve al login con el selector de cuenta. Debe ser una petición
// HTTP real (no el circuito Blazor), por eso el layout navega aquí con forceLoad.
if (autenticacionSimulada)
{
    AutenticacionSimulada.MapearEndpoints(app);
    app.MapGet("/account/logout", async (HttpContext http) =>
    {
        await http.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
        return Results.Redirect(AutenticacionSimulada.RutaLogin);
    }).AllowAnonymous();
}
else
{
    app.MapGet("/account/logout", async (HttpContext http) =>
    {
        await http.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
        await http.ChallengeAsync(OpenIdConnectDefaults.AuthenticationScheme,
            new AuthenticationProperties(new Dictionary<string, string?> { ["prompt"] = "select_account" }) { RedirectUri = "/" });
    }).AllowAnonymous();
}

// La pantalla principal es Gastos: la raíz redirige por HTTP (evita la NavigationException del prerrenderizado
// que se produce al redirigir desde un componente durante el render estático).
app.MapGet("/", () => Results.Redirect("/gastos"));

app.MapRazorComponents<App>().AddInteractiveServerRenderMode();

app.Run();
