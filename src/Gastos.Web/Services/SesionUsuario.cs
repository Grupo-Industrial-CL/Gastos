using System.Security.Claims;
using Gastos.Domain.Abstractions.Services;
using Gastos.Domain.Models;
using Microsoft.AspNetCore.Components.Authorization;

namespace Gastos.Web.Services;

/// <summary>
/// Contexto de sesión (scoped por circuito Blazor): resuelve al usuario de Azure contra MP_Usuarios y carga su
/// rol y sociedades del portal (GT_UsuariosRoles / GT_UsuariosSociedades). Sin rol = sin acceso (gate en MainLayout).
/// </summary>
public sealed class SesionUsuario
{
    private readonly AuthenticationStateProvider _auth;
    private readonly IUsuarioService _usuarios;
    private readonly IPermisosService _permisos;
    private readonly SemaphoreSlim _gate = new(1, 1);

    public SesionUsuario(AuthenticationStateProvider auth, IUsuarioService usuarios, IPermisosService permisos)
    {
        _auth = auth;
        _usuarios = usuarios;
        _permisos = permisos;
    }

    public bool Cargado { get; private set; }
    /// <summary>Correo de identidad (usId de MP_Usuarios) o, si no está en el directorio, el del token.</summary>
    public string Email { get; private set; } = string.Empty;
    public string Nombre { get; private set; } = string.Empty;
    public string Empresa { get; private set; } = string.Empty;
    public bool EnDirectorio { get; private set; }
    public PermisosUsuario Permisos { get; private set; } = new();

    public bool TieneAcceso => Permisos.TieneAcceso;
    public bool EsAdministrador => Permisos.EsAdministrador;
    public bool EsConsultor => Permisos.EsConsultor;
    public string Rol => Permisos.Rol ?? string.Empty;
    /// <summary>Sociedades que puede ver: null = todas (Administrador); lista = las asignadas (Consultor).</summary>
    public IReadOnlyCollection<int>? SociedadesVisibles => EsAdministrador ? null : Permisos.Sociedades;

    public async Task EnsureLoadedAsync()
    {
        if (Cargado) return;
        await _gate.WaitAsync();
        try
        {
            if (Cargado) return;
            var user = (await _auth.GetAuthenticationStateAsync()).User;
            var email = user.FindFirst("preferred_username")?.Value
                        ?? user.FindFirst(ClaimTypes.Email)?.Value
                        ?? user.FindFirst(ClaimTypes.Upn)?.Value
                        ?? string.Empty;
            var oid = user.FindFirst("oid")?.Value
                      ?? user.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value;
            Email = email;
            Nombre = user.Identity?.Name ?? email;

            try
            {
                var u = await _usuarios.DameUsuarioAsync(email, oid);
                if (u is not null)
                {
                    EnDirectorio = true;
                    Email = string.IsNullOrWhiteSpace(u.usId) ? email : u.usId;
                    if (!string.IsNullOrWhiteSpace(u.usDisplayName)) Nombre = u.usDisplayName;
                    Empresa = u.usCompanyName ?? string.Empty;
                }
            }
            catch { /* directorio no disponible: se sigue con los datos del token */ }

            try { Permisos = await _permisos.DamePermisosAsync(Email); }
            catch { Permisos = new PermisosUsuario { Email = Email }; }

            Cargado = true;
        }
        finally { _gate.Release(); }
    }

    /// <summary>Fuerza recarga (p. ej. tras cambiar permisos del propio usuario).</summary>
    public void Invalidar() => Cargado = false;
}
