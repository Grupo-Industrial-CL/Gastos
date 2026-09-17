using Gastos.Domain.Models;

namespace Gastos.Domain.Abstractions.Services;

/// <summary>Roles (GT_UsuariosRoles) y sociedades permitidas (GT_UsuariosSociedades) por usuario.</summary>
public interface IPermisosService
{
    Task<PermisosUsuario> DamePermisosAsync(string email, CancellationToken ct = default);
    /// <summary>Asigna el rol indicado; con rol nulo o vacío se retira el acceso (y sus sociedades).</summary>
    Task AsignarRolAsync(string email, string? rol, string usuarioAuditor, CancellationToken ct = default);
    Task EstablecerSociedadAsync(string email, int codPais, bool asignada, CancellationToken ct = default);
    Task EstablecerTodasSociedadesAsync(string email, bool asignar, CancellationToken ct = default);
}
