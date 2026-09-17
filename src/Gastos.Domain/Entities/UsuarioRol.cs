namespace Gastos.Domain.Entities;

/// <summary>Rol de acceso de un usuario de Azure al portal (GT_UsuariosRoles). Email = MP_Usuarios.usId.</summary>
public class UsuarioRol
{
    public string Email { get; set; } = string.Empty;
    public string Rol { get; set; } = string.Empty;
    public DateTime FechaAlta { get; set; }
    public string? UsuarioAlta { get; set; }
}
