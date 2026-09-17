namespace Gastos.Domain.Models;

/// <summary>Rol y sociedades con acceso de un usuario del portal.</summary>
public class PermisosUsuario
{
    public string Email { get; set; } = string.Empty;
    public string? Rol { get; set; }
    public List<int> Sociedades { get; set; } = new();
    public bool TieneAcceso => !string.IsNullOrWhiteSpace(Rol);
    public bool EsAdministrador => string.Equals(Rol, Roles.Administrador, StringComparison.OrdinalIgnoreCase);
    public bool EsConsultor => string.Equals(Rol, Roles.Consultor, StringComparison.OrdinalIgnoreCase);
}
