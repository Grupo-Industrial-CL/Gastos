namespace Gastos.Domain.Entities;

/// <summary>
/// Usuario sincronizado desde Azure AD (tabla MP_Usuarios, compartida con MasterPlan, solo lectura).
/// La identidad es el correo @gicl.es (usId). Incluye el rol asignado en GT_UsuariosRoles (nulo = sin acceso)
/// y el número de sociedades a las que tiene acceso.
/// </summary>
public class MPUsuario
{
    public string usId { get; set; } = string.Empty;
    public Guid usObjectId { get; set; }
    public string usDisplayName { get; set; } = string.Empty;
    public string usFirstName { get; set; } = string.Empty;
    public string usLastName { get; set; } = string.Empty;
    public string? usUserMail { get; set; }
    public string? usCompanyName { get; set; }
    public string? usDepartment { get; set; }
    public string? usJobTitle { get; set; }
    public string? usManagerId { get; set; }
    public string? usManagerDisplayName { get; set; }
    public bool usIsActive { get; set; }

    /// <summary>Rol en el portal de Gastos (GT_UsuariosRoles). Nulo si no tiene acceso.</summary>
    public string? Rol { get; set; }
    /// <summary>Número de sociedades con acceso (GT_UsuariosSociedades).</summary>
    public int NumSociedades { get; set; }
}
