using System.ComponentModel.DataAnnotations.Schema;

namespace Gastos.Domain.Entities;

/// <summary>Empleado que genera gastos (GT_UsuariosPersonal; Informix "usuarios_personal"). Clave: CodPais + Departamento + Codigo.</summary>
public class UsuarioPersonal
{
    public int CodPais { get; set; }
    public int Departamento { get; set; }
    public long Codigo { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public string? DNI { get; set; }
    public string Jerarquia { get; set; } = string.Empty;
    /// <summary>Identificador del empleado en Epsilon; es la clave con la que LayTours identifica al viajero ("User ID" del CSV).</summary>
    [Column("Id_Epsilon")] public string? IdEpsilon { get; set; }
    public string? Usuario { get; set; }
    [Column("Cuenta_Bancaria")] public string? CuentaBancaria { get; set; }
    public DateTime? FechaBaja { get; set; }
    [Column("Bloqueado_Por")] public string? BloqueadoPor { get; set; }
}
