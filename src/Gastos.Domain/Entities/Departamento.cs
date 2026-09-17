namespace Gastos.Domain.Entities;

/// <summary>Departamento (GT_Departamentos; Informix "departamentoscl"). Clave: CodPais + Codigo.</summary>
public class Departamento
{
    public int CodPais { get; set; }
    public int Codigo { get; set; }
    public string Descripcion { get; set; } = string.Empty;
    public string? CECO { get; set; }
    public string? Divisa { get; set; }
    public bool Activo { get; set; } = true;
}
