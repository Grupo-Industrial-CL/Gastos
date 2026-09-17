namespace Gastos.Domain.Entities;

/// <summary>Sociedad / empresa del grupo (tabla GT_Sociedades; en Informix "paises": el VB6 modelaba cada empresa como un "país").</summary>
public class Sociedad
{
    public int CodPais { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public string? NombreIngles { get; set; }
    public bool Activo { get; set; } = true;
}
