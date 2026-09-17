namespace Gastos.Domain.Entities;

/// <summary>Registro de auditoría de cada carga ejecutada (GT_Cargas).</summary>
public class Carga
{
    public int Id { get; set; }
    public string Origen { get; set; } = string.Empty;
    public DateTime Fecha { get; set; }
    public string Usuario { get; set; } = string.Empty;
    public DateTime? Desde { get; set; }
    public DateTime? Hasta { get; set; }
    public int? CodPais { get; set; }
    public string? NombreFichero { get; set; }
    public int Filas { get; set; }
    public int Insertados { get; set; }
    public int Omitidos { get; set; }
    public string Estado { get; set; } = "OK";   // OK | Error
    public string? Mensaje { get; set; }
}
