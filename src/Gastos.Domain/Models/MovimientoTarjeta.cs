namespace Gastos.Domain.Models;

/// <summary>
/// Movimiento de tarjeta de carburante tal y como lo devuelve el proveedor (API de Repsol/Moeve) o el CSV
/// antiguo de Solred. Es el formato neutro que consume la carga; cada cliente de API traduce su JSON a esto.
/// </summary>
public class MovimientoTarjeta
{
    public string NumTarjeta { get; set; } = string.Empty;
    public int AnyoFactura { get; set; }
    public string Factura { get; set; } = string.Empty;
    public string? Matricula { get; set; }
    public string? Conductor { get; set; }
    public DateTime FechaOperacion { get; set; }
    /// <summary>Hora en formato HH:mm (o HH:mm:ss).</summary>
    public string HoraOperacion { get; set; } = "00:00";
    public string? Poblacion { get; set; }
    public decimal Km { get; set; }
    public string? DescripcionProducto { get; set; }
    public decimal Litros { get; set; }
    /// <summary>Porcentaje de IVA (p. ej. 21).</summary>
    public decimal IVA { get; set; }
    public decimal ImporteTotal { get; set; }
    public decimal PrecioLitro { get; set; }
    public string? InfoAux { get; set; }
}
