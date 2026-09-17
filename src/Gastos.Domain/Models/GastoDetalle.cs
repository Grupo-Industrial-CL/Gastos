using Gastos.Domain.Entities;

namespace Gastos.Domain.Models;

/// <summary>Detalle de un gasto interno: la fila del grid más el movimiento de tarjeta (si lo hay) y la carga que lo generó.</summary>
public class GastoDetalle
{
    public GastoVista Gasto { get; set; } = new();
    public GastoInterno Registro { get; set; } = new();
    public MovimientoCarburante? Movimiento { get; set; }
    /// <summary>"Repsol" o "Moeve" si el gasto viene de una tarjeta.</summary>
    public string? ProveedorMovimiento { get; set; }
    public Carga? Carga { get; set; }
}
