using Gastos.Domain.Models;

namespace Gastos.Domain.Abstractions.Services;

/// <summary>Consulta de gastos internos para el grid.</summary>
public interface IGastoService
{
    /// <summary>
    /// Devuelve los gastos de las sociedades indicadas (null = todas; lista vacía = ninguno) entre las fechas dadas.
    /// El filtrado por sociedades es la barrera de seguridad de los Consultores: la UI SIEMPRE pasa sus sociedades permitidas.
    /// </summary>
    Task<IReadOnlyList<GastoVista>> DameGastosAsync(IReadOnlyCollection<int>? sociedades, DateTime? desde, DateTime? hasta, CancellationToken ct = default);

    /// <summary>Detalle de un gasto (registro completo, movimiento de tarjeta y carga). Nulo si no existe o no está entre las sociedades permitidas.</summary>
    Task<GastoDetalle?> DameDetalleAsync(long contador, IReadOnlyCollection<int>? sociedades, CancellationToken ct = default);
}
