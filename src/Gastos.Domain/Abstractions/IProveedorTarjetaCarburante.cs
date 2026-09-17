using Gastos.Domain.Models;

namespace Gastos.Domain.Abstractions;

/// <summary>
/// Cliente de la API de un proveedor de tarjetas de carburante (Repsol / Moeve). Cada implementación traduce
/// la respuesta del proveedor al formato neutro <see cref="MovimientoTarjeta"/>.
/// </summary>
public interface IProveedorTarjetaCarburante
{
    OrigenCarga Origen { get; }
    /// <summary>False si faltan datos de conexión (URL, credenciales): la UI deshabilita la opción.</summary>
    bool Configurado { get; }
    /// <summary>Texto para la UI: "Configurado", "Simulado" o el motivo por el que no lo está.</summary>
    string Detalle { get; }
    Task<IReadOnlyList<MovimientoTarjeta>> ObtenerMovimientosAsync(DateTime desde, DateTime hasta, CancellationToken ct = default);
}
