using Gastos.Domain.Entities;
using Gastos.Domain.Models;

namespace Gastos.Domain.Abstractions.Services;

/// <summary>Cargas de gastos: fichero LayTours y APIs de Repsol/Moeve. Solo para Administradores (lo controla la UI).</summary>
public interface ICargaService
{
    /// <summary>Carga el CSV de LayTours (formato del antiguo GeneracionCSV\LayTours.csv) para la sociedad indicada.</summary>
    Task<ResultadoCarga> CargarLayToursAsync(Stream csv, int codPais, string usuario, string nombreFichero, CancellationToken ct = default);

    /// <summary>Pide a la API del proveedor (Repsol o Moeve) los movimientos entre fechas y los carga.</summary>
    Task<ResultadoCarga> CargarProveedorAsync(OrigenCarga origen, DateTime desde, DateTime hasta, string usuario, CancellationToken ct = default);

    /// <summary>Carga movimientos ya obtenidos (p. ej. de un CSV de Solred antiguo o de pruebas) como si vinieran del proveedor.</summary>
    Task<ResultadoCarga> CargarMovimientosAsync(OrigenCarga origen, IReadOnlyList<MovimientoTarjeta> movimientos, string usuario, string? nombreFichero, DateTime? desde, DateTime? hasta, CancellationToken ct = default);

    Task<IReadOnlyList<Carga>> DameCargasAsync(int ultimas = 200, CancellationToken ct = default);

    /// <summary>Comprobación previa del CSV de LayTours: formato y usuarios Epsilon sin asignar en la sociedad. No guarda nada.</summary>
    Task<ResultadoComprobacion> ComprobarLayToursAsync(Stream csv, int codPais, CancellationToken ct = default);

    /// <summary>Comprobación previa de una carga por API: pide los movimientos y lista las tarjetas que no están dadas de alta. No guarda nada.</summary>
    Task<ResultadoComprobacion> ComprobarProveedorAsync(OrigenCarga origen, DateTime desde, DateTime hasta, CancellationToken ct = default);

    /// <summary>Proveedores de API registrados y si están configurados (para habilitar/deshabilitar opciones en la UI).</summary>
    IReadOnlyList<(OrigenCarga Origen, bool Configurado, string Detalle)> EstadoProveedores();
}
