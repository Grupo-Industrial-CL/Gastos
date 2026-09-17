using Gastos.Domain;
using Gastos.Domain.Abstractions.Services;
using Microsoft.Extensions.Options;

namespace Gastos.Web.Services;

/// <summary>Configuración de la carga programada (sección CargasProgramadas).</summary>
public sealed class CargasProgramadasOptions
{
    public bool Habilitado { get; set; }
    /// <summary>Hora local de ejecución diaria, formato HH:mm.</summary>
    public string Hora { get; set; } = "03:00";
    /// <summary>Orígenes a cargar: Repsol y/o Moeve. Los no configurados se saltan con aviso.</summary>
    public List<string> Origenes { get; set; } = new() { "Repsol", "Moeve" };
    /// <summary>Días hacia atrás que se piden a la API en cada ejecución (los movimientos ya cargados se omiten).</summary>
    public int DiasAtras { get; set; } = 7;
    public const string Usuario = "Carga programada";
}

/// <summary>
/// Carga programada de movimientos de tarjeta por API: una vez al día, a la hora configurada, pide a Repsol/Moeve los
/// movimientos de los últimos N días y los carga con el mismo servicio que el botón manual (idempotente: lo ya cargado se
/// omite). Queda registrada en GT_Cargas con usuario "Carga programada". Deshabilitada por defecto.
/// </summary>
public sealed class CargaProgramadaBackgroundService : BackgroundService
{
    private readonly IServiceScopeFactory _scopes;
    private readonly IOptionsMonitor<CargasProgramadasOptions> _opt;
    private readonly ILogger<CargaProgramadaBackgroundService> _log;

    public CargaProgramadaBackgroundService(IServiceScopeFactory scopes, IOptionsMonitor<CargasProgramadasOptions> opt, ILogger<CargaProgramadaBackgroundService> log)
    {
        _scopes = scopes;
        _opt = opt;
        _log = log;
    }

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            var o = _opt.CurrentValue;
            if (!o.Habilitado)
            {
                // Se re-evalúa cada hora por si se habilita en caliente.
                await Task.Delay(TimeSpan.FromHours(1), ct);
                continue;
            }

            var espera = HastaProximaEjecucion(o.Hora);
            _log.LogInformation("Carga programada: próxima ejecución en {Espera} (a las {Hora})", espera, o.Hora);
            try { await Task.Delay(espera, ct); } catch (OperationCanceledException) { break; }
            if (ct.IsCancellationRequested) break;

            await EjecutarAsync(_opt.CurrentValue, ct);
        }
    }

    private async Task EjecutarAsync(CargasProgramadasOptions o, CancellationToken ct)
    {
        var hasta = DateTime.Today;
        var desde = hasta.AddDays(-Math.Max(1, o.DiasAtras));
        foreach (var nombre in o.Origenes)
        {
            if (!Enum.TryParse<OrigenCarga>(nombre, true, out var origen) || origen == OrigenCarga.LayTours)
            {
                _log.LogWarning("Carga programada: origen '{Origen}' no válido (Repsol o Moeve)", nombre);
                continue;
            }
            try
            {
                using var scope = _scopes.CreateScope();
                var cargas = scope.ServiceProvider.GetRequiredService<ICargaService>();
                if (!cargas.EstadoProveedores().Any(p => p.Origen == origen && p.Configurado))
                {
                    _log.LogWarning("Carga programada: la API de {Origen} no está configurada; se omite", origen);
                    continue;
                }
                var r = await cargas.CargarProveedorAsync(origen, desde, hasta, CargasProgramadasOptions.Usuario, ct);
                _log.LogInformation("Carga programada {Origen} ({Desde:d} - {Hasta:d}): {Estado}. {Mensaje}", origen, desde, hasta, r.Ok ? "OK" : "ERROR", r.Mensaje);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                _log.LogError(ex, "Carga programada {Origen}: error inesperado", origen);
            }
        }
    }

    private static TimeSpan HastaProximaEjecucion(string hora)
    {
        if (!TimeSpan.TryParse(hora, out var h)) h = new TimeSpan(3, 0, 0);
        var ahora = DateTime.Now;
        var proxima = ahora.Date + h;
        if (proxima <= ahora) proxima = proxima.AddDays(1);
        return proxima - ahora;
    }
}
