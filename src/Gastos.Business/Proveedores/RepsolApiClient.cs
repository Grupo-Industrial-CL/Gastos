using Gastos.Domain;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace Gastos.Business.Proveedores;

/// <summary>API de Repsol (tarjetas Solred). Configuración en Proveedores:Repsol. PENDIENTE: contrato real de la API.</summary>
public sealed class RepsolApiClient : ProveedorApiBase
{
    public RepsolApiClient(HttpClient http, IOptionsMonitor<ProveedorApiOptions> opts, ILogger<RepsolApiClient> log)
        : base(http, opts.Get(Nombre), log) { }

    public const string Nombre = "Repsol";
    public override OrigenCarga Origen => OrigenCarga.Repsol;
}
