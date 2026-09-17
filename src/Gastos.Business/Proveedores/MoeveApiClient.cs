using Gastos.Domain;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace Gastos.Business.Proveedores;

/// <summary>API de Moeve (tarjetas de carburante). Configuración en Proveedores:Moeve. PENDIENTE: contrato real de la API.</summary>
public sealed class MoeveApiClient : ProveedorApiBase
{
    public MoeveApiClient(HttpClient http, IOptionsMonitor<ProveedorApiOptions> opts, ILogger<MoeveApiClient> log)
        : base(http, opts.Get(Nombre), log) { }

    public const string Nombre = "Moeve";
    public override OrigenCarga Origen => OrigenCarga.Moeve;
}
