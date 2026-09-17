using System.Globalization;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Serialization;
using Gastos.Business.Cargas;
using Gastos.Domain;
using Gastos.Domain.Abstractions;
using Gastos.Domain.Models;
using Microsoft.Extensions.Logging;

namespace Gastos.Business.Proveedores;

/// <summary>
/// Cliente genérico de API de proveedor de tarjetas de carburante. Hace GET a la ruta de movimientos con autenticación
/// por API key y traduce el JSON al formato neutro <see cref="MovimientoTarjeta"/>. Las clases hijas (Repsol/Moeve)
/// solo fijan el origen; cuando llegue la documentación real de cada API, se sobreescribe <see cref="TraducirAsync"/>.
/// </summary>
public abstract class ProveedorApiBase : IProveedorTarjetaCarburante
{
    private readonly HttpClient _http;
    private readonly ProveedorApiOptions _opt;
    private readonly ILogger _log;

    protected ProveedorApiBase(HttpClient http, ProveedorApiOptions opt, ILogger log)
    {
        _http = http;
        _opt = opt;
        _log = log;
    }

    public abstract OrigenCarga Origen { get; }

    public bool Configurado => _opt.Simulado
        ? !string.IsNullOrWhiteSpace(_opt.FicheroSimulacion)
        : !string.IsNullOrWhiteSpace(_opt.BaseUrl) && !string.IsNullOrWhiteSpace(_opt.ApiKey);

    public string Detalle => _opt.Simulado
        ? (Configurado ? $"Simulado (fichero {Path.GetFileName(_opt.FicheroSimulacion)})" : "Simulado sin fichero de simulación")
        : (Configurado ? "Configurado" : "Pendiente de configurar (URL y credenciales de la API)");

    public async Task<IReadOnlyList<MovimientoTarjeta>> ObtenerMovimientosAsync(DateTime desde, DateTime hasta, CancellationToken ct = default)
    {
        if (_opt.Simulado)
            return Simular(desde, hasta);

        if (!Configurado) throw new InvalidOperationException($"API de {Origen.Codigo()} no configurada.");

        var ruta = _opt.RutaMovimientos
            .Replace("{desde}", desde.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
            .Replace("{hasta}", hasta.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        var url = new Uri(new Uri(_opt.BaseUrl!.TrimEnd('/') + "/"), ruta);

        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        if (string.IsNullOrWhiteSpace(_opt.CabeceraApiKey))
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _opt.ApiKey);
        else
            req.Headers.TryAddWithoutValidation(_opt.CabeceraApiKey, _opt.ApiKey);
        req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        _log.LogInformation("Llamando a la API de {Origen}: {Url}", Origen, url);
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        cts.CancelAfter(TimeSpan.FromSeconds(Math.Max(10, _opt.TimeoutSegundos)));
        using var resp = await _http.SendAsync(req, cts.Token);
        resp.EnsureSuccessStatusCode();
        return await TraducirAsync(resp, cts.Token);
    }

    /// <summary>
    /// Traducción por defecto: la API devuelve un array JSON de movimientos con estos nombres (o un objeto con la propiedad
    /// "movimientos"/"data"/"items" que los contiene). PENDIENTE: adaptar al contrato real de cada proveedor.
    /// </summary>
    protected virtual async Task<IReadOnlyList<MovimientoTarjeta>> TraducirAsync(HttpResponseMessage resp, CancellationToken ct)
    {
        var opciones = new JsonSerializerOptions(JsonSerializerDefaults.Web) { NumberHandling = JsonNumberHandling.AllowReadingFromString };
        using var doc = await resp.Content.ReadFromJsonAsync<JsonDocument>(opciones, ct) ?? throw new InvalidOperationException("Respuesta vacía.");
        var raiz = doc.RootElement;
        if (raiz.ValueKind == JsonValueKind.Object)
        {
            foreach (var nombre in new[] { "movimientos", "data", "items", "results", "transactions" })
                if (raiz.TryGetProperty(nombre, out var arr) && arr.ValueKind == JsonValueKind.Array) { raiz = arr; break; }
        }
        if (raiz.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("La respuesta de la API no contiene un array de movimientos.");

        var lista = new List<MovimientoTarjeta>();
        foreach (var e in raiz.EnumerateArray())
            lista.Add(e.Deserialize<MovimientoApiDto>(opciones)!.AMovimiento());
        return lista;
    }

    private IReadOnlyList<MovimientoTarjeta> Simular(DateTime desde, DateTime hasta)
    {
        if (!File.Exists(_opt.FicheroSimulacion)) throw new FileNotFoundException($"No existe el fichero de simulación '{_opt.FicheroSimulacion}'.");
        using var lector = new StreamReader(_opt.FicheroSimulacion!, System.Text.Encoding.Latin1, true);
        var parseo = SolredCsvParser.Parsear(lector);
        if (!parseo.Ok) throw new InvalidOperationException($"Fichero de simulación incorrecto: {parseo.Error}");
        return parseo.Movimientos.Where(m => m.FechaOperacion.Date >= desde.Date && m.FechaOperacion.Date <= hasta.Date).ToList();
    }

    /// <summary>Forma esperada (provisional) de un movimiento en el JSON de la API. Se acepta también en inglés.</summary>
    protected sealed class MovimientoApiDto
    {
        public string? NumTarjeta { get; set; }
        public string? Tarjeta { get; set; }
        public string? CardNumber { get; set; }
        public int? AnyoFactura { get; set; }
        public string? Factura { get; set; }
        public string? InvoiceNumber { get; set; }
        public string? Matricula { get; set; }
        public string? Conductor { get; set; }
        public DateTime? FechaOperacion { get; set; }
        public DateTime? Fecha { get; set; }
        public DateTime? TransactionDate { get; set; }
        public string? HoraOperacion { get; set; }
        public string? Hora { get; set; }
        public string? Poblacion { get; set; }
        public decimal? Km { get; set; }
        public string? DescripcionProducto { get; set; }
        public string? Producto { get; set; }
        public decimal? Litros { get; set; }
        public decimal? IVA { get; set; }
        public decimal? ImporteTotal { get; set; }
        public decimal? Importe { get; set; }
        public decimal? PrecioLitro { get; set; }
        public string? InfoAux { get; set; }

        public MovimientoTarjeta AMovimiento()
        {
            var fecha = FechaOperacion ?? Fecha ?? TransactionDate ?? throw new InvalidOperationException("Movimiento sin fecha.");
            var hora = HoraOperacion ?? Hora ?? (fecha.TimeOfDay > TimeSpan.Zero ? fecha.ToString("HH:mm") : "00:00");
            var factura = Factura ?? InvoiceNumber ?? string.Empty;
            return new MovimientoTarjeta
            {
                NumTarjeta = ClasificadorCarburante.NormalizarTarjeta(NumTarjeta ?? Tarjeta ?? CardNumber),
                AnyoFactura = AnyoFactura ?? fecha.Year,
                Factura = factura,
                Matricula = Matricula,
                Conductor = Conductor,
                FechaOperacion = fecha.Date,
                HoraOperacion = hora,
                Poblacion = Poblacion,
                Km = Km ?? 0,
                DescripcionProducto = DescripcionProducto ?? Producto,
                Litros = Litros ?? 0,
                IVA = IVA ?? 21,
                ImporteTotal = ImporteTotal ?? Importe ?? 0,
                PrecioLitro = PrecioLitro ?? 0,
                InfoAux = InfoAux
            };
        }
    }
}
