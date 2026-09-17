namespace Gastos.Business.Proveedores;

/// <summary>
/// Configuración de la API de un proveedor de tarjetas (secciones Proveedores:Repsol y Proveedores:Moeve).
/// PENDIENTE: ajustar cuando el proveedor facilite la documentación de su API (URL, autenticación, formato).
/// </summary>
public sealed class ProveedorApiOptions
{
    public string? BaseUrl { get; set; }
    /// <summary>Ruta relativa del recurso de movimientos. Admite {desde} y {hasta} (yyyy-MM-dd).</summary>
    public string RutaMovimientos { get; set; } = "movimientos?desde={desde}&hasta={hasta}";
    /// <summary>Token/API key que se envía como Bearer. NO se commitea: user-secrets o variable de entorno.</summary>
    public string? ApiKey { get; set; }
    /// <summary>Nombre de la cabecera si el proveedor no usa "Authorization: Bearer" (p. ej. "X-Api-Key").</summary>
    public string? CabeceraApiKey { get; set; }
    public int TimeoutSegundos { get; set; } = 120;
    /// <summary>Modo simulado (desarrollo): en lugar de llamar a la API lee un CSV con el formato antiguo de Solred.</summary>
    public bool Simulado { get; set; }
    public string? FicheroSimulacion { get; set; }
}
