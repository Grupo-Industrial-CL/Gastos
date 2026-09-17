using Gastos.Domain.Models;

namespace Gastos.Business.Cargas;

/// <summary>Resultado de clasificar un movimiento de tarjeta como gasto interno.</summary>
public sealed record GastoClasificado(string TipoGasto, string Donde, int Numero, decimal ImporteSinIVA);

/// <summary>
/// Port de la clasificación de movimientos de tarjeta de carburante del VB6 (frmCargaFicheroSolred2, rama "Interno"):
/// por defecto Gasolina; según la descripción del producto pasa a Parking, Peaje u Otros gastos (lavados).
/// Los códigos de tipo de gasto son los del VB6: 1 Gasolina, 9 Peaje, 10 Otros gastos, 11 Parking.
/// </summary>
public static class ClasificadorCarburante
{
    public const string Gasolina = "1";
    public const string Peaje = "9";
    public const string OtrosGastos = "10";
    public const string Parking = "11";

    public static GastoClasificado Clasificar(MovimientoTarjeta m)
    {
        var desc = (m.DescripcionProducto ?? string.Empty).ToLowerInvariant();
        var poblacion = m.Poblacion ?? string.Empty;
        var tipo = Gasolina;
        var numero = 0;
        var donde = poblacion;

        if (desc.Contains("aparcamiento") || desc.Contains("aparcamien"))
        {
            tipo = Parking;
            donde = $"Parking {poblacion} ({m.InfoAux})";
            numero = 0;
        }
        if (desc.Contains("estradas") || desc.Contains("autopista") || desc.Contains("teroperable") || desc.Contains("autoestrad"))
        {
            tipo = Peaje;
            donde = $"{poblacion} ({m.InfoAux})";
            numero = 1;
        }
        if (desc.Contains("lavados"))
        {
            tipo = OtrosGastos;
            numero = 1;
        }
        return new GastoClasificado(tipo, Texto.Donde(donde), numero, ImporteSinIVA(m.ImporteTotal, m.IVA));
    }

    /// <summary>Round(Importe / (1 + IVA/100), 2) como en el VB6 (redondeo bancario de VB6 = MidpointRounding.ToEven).</summary>
    public static decimal ImporteSinIVA(decimal importeTotal, decimal iva)
        => Math.Round(importeTotal / (1 + iva / 100m), 2, MidpointRounding.ToEven);

    /// <summary>
    /// Normaliza el número de tarjeta como el VB6: quita ceros a la izquierda y decimales ("1234.0"), y las tarjetas
    /// que empiezan por 9000000 se agrupan en la genérica 9000000000000000.
    /// </summary>
    public static string NormalizarTarjeta(string? tarjeta)
    {
        var t = (tarjeta ?? string.Empty).Trim();
        t = t.TrimStart('0');
        var punto = t.IndexOf('.'); if (punto >= 0) t = t[..punto];
        var coma = t.IndexOf(','); if (coma >= 0) t = t[..coma];
        if (t.StartsWith("9000000", StringComparison.Ordinal)) t = "9000000000000000";
        return t;
    }
}
