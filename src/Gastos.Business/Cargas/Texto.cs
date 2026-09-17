using System.Globalization;

namespace Gastos.Business.Cargas;

/// <summary>Utilidades de texto/número heredadas del VB6 (LimpiarCadena, IsNumeric, truncado de "Donde").</summary>
public static class Texto
{
    private static readonly CultureInfo Es = CultureInfo.GetCultureInfo("es-ES");

    /// <summary>Port de LimpiarCadena: apóstrofos y comillas dobles pasan a acento grave, los ';' a coma, y se recorta.</summary>
    public static string Limpiar(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        return s.Replace('\'', '`').Replace('"', '`').Replace(';', ',').Trim();
    }

    /// <summary>"Donde" en el VB6: primeros 50 caracteres y los apóstrofos sustituidos por espacio.</summary>
    public static string Donde(string? s)
    {
        s ??= string.Empty;
        if (s.Length > 50) s = s[..50];
        return s.Replace('\'', ' ');
    }

    /// <summary>
    /// Equivalente práctico a IsNumeric + conversión: admite "1234.56" (formato del CSV) y "1234,56" (formato es-ES).
    /// Si hay coma se interpreta como decimal español (los puntos serían de miles); si no, formato invariante.
    /// </summary>
    public static bool TryNumero(string? s, out decimal valor)
    {
        valor = 0;
        if (string.IsNullOrWhiteSpace(s)) return false;
        s = s.Trim();
        if (s.Contains(','))
            return decimal.TryParse(s, NumberStyles.Number, Es, out valor);
        return decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out valor);
    }

    public static bool TryEntero(string? s, out int valor)
    {
        valor = 0;
        if (!TryNumero(s, out var d) || d != Math.Truncate(d) || d < int.MinValue || d > int.MaxValue) return false;
        valor = (int)d;
        return true;
    }
}
