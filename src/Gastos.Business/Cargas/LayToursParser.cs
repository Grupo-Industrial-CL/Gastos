using System.Globalization;

namespace Gastos.Business.Cargas;

/// <summary>Línea válida del CSV de LayTours (9 campos separados por ';').</summary>
public sealed class LineaLayTours
{
    public int Fila { get; init; }
    public string Factura { get; init; } = string.Empty;          // "Invoice number"
    public string Concepto { get; init; } = string.Empty;         // "Accounting concept" (normalizado en minúsculas)
    public decimal Importe { get; init; }                         // "Amount"
    public string UsuarioEpsilon { get; init; } = string.Empty;   // "User ID" (numérico)
    public string NumeroAlbaran { get; init; } = string.Empty;
    public string Detalle { get; init; } = string.Empty;
    public DateTime Fecha { get; init; }                          // "Fecha inicio servicio" (ddMMyyyy)
    public int NumeroDias { get; init; }                          // "Num de dias"
    public string Destino { get; init; } = string.Empty;

    /// <summary>Familia del concepto según el VB6: Billetes / Hoteles / Otros gastos.</summary>
    public string NombreGasto => LayToursParser.NombreGastoDe(Concepto);
    public bool EsHotel => Concepto == "hotel";
    /// <summary>"Donde" del gasto: albarán/destino para billetes; albarán/detalle para el resto.</summary>
    public string Donde => Texto.Donde(NombreGasto == "Billetes" ? $"{NumeroAlbaran}/{Destino}" : $"{NumeroAlbaran}/{Detalle}");
}

public sealed class ResultadoParseoLayTours
{
    public List<LineaLayTours> Lineas { get; } = new();
    public string? Error { get; set; }
    public bool Ok => Error is null;
}

/// <summary>
/// Port fiel de la lectura del fichero GeneracionCSV\LayTours.csv del VB6 (frmCargarFicheroORBIS.cmdCargar_Click).
/// Es "todo o nada": el primer error detiene el proceso con el número de fila (igual que el VB6, que hacía rollback).
/// </summary>
public static class LayToursParser
{
    public static readonly string[] ConceptosValidos =
    {
        "avion", "gasto gestion", "hotel", "mayoristas internac.", "mayorista internac.", "mayoristas nacional",
        "mayorista nacional", "no usar  avion / tren / barco", "otros", "rent a car", "servicios de tierra",
        "servicio de tierra", "tren"
    };

    public static string NombreGastoDe(string concepto) => concepto switch
    {
        "avion" or "no usar  avion / tren / barco" or "tren" => "Billetes",
        "hotel" => "Hoteles",
        _ => "Otros gastos"
    };

    public static ResultadoParseoLayTours Parsear(TextReader lector)
    {
        var r = new ResultadoParseoLayTours();
        var fila = 0;
        string? linea;
        while ((linea = lector.ReadLine()) is not null)
        {
            fila++;
            linea = linea.Trim();
            if (linea.Length == 0)
            {
                // El VB6 abortaba en una línea vacía. Se tolera únicamente una última línea vacía (salto final del fichero).
                if (lector.Peek() == -1) break;
                return Fallo(r, $"Línea vacía en la fila {fila}.");
            }

            var campos = linea.Split(';');
            if (campos.Length != 9) return Fallo(r, $"Número de columnas incorrectas en la fila {fila} (se esperaban 9 y hay {campos.Length}).");

            var factura = campos[0].Trim();
            if (factura.Length == 0) return Fallo(r, $"Campo 'Invoice number' incorrecto en la fila {fila}.");

            var concepto = campos[1].Trim().ToLowerInvariant();
            if (!ConceptosValidos.Contains(concepto)) return Fallo(r, $"Campo 'Accounting concept' incorrecto en la fila {fila} ('{campos[1].Trim()}').");

            if (!Texto.TryNumero(campos[2], out var importe)) return Fallo(r, $"Campo 'Amount' incorrecto en la fila {fila}.");

            var usuario = campos[3].Trim();
            if (!Texto.TryNumero(usuario, out _)) return Fallo(r, $"Campo 'User ID' incorrecto en la fila {fila}.");

            var albaran = Texto.Limpiar(campos[4]);
            var detalle = Texto.Limpiar(campos[5]);

            if (!TryFecha(campos[6].Trim(), out var fecha)) return Fallo(r, $"Campo 'Fecha inicio servicio' incorrecto en la fila {fila}.");

            if (!Texto.TryEntero(campos[7], out var dias)) return Fallo(r, $"Campo 'Num de dias' incorrecto en la fila {fila}.");

            var destino = Texto.Limpiar(campos[8]);

            r.Lineas.Add(new LineaLayTours
            {
                Fila = fila, Factura = factura, Concepto = concepto, Importe = importe, UsuarioEpsilon = usuario,
                NumeroAlbaran = albaran, Detalle = detalle, Fecha = fecha, NumeroDias = dias, Destino = destino
            });
        }
        if (r.Lineas.Count == 0 && r.Ok) r.Error = "El fichero está vacío.";
        return r;
    }

    // Fecha numérica ddMMyyyy (7 u 8 dígitos; con 7 se antepone un 0), validando día/mes/bisiesto como el VB6.
    private static bool TryFecha(string s, out DateTime fecha)
    {
        fecha = default;
        if (s.Length == 7 && s.All(char.IsDigit)) s = "0" + s;
        if (s.Length != 8 || !s.All(char.IsDigit)) return false;
        return DateTime.TryParseExact(s, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha);
    }

    private static ResultadoParseoLayTours Fallo(ResultadoParseoLayTours r, string error) { r.Error = error; return r; }
}
