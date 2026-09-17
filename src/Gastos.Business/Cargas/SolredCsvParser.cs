using System.Globalization;
using Gastos.Domain.Models;

namespace Gastos.Business.Cargas;

public sealed class ResultadoParseoSolred
{
    public List<MovimientoTarjeta> Movimientos { get; } = new();
    public string? Error { get; set; }
    public bool Ok => Error is null;
}

/// <summary>
/// Port del lector del antiguo GeneracionCSV\Solred.csv (frmCargaFicheroSolred2). Se conserva como vía transitoria
/// hasta que la API de Repsol esté operativa y para poder cargar históricos. Campos (separados por ';'):
/// Año factura; Factura; Tarjeta; Matrícula; Conductor; Fecha (dd/MM/yyyy); Hora (HH:mm); Población; Km;
/// Descripción producto; Litros; IVA; Importe total; Precio litro; Info auxiliar (resto de la línea).
/// </summary>
public static class SolredCsvParser
{
    public static ResultadoParseoSolred Parsear(TextReader lector)
    {
        var r = new ResultadoParseoSolred();
        var fila = 0;
        string? linea;
        while ((linea = lector.ReadLine()) is not null)
        {
            fila++;
            linea = linea.Trim();
            if (linea.Length == 0)
            {
                if (lector.Peek() == -1) break;
                return Fallo(r, $"Línea vacía en la fila {fila}.");
            }
            var c = linea.Split(';');
            if (c.Length < 14) return Fallo(r, $"Número de columnas incorrectas en la fila {fila} (se esperaban al menos 14 y hay {c.Length}).");

            if (!Texto.TryEntero(c[0], out var anyo)) return Fallo(r, $"Campo 'Año Factura' incorrecto en la fila {fila}.");
            var factura = c[1].Trim();
            if (factura.Length == 0) return Fallo(r, $"Campo 'Factura' incorrecto en la fila {fila}.");
            var tarjeta = ClasificadorCarburante.NormalizarTarjeta(c[2]);
            if (tarjeta.Length == 0) return Fallo(r, $"Campo 'Tarjeta' incorrecto en la fila {fila}.");
            var matricula = c[3].Trim();
            var conductor = c[4].Trim();
            var sFecha = c[5].Trim();
            if (!sFecha.Contains('/') || !DateTime.TryParseExact(sFecha, new[] { "dd/MM/yyyy", "d/M/yyyy", "dd/MM/yy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out var fecha))
                return Fallo(r, $"Campo fecha incorrecto en la fila {fila}.");
            var hora = c[6].Trim();
            if (!hora.Contains(':')) return Fallo(r, $"Campo hora incorrecto en la fila {fila}.");
            var poblacion = c[7].Trim();
            if (!Texto.TryNumero(c[8], out var km)) return Fallo(r, $"Campo 'Kilómetros' incorrecto en la fila {fila}.");
            var descProd = c[9].Trim();
            if (!Texto.TryNumero(c[10], out var litros)) return Fallo(r, $"Campo 'Nº Litros' incorrecto en la fila {fila}.");
            if (!Texto.TryNumero(c[11], out var iva)) return Fallo(r, $"Campo 'IVA' incorrecto en la fila {fila}.");
            if (!Texto.TryNumero(c[12], out var total)) return Fallo(r, $"Campo 'Importe Total' incorrecto en la fila {fila}.");
            if (!Texto.TryNumero(c[13], out var precioLitro)) return Fallo(r, $"Campo 'Precio litro' incorrecto en la fila {fila}.");
            var infoAux = c.Length > 14 ? string.Join(";", c.Skip(14)).Trim() : string.Empty;

            r.Movimientos.Add(new MovimientoTarjeta
            {
                NumTarjeta = tarjeta, AnyoFactura = anyo, Factura = factura, Matricula = matricula, Conductor = conductor,
                FechaOperacion = fecha, HoraOperacion = hora, Poblacion = poblacion, Km = km, DescripcionProducto = descProd,
                Litros = litros, IVA = iva, ImporteTotal = total, PrecioLitro = precioLitro, InfoAux = infoAux
            });
        }
        if (r.Movimientos.Count == 0 && r.Ok) r.Error = "El fichero está vacío.";
        return r;
    }

    private static ResultadoParseoSolred Fallo(ResultadoParseoSolred r, string error) { r.Error = error; return r; }
}
