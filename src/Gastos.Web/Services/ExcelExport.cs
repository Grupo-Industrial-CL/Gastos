using System.Globalization;
using System.IO.Compression;
using System.Text;

namespace Gastos.Web.Services;

/// <summary>Celda de exportación: texto o número (para poder sumar en Excel).</summary>
public readonly struct ExcelCell
{
    public string? Text { get; }
    public double? Number { get; }
    private ExcelCell(string? t, double? n) { Text = t; Number = n; }
    public static ExcelCell Txt(string? t) => new(t ?? string.Empty, null);
    public static ExcelCell Num(double n) => new(null, n);
}

/// <summary>
/// Generador de .xlsx (OOXML) sin dependencias externas: escribe el paquete zip con las partes mínimas.
/// Exporta lo que el usuario ve en el grid (valores ya formateados; números como número). Tomado de MasterPlan v2.
/// </summary>
public static class ExcelExport
{
    private static readonly UTF8Encoding Utf8 = new(false);

    public static byte[] Crear(IEnumerable<string> headers, IEnumerable<IEnumerable<ExcelCell>> rows, string sheetName = "Datos")
    {
        var hs = headers.ToList();
        var data = rows.Select(r => r.ToList()).ToList();
        var hoja = SanearNombreHoja(sheetName);

        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            Escribir(zip, "[Content_Types].xml", ContentTypes());
            Escribir(zip, "_rels/.rels", RootRels());
            Escribir(zip, "xl/workbook.xml", Workbook(hoja));
            Escribir(zip, "xl/_rels/workbook.xml.rels", WorkbookRels());
            Escribir(zip, "xl/worksheets/sheet1.xml", Sheet(hs, data));
        }
        return ms.ToArray();
    }

    private static void Escribir(ZipArchive zip, string ruta, string contenido)
    {
        var e = zip.CreateEntry(ruta, CompressionLevel.Optimal);
        using var s = e.Open();
        var bytes = Utf8.GetBytes(contenido);
        s.Write(bytes, 0, bytes.Length);
    }

    private static string Sheet(List<string> headers, List<List<ExcelCell>> data)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
        var fila = 1;
        if (headers.Count > 0) Fila(sb, fila++, headers.Select(ExcelCell.Txt).ToList());
        foreach (var r in data) Fila(sb, fila++, r);
        sb.Append("</sheetData></worksheet>");
        return sb.ToString();
    }

    private static void Fila(StringBuilder sb, int fila, List<ExcelCell> celdas)
    {
        sb.Append("<row r=\"").Append(fila).Append("\">");
        for (var i = 0; i < celdas.Count; i++)
        {
            var refc = ColLetra(i + 1) + fila;
            var c = celdas[i];
            if (c.Number.HasValue)
                sb.Append("<c r=\"").Append(refc).Append("\"><v>").Append(c.Number.Value.ToString(CultureInfo.InvariantCulture)).Append("</v></c>");
            else
                sb.Append("<c r=\"").Append(refc).Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">").Append(EscaparXml(c.Text)).Append("</t></is></c>");
        }
        sb.Append("</row>");
    }

    private static string ColLetra(int idx)
    {
        var s = string.Empty;
        while (idx > 0) { var r = (idx - 1) % 26; s = (char)('A' + r) + s; idx = (idx - 1) / 26; }
        return s;
    }

    private static string EscaparXml(string? v)
    {
        if (string.IsNullOrEmpty(v)) return string.Empty;
        var sb = new StringBuilder(v.Length);
        foreach (var ch in v)
        {
            switch (ch)
            {
                case '&': sb.Append("&amp;"); break;
                case '<': sb.Append("&lt;"); break;
                case '>': sb.Append("&gt;"); break;
                case '"': sb.Append("&quot;"); break;
                case '\'': sb.Append("&apos;"); break;
                default: if (ch == '\t' || ch == '\n' || ch == '\r' || ch >= 0x20) sb.Append(ch); break;
            }
        }
        return sb.ToString();
    }

    private static string SanearNombreHoja(string? nombre)
    {
        var n = (nombre ?? "Datos").Trim();
        foreach (var c in new[] { ':', '\\', '/', '?', '*', '[', ']' }) n = n.Replace(c, ' ');
        if (n.Length == 0) n = "Datos";
        return n.Length > 31 ? n[..31] : n;
    }

    private static string ContentTypes() =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
        "</Types>";

    private static string RootRels() =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
        "</Relationships>";

    private static string Workbook(string hoja) =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
        "<sheets><sheet name=\"" + EscaparXml(hoja) + "\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";

    private static string WorkbookRels() =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
        "</Relationships>";
}
