using Gastos.Business.Cargas;
using Xunit;

namespace Gastos.Tests;

public class SolredCsvParserTests
{
    private const string LineaOk = "2026;F-77;0012345;1234ABC;PEREZ JUAN;15/01/2026;08:35;BADAJOZ;12345;DIESEL E+;45.2;21;60.50;1.339;ES-SUR;extra";

    [Fact]
    public void Linea_valida_se_convierte_en_movimiento()
    {
        var r = SolredCsvParser.Parsear(new StringReader(LineaOk));
        Assert.True(r.Ok, r.Error);
        var m = Assert.Single(r.Movimientos);
        Assert.Equal(2026, m.AnyoFactura);
        Assert.Equal("F-77", m.Factura);
        Assert.Equal("12345", m.NumTarjeta);            // ceros a la izquierda eliminados
        Assert.Equal(new DateTime(2026, 1, 15), m.FechaOperacion);
        Assert.Equal("08:35", m.HoraOperacion);
        Assert.Equal(45.2m, m.Litros);
        Assert.Equal(60.50m, m.ImporteTotal);
        Assert.Equal(1.339m, m.PrecioLitro);
        Assert.Equal("ES-SUR;extra", m.InfoAux);        // el resto de la línea, con sus ';'
    }

    [Theory]
    [InlineData("abcd;F;1;M;C;15/01/2026;08:35;B;1;D;1;21;10;1;x", "Año Factura")]
    [InlineData("2026;F;1;M;C;15012026;08:35;B;1;D;1;21;10;1;x", "fecha")]
    [InlineData("2026;F;1;M;C;15/01/2026;0835;B;1;D;1;21;10;1;x", "hora")]
    [InlineData("2026;F;1;M;C;15/01/2026;08:35;B;km;D;1;21;10;1;x", "Kilómetros")]
    [InlineData("2026;F;1;M;C;15/01/2026;08:35;B;1;D;1;21;total;1;x", "Importe Total")]
    public void Validaciones(string linea, string texto)
    {
        var r = SolredCsvParser.Parsear(new StringReader(linea));
        Assert.False(r.Ok);
        Assert.Contains(texto, r.Error);
    }
}
