using Gastos.Business.Cargas;
using Gastos.Domain.Models;
using Xunit;

namespace Gastos.Tests;

public class ClasificadorCarburanteTests
{
    private static MovimientoTarjeta Mov(string producto, string poblacion = "Badajoz", string infoAux = "A-5", decimal total = 121m, decimal iva = 21m)
        => new() { DescripcionProducto = producto, Poblacion = poblacion, InfoAux = infoAux, ImporteTotal = total, IVA = iva };

    [Fact]
    public void Por_defecto_es_gasolina_con_donde_poblacion()
    {
        var g = ClasificadorCarburante.Clasificar(Mov("Diesel e+"));
        Assert.Equal(ClasificadorCarburante.Gasolina, g.TipoGasto);
        Assert.Equal("Badajoz", g.Donde);
        Assert.Equal(0, g.Numero);
        Assert.Equal(100m, g.ImporteSinIVA);
    }

    [Fact]
    public void Aparcamiento_es_parking()
    {
        var g = ClasificadorCarburante.Clasificar(Mov("APARCAMIENTO SABA"));
        Assert.Equal(ClasificadorCarburante.Parking, g.TipoGasto);
        Assert.Equal("Parking Badajoz (A-5)", g.Donde);
    }

    [Theory]
    [InlineData("Autopista AP-7")]
    [InlineData("Peaje interoperable")]
    [InlineData("Autoestradas Portugal")]
    public void Autopistas_son_peaje_con_numero_1(string producto)
    {
        var g = ClasificadorCarburante.Clasificar(Mov(producto));
        Assert.Equal(ClasificadorCarburante.Peaje, g.TipoGasto);
        Assert.Equal("Badajoz (A-5)", g.Donde);
        Assert.Equal(1, g.Numero);
    }

    [Fact]
    public void Lavados_son_otros_gastos()
    {
        var g = ClasificadorCarburante.Clasificar(Mov("Lavados premium"));
        Assert.Equal(ClasificadorCarburante.OtrosGastos, g.TipoGasto);
        Assert.Equal(1, g.Numero);
    }

    [Fact]
    public void Donde_se_trunca_a_50_y_quita_apostrofos()
    {
        var g = ClasificadorCarburante.Clasificar(Mov("Diesel", new string('x', 60) + "'"));
        Assert.Equal(50, g.Donde.Length);
        Assert.DoesNotContain("'", g.Donde);
    }

    [Theory]
    [InlineData("0001234567", "1234567")]
    [InlineData("1234567.0", "1234567")]
    [InlineData("1234567,00", "1234567")]
    [InlineData("9000000123456789", "9000000000000000")]
    [InlineData("  777 ", "777")]
    public void Normalizar_tarjeta_como_el_vb6(string entrada, string esperada)
        => Assert.Equal(esperada, ClasificadorCarburante.NormalizarTarjeta(entrada));

    [Fact]
    public void Importe_sin_iva_redondea_a_dos_decimales()
        => Assert.Equal(41.32m, ClasificadorCarburante.ImporteSinIVA(50m, 21m));
}
