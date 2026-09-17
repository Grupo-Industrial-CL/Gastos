using Gastos.Business.Cargas;
using Xunit;

namespace Gastos.Tests;

public class LayToursParserTests
{
    private static ResultadoParseoLayTours Parse(string contenido) => LayToursParser.Parsear(new StringReader(contenido));

    [Fact]
    public void Linea_valida_de_hotel_se_parsea_con_dias_y_donde_albaran_detalle()
    {
        var r = Parse("F-001;Hotel;250.50;12345;ALB-9;Hotel NH Madrid;15012026;3;Madrid");
        Assert.True(r.Ok, r.Error);
        var l = Assert.Single(r.Lineas);
        Assert.Equal("F-001", l.Factura);
        Assert.Equal("Hoteles", l.NombreGasto);
        Assert.True(l.EsHotel);
        Assert.Equal(250.50m, l.Importe);
        Assert.Equal(new DateTime(2026, 1, 15), l.Fecha);
        Assert.Equal(3, l.NumeroDias);
        Assert.Equal("ALB-9/Hotel NH Madrid", l.Donde);
    }

    [Fact]
    public void Billete_usa_destino_en_donde_y_admite_fecha_de_7_digitos()
    {
        var r = Parse("F-2;AVION;120;7;ALB;Detalle;5032026;0;Lisboa");
        Assert.True(r.Ok, r.Error);
        var l = r.Lineas[0];
        Assert.Equal("Billetes", l.NombreGasto);
        Assert.Equal("ALB/Lisboa", l.Donde);
        Assert.Equal(new DateTime(2026, 3, 5), l.Fecha);
    }

    [Theory]
    [InlineData("rent a car", "Otros gastos")]
    [InlineData("Gasto Gestion", "Otros gastos")]
    [InlineData("TREN", "Billetes")]
    [InlineData("no usar  avion / tren / barco", "Billetes")]
    public void Conceptos_se_clasifican_como_el_vb6(string concepto, string esperado)
    {
        var r = Parse($"F;{concepto};10;1;A;D;01012026;1;X");
        Assert.True(r.Ok, r.Error);
        Assert.Equal(esperado, r.Lineas[0].NombreGasto);
    }

    [Fact]
    public void Concepto_desconocido_falla_indicando_la_fila()
    {
        var r = Parse("F;Hotel;10;1;A;D;01012026;1;X\nF;Taxi;10;1;A;D;01012026;1;X");
        Assert.False(r.Ok);
        Assert.Contains("Accounting concept", r.Error);
        Assert.Contains("fila 2", r.Error);
    }

    [Theory]
    [InlineData("F;Hotel;10;1;A;D;01012026;1", "columnas")]                 // 8 campos
    [InlineData(";Hotel;10;1;A;D;01012026;1;X", "Invoice number")]
    [InlineData("F;Hotel;abc;1;A;D;01012026;1;X", "Amount")]
    [InlineData("F;Hotel;10;pepe;A;D;01012026;1;X", "User ID")]
    [InlineData("F;Hotel;10;1;A;D;31022026;1;X", "Fecha inicio servicio")]  // 31 de febrero
    [InlineData("F;Hotel;10;1;A;D;01132026;1;X", "Fecha inicio servicio")]  // mes 13
    [InlineData("F;Hotel;10;1;A;D;01012026;x;X", "Num de dias")]
    public void Validaciones_del_vb6(string linea, string textoError)
    {
        var r = Parse(linea);
        Assert.False(r.Ok);
        Assert.Contains(textoError, r.Error);
    }

    [Fact]
    public void Bisiesto_acepta_29_de_febrero_y_no_bisiesto_lo_rechaza()
    {
        Assert.True(Parse("F;Hotel;10;1;A;D;29022024;1;X").Ok);
        Assert.False(Parse("F;Hotel;10;1;A;D;29022026;1;X").Ok);
    }

    [Fact]
    public void Limpiar_cadena_sustituye_apostrofos_y_comillas()
    {
        var r = Parse("F;Otros;10;1;AL'B;De\"talle;01012026;1;X");
        Assert.True(r.Ok, r.Error);
        Assert.Equal("AL`B", r.Lineas[0].NumeroAlbaran);
        Assert.Equal("De`talle", r.Lineas[0].Detalle);
    }

    [Fact]
    public void Ultima_linea_vacia_se_tolera_pero_una_intermedia_no()
    {
        Assert.True(Parse("F;Hotel;10;1;A;D;01012026;1;X\n").Ok);
        var r = Parse("F;Hotel;10;1;A;D;01012026;1;X\n\nF;Hotel;10;1;A;D;01012026;1;X");
        Assert.False(r.Ok);
        Assert.Contains("Línea vacía", r.Error);
    }

    [Fact]
    public void Importe_con_coma_decimal_espanola_se_acepta()
    {
        var r = Parse("F;Hotel;1.250,75;1;A;D;01012026;1;X");
        Assert.True(r.Ok, r.Error);
        Assert.Equal(1250.75m, r.Lineas[0].Importe);
    }
}
