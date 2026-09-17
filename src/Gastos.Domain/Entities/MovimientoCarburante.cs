using System.ComponentModel.DataAnnotations.Schema;

namespace Gastos.Domain.Entities;

/// <summary>
/// Detalle de un movimiento de tarjeta de carburante. Misma estructura para GT_GastosSolred (Repsol) y
/// GT_GastosMoeve (Moeve); Informix "gastos_solred". Cont_GastoInterno enlaza con el gasto interno
/// generado (nulo si el movimiento no generó gasto: descuentos en negativo o tarjeta sin "insertar gasto").
/// </summary>
public class MovimientoCarburante
{
    public long Id { get; set; }
    [Column("Cont_GastoInterno")] public long? ContGastoInterno { get; set; }
    [Column("Num_Tarjeta")] public string NumTarjeta { get; set; } = string.Empty;
    [Column("Anyo_Fact")] public int AnyoFact { get; set; }
    public string Factura { get; set; } = string.Empty;
    public string? Matricula { get; set; }
    public string? Conductor { get; set; }
    [Column("Fec_Operacion")] public DateTime FecOperacion { get; set; }
    [Column("Hor_Operacion")] public string HorOperacion { get; set; } = string.Empty;
    public string? Poblacion { get; set; }
    public decimal Km { get; set; }
    [Column("Des_Prod")] public string? DesProd { get; set; }
    [Column("Num_Litros")] public decimal NumLitros { get; set; }
    public decimal IVA { get; set; }
    [Column("Imp_Total")] public decimal ImpTotal { get; set; }
    [Column("Imp_Total_SinIVA")] public decimal ImpTotalSinIVA { get; set; }
    [Column("Precio_Litro")] public decimal PrecioLitro { get; set; }
    [Column("Info_Aux")] public string? InfoAux { get; set; }
    public int? IdCarga { get; set; }
    public DateTime? FechaCarga { get; set; }
}
