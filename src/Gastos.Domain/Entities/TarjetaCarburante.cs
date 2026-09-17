using System.ComponentModel.DataAnnotations.Schema;

namespace Gastos.Domain.Entities;

/// <summary>Tarjeta de carburante asignada a un empleado (GT_TarjetasSolred; Informix "tarjeta_solred"). Proveedor: Repsol o Moeve.</summary>
public class TarjetaCarburante
{
    public string NumTarjeta { get; set; } = string.Empty;
    public string Proveedor { get; set; } = "Repsol";
    public int CodPais { get; set; }
    public string? Titular { get; set; }
    public string? Empresa { get; set; }
    public string? CIF { get; set; }
    /// <summary>-1 en el VB6 significaba "red comercial" (no entra en este portal); en otro caso, departamento del empleado.</summary>
    public int Departamento { get; set; }
    public long Codigo { get; set; }
    public string Jerarquia { get; set; } = string.Empty;
    /// <summary>'S' = los movimientos de la tarjeta generan gasto interno; otro valor = solo se guarda el detalle del movimiento.</summary>
    [Column("Ins_Gasto_Int")] public string? InsGastoInt { get; set; }
    public bool Activo { get; set; } = true;
}
