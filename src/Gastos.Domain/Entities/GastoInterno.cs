using System.ComponentModel.DataAnnotations.Schema;

namespace Gastos.Domain.Entities;

/// <summary>Gasto interno de un empleado (GT_GastosInternos; Informix "gastos_internos"). Es la tabla que alimentan las cargas.</summary>
public class GastoInterno
{
    public long Contador { get; set; }
    public int CodPais { get; set; }
    public long Codigo { get; set; }
    public int Departamento { get; set; }
    public string Jerarquia { get; set; } = string.Empty;
    public string TipoGasto { get; set; } = string.Empty;
    public DateTime Fecha { get; set; }
    public decimal Importe { get; set; }
    [Column("Importe_Ori")] public decimal? ImporteOri { get; set; }
    public string Moneda { get; set; } = "EUR";
    public int Numero { get; set; }
    public int Numero2 { get; set; }
    public string Donde { get; set; } = string.Empty;
    [Column("Tipo_Com_Inv")] public string? TipoComInv { get; set; }
    public string? VISA { get; set; }
    public string? Validado { get; set; }
    public string? Validacion { get; set; }
    public string? Pagado { get; set; }
    public string? Factura { get; set; }
    public string? FamiliaGasto { get; set; }
    public string? Proveedor { get; set; }
    public int? Personas { get; set; }
    [Column("PEP_CECO")] public string? PepCeco { get; set; }
    [Column("Conta_SAP")] public string? ContaSAP { get; set; }
    [Column("Documento_SAP")] public string? DocumentoSAP { get; set; }
    [Column("Archivado_En")] public string? ArchivadoEn { get; set; }
    [Column("Fecha_Insert")] public DateTime? FechaInsert { get; set; }
    /// <summary>Origen de la carga (LayTours / Repsol / Moeve). Nulo en registros migrados sin origen conocido.</summary>
    public string? Origen { get; set; }
    public int? IdCarga { get; set; }
}
