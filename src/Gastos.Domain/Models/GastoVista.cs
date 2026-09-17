namespace Gastos.Domain.Models;

/// <summary>Fila del grid de Gastos: gasto interno con los nombres de sociedad, departamento, empleado y tipo de gasto ya resueltos.</summary>
public class GastoVista
{
    public long Contador { get; set; }
    public int CodPais { get; set; }
    public string Sociedad { get; set; } = string.Empty;
    public int Departamento { get; set; }
    public string DepartamentoNombre { get; set; } = string.Empty;
    public long Codigo { get; set; }
    public string Empleado { get; set; } = string.Empty;
    public string Jerarquia { get; set; } = string.Empty;
    public string TipoGasto { get; set; } = string.Empty;
    public string TipoGastoNombre { get; set; } = string.Empty;
    public DateTime Fecha { get; set; }
    public decimal Importe { get; set; }
    public string Moneda { get; set; } = "EUR";
    public int Numero { get; set; }
    public string Donde { get; set; } = string.Empty;
    public string? Factura { get; set; }
    public string? Origen { get; set; }
    public string? Validado { get; set; }
    public string? Validacion { get; set; }
    public string? Pagado { get; set; }
    public string? ContaSAP { get; set; }
    public DateTime? FechaInsert { get; set; }
}
