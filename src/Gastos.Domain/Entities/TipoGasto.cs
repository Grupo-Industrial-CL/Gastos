namespace Gastos.Domain.Entities;

/// <summary>Tipo de gasto (GT_TiposGasto; Informix "tiposgasto"). "Tipo" agrupa por sociedad ('I' = interno estándar o el Id de GT_GastosSociedades).</summary>
public class TipoGasto
{
    public string Tipo { get; set; } = "I";
    public string Codigo { get; set; } = string.Empty;      // Informix: TipoGasto (p. ej. "1" Gasolina, "9" Peaje, "10" Otros gastos, "11" Parking)
    public string NombreGasto { get; set; } = string.Empty; // p. ej. "Billetes", "Hoteles", "Otros gastos"
    public int Orden { get; set; }
}
