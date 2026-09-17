namespace Gastos.Domain.Entities;

/// <summary>Empresa de gasto (GT_GastosSociedades; Informix "gastos_empresas"). Su Id es el "Tipo" con el que se buscan los tipos de gasto de sociedades no estándar.</summary>
public class GastoSociedad
{
    public string Id { get; set; } = string.Empty;
    public int CodPais { get; set; }
    public string Empresa { get; set; } = string.Empty;
    public int Orden { get; set; }
}
