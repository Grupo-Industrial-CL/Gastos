namespace Gastos.Domain.Models;

/// <summary>Resultado de la comprobación previa de una carga (equivale a los botones "Comprobar" del VB6): qué falta por dar de alta antes de cargar.</summary>
public class ResultadoComprobacion
{
    /// <summary>True si el fichero/la API se han leído bien (aunque haya elementos faltantes).</summary>
    public bool Ok { get; set; }
    public int Filas { get; set; }
    public string Mensaje { get; set; } = string.Empty;
    /// <summary>Tarjetas o usuarios Epsilon que no existen en el sistema. Vacío = se puede cargar.</summary>
    public List<ElementoFaltante> Faltantes { get; set; } = new();
    public List<string> Avisos { get; set; } = new();
    public bool PuedeCargar => Ok && Faltantes.Count == 0;
}

public class ElementoFaltante
{
    public string Clave { get; set; } = string.Empty;    // número de tarjeta / id Epsilon
    public int Ocurrencias { get; set; }
    public string Detalle { get; set; } = string.Empty;  // conductor, matrícula, importe acumulado, fila...
}
