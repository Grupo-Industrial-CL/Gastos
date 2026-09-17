using Microsoft.AspNetCore.Components;

namespace Gastos.Web.Components.Common;

/// <summary>Descriptor de columna para <c>GtGrid&lt;T&gt;</c>: título, texto mostrado/buscado/filtrado, orden y plantilla opcionales.</summary>
public sealed class GtColumna<T>
{
    /// <summary>Título (también clave del filtro de columna).</summary>
    public string Titulo { get; set; } = string.Empty;
    /// <summary>Texto que se muestra, se busca y se filtra. Nunca nulo.</summary>
    public Func<T, string> Valor { get; set; } = _ => string.Empty;
    /// <summary>Valor de ordenación (fecha/número). Si es nulo se ordena por <see cref="Valor"/>.</summary>
    public Func<T, object>? Orden { get; set; }
    /// <summary>Valor numérico para exportar a Excel como número y para el filtro numérico (&gt;, &lt;, =).</summary>
    public Func<T, double?>? Numero { get; set; }
    public string? Ancho { get; set; }
    /// <summary>Celda personalizada (chip...). Si es nula se muestra el texto con resaltado del buscador.</summary>
    public RenderFragment<T>? Plantilla { get; set; }
    public bool Filtrable { get; set; } = true;
    /// <summary>Oculta la lista de valores (embudo) y deja solo el cuadro de texto.</summary>
    public bool SinEmbudo { get; set; }
    /// <summary>Filtro por comparación numérica (requiere <see cref="Numero"/>).</summary>
    public bool FiltroNumerico { get; set; }
    /// <summary>Alinea la celda a la derecha (importes).</summary>
    public bool Derecha { get; set; }
}
