namespace Gastos.Domain.Models;

/// <summary>Resultado de una carga (LayTours, Repsol o Moeve). Las cargas son "todo o nada": si hay error no se guarda ningún dato.</summary>
public class ResultadoCarga
{
    public bool Ok { get; set; }
    public int IdCarga { get; set; }
    public int Filas { get; set; }
    public int Insertados { get; set; }
    /// <summary>Movimientos ya existentes (duplicados) o que no generan gasto (importe negativo, tarjeta sin "insertar gasto").</summary>
    public int Omitidos { get; set; }
    public string Mensaje { get; set; } = string.Empty;
    public List<string> Avisos { get; set; } = new();

    public static ResultadoCarga Error(string mensaje, int filas = 0) => new() { Ok = false, Mensaje = mensaje, Filas = filas };
}
