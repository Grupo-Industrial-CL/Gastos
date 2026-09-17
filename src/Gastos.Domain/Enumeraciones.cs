namespace Gastos.Domain;

/// <summary>Roles de acceso al portal (tabla GT_UsuariosRoles). Sin fila en esa tabla = sin acceso.</summary>
public static class Roles
{
    /// <summary>Ve todos los gastos de todas las sociedades, carga datos y configura usuarios.</summary>
    public const string Administrador = "Administrador";
    /// <summary>Solo consulta los gastos de las sociedades a las que tiene acceso (GT_UsuariosSociedades).</summary>
    public const string Consultor = "Consultor";

    public static readonly string[] Todos = { Administrador, Consultor };
    public static bool EsValido(string? rol) => rol is not null && Todos.Contains(rol, StringComparer.OrdinalIgnoreCase);
}

/// <summary>Origen de una carga de gastos.</summary>
public enum OrigenCarga
{
    /// <summary>Fichero CSV de LayTours (antes "Carga fichero LayTours"/ORBIS en el VB6).</summary>
    LayTours = 1,
    /// <summary>API de Repsol (tarjetas Solred). Sustituye a la carga del CSV de Solred del VB6.</summary>
    Repsol = 2,
    /// <summary>API de Moeve (tarjetas de carburante). Nueva en el portal.</summary>
    Moeve = 3
}

public static class OrigenCargaExtensiones
{
    /// <summary>Texto que se guarda en GT_GastosInternos.Origen y se muestra en el grid.</summary>
    public static string Codigo(this OrigenCarga o) => o switch
    {
        OrigenCarga.LayTours => "LayTours",
        OrigenCarga.Repsol => "Repsol",
        OrigenCarga.Moeve => "Moeve",
        _ => o.ToString()
    };
}
