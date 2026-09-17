using Gastos.Domain.Entities;

namespace Gastos.Domain.Models;

/// <summary>Tarjeta con los nombres resueltos para el grid de mantenimiento.</summary>
public class TarjetaVista : TarjetaCarburante
{
    public string SociedadNombre { get; set; } = string.Empty;
    public string DepartamentoNombre { get; set; } = string.Empty;
    public string EmpleadoNombre { get; set; } = string.Empty;
}

/// <summary>Empleado con los nombres resueltos para el grid de mantenimiento.</summary>
public class EmpleadoVista : UsuarioPersonal
{
    public string SociedadNombre { get; set; } = string.Empty;
    public string DepartamentoNombre { get; set; } = string.Empty;
}

/// <summary>Departamento con el nombre de la sociedad.</summary>
public class DepartamentoVista : Departamento
{
    public string SociedadNombre { get; set; } = string.Empty;
}
