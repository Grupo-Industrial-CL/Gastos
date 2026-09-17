using Gastos.Domain.Entities;
using Gastos.Domain.Models;

namespace Gastos.Domain.Abstractions.Services;

/// <summary>Mantenimiento de catálogos desde el portal (solo Administrador): departamentos, empleados y tarjetas de carburante. Sin borrados físicos.</summary>
public interface IMantenimientoService
{
    Task<IReadOnlyList<DepartamentoVista>> DameDepartamentosAsync(CancellationToken ct = default);
    Task GuardarDepartamentoAsync(Departamento d, bool esAlta, CancellationToken ct = default);

    Task<IReadOnlyList<EmpleadoVista>> DameEmpleadosAsync(int? codPais = null, CancellationToken ct = default);
    Task GuardarEmpleadoAsync(UsuarioPersonal u, bool esAlta, CancellationToken ct = default);

    Task<IReadOnlyList<TarjetaVista>> DameTarjetasAsync(CancellationToken ct = default);
    Task GuardarTarjetaAsync(TarjetaCarburante t, bool esAlta, CancellationToken ct = default);
}
