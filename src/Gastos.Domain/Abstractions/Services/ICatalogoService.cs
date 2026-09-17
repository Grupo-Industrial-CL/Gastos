using Gastos.Domain.Entities;

namespace Gastos.Domain.Abstractions.Services;

/// <summary>Catálogos maestros: sociedades, departamentos, tipos de gasto.</summary>
public interface ICatalogoService
{
    Task<IReadOnlyList<Sociedad>> DameSociedadesAsync(bool soloActivas = true, CancellationToken ct = default);
    Task<IReadOnlyList<Departamento>> DameDepartamentosAsync(int? codPais = null, CancellationToken ct = default);
    Task<IReadOnlyList<TipoGasto>> DameTiposGastoAsync(CancellationToken ct = default);
}
