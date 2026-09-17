using Gastos.Domain.Abstractions;
using Gastos.Domain.Abstractions.Services;
using Gastos.Domain.Entities;

namespace Gastos.Business.Services;

public sealed class CatalogoService : ICatalogoService
{
    private readonly ISqlExecutor _sql;
    public CatalogoService(ISqlExecutor sql) => _sql = sql;

    public Task<IReadOnlyList<Sociedad>> DameSociedadesAsync(bool soloActivas = true, CancellationToken ct = default)
        => _sql.QueryAsync<Sociedad>(
            "SELECT CodPais, Nombre, NombreIngles, Activo FROM dbo.GT_Sociedades WHERE (@todas = 1 OR Activo = 1) ORDER BY Nombre",
            new { todas = soloActivas ? 0 : 1 }, ct);

    public Task<IReadOnlyList<Departamento>> DameDepartamentosAsync(int? codPais = null, CancellationToken ct = default)
        => _sql.QueryAsync<Departamento>(
            "SELECT CodPais, Codigo, Descripcion, CECO, Divisa, Activo FROM dbo.GT_Departamentos WHERE (@codPais IS NULL OR CodPais = @codPais) ORDER BY Descripcion",
            new { codPais }, ct);

    public Task<IReadOnlyList<TipoGasto>> DameTiposGastoAsync(CancellationToken ct = default)
        => _sql.QueryAsync<TipoGasto>("SELECT Tipo, Codigo, NombreGasto, Orden FROM dbo.GT_TiposGasto ORDER BY Tipo, Orden, NombreGasto", null, ct);
}
