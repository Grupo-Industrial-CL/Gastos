using Microsoft.EntityFrameworkCore;

namespace Gastos.Data;

/// <summary>
/// DbContext MÍNIMO (misma estrategia que MasterPlan v2): EF Core gestiona la conexión y su ciclo de vida
/// por circuito Blazor; el acceso a datos se hace con SQL parametrizado vía <see cref="SqlExecutor"/> (Dapper).
/// Si más adelante conviene mapear entidades con LINQ, se añaden DbSet aquí sin tocar la capa de negocio.
/// </summary>
public class GastosDbContext : DbContext
{
    public GastosDbContext(DbContextOptions<GastosDbContext> options) : base(options) { }
}
