using Gastos.Domain.Abstractions;
using Gastos.Domain.Entities;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Gastos.Data;

public static class DependencyInjection
{
    /// <summary>Registra el acceso a datos. Espera la cadena de conexión "Gastos" (ConnectionStrings:Gastos).</summary>
    public static IServiceCollection AddGastosData(this IServiceCollection services, IConfiguration configuration, string connectionStringName = "Gastos")
    {
        var cs = configuration.GetConnectionString(connectionStringName);
        if (string.IsNullOrWhiteSpace(cs) || cs.StartsWith("__"))
            throw new InvalidOperationException(
                $"No se encontró la cadena de conexión '{connectionStringName}'. Defínela con 'dotnet user-secrets set ConnectionStrings:{connectionStringName} \"...\"' (desarrollo) " +
                $"o con la variable de entorno ConnectionStrings__{connectionStringName} (servidor). Nunca en appsettings*.json.");
        return services.AddGastosData(cs);
    }

    public static IServiceCollection AddGastosData(this IServiceCollection services, string connectionString)
    {
        // MARS: varios componentes del mismo circuito pueden consultar a la vez.
        connectionString = new Microsoft.Data.SqlClient.SqlConnectionStringBuilder(connectionString) { MultipleActiveResultSets = true }.ConnectionString;

        services.AddDbContext<GastosDbContext>(o => o.UseSqlServer(connectionString, sql =>
        {
            sql.EnableRetryOnFailure(maxRetryCount: 3);
            sql.CommandTimeout(600);
        }));
        services.AddScoped<ISqlExecutor, SqlExecutor>();

        DapperColumnMapping.Initialize(typeof(GastoInterno).Assembly);
        return services;
    }
}
