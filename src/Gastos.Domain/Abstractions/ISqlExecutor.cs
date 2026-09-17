namespace Gastos.Domain.Abstractions;

/// <summary>
/// Contrato de acceso a datos con SQL parametrizado (Dapper sobre la conexión de EF Core).
/// Usar SIEMPRE parámetros (objeto anónimo) para evitar inyección SQL.
/// </summary>
public interface ISqlExecutor : ISqlComandos
{
    /// <summary>
    /// Ejecuta el trabajo dentro de una transacción: si el delegado lanza una excepción se hace ROLLBACK;
    /// si termina, COMMIT. Es la base de las cargas "todo o nada".
    /// </summary>
    Task<T> EnTransaccionAsync<T>(Func<ISqlComandos, Task<T>> trabajo, CancellationToken ct = default);
}

/// <summary>Operaciones SQL básicas; disponibles fuera y dentro de una transacción.</summary>
public interface ISqlComandos
{
    Task<IReadOnlyList<T>> QueryAsync<T>(string sql, object? parameters = null, CancellationToken ct = default);
    Task<T?> QueryFirstOrDefaultAsync<T>(string sql, object? parameters = null, CancellationToken ct = default);
    Task<int> ExecuteAsync(string sql, object? parameters = null, CancellationToken ct = default);
    Task<T?> ExecuteScalarAsync<T>(string sql, object? parameters = null, CancellationToken ct = default);
}
