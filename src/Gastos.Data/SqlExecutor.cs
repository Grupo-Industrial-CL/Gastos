using System.Data;
using System.Data.Common;
using Dapper;
using Gastos.Domain.Abstractions;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;

namespace Gastos.Data;

/// <summary>
/// Implementación de <see cref="ISqlExecutor"/> con Dapper sobre la conexión del <see cref="GastosDbContext"/>.
/// La conexión se comparte por circuito Blazor, así que los comandos se serializan con un semáforo (evita
/// "connection is busy" cuando dos componentes consultan a la vez). Las cargas usan <see cref="EnTransaccionAsync"/>.
/// </summary>
public sealed class SqlExecutor : ISqlExecutor
{
    private const int TimeoutSegundos = 600;
    private readonly GastosDbContext _db;
    private readonly ILogger<SqlExecutor> _logger;
    private readonly SemaphoreSlim _gate = new(1, 1);

    public SqlExecutor(GastosDbContext db, ILogger<SqlExecutor> logger)
    {
        _db = db;
        _logger = logger;
    }

    private DbConnection Connection => _db.Database.GetDbConnection();

    private async Task<T> RunAsync<T>(Func<Task<T>> op, CancellationToken ct)
    {
        await _gate.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            if (Connection.State != ConnectionState.Open) await Connection.OpenAsync(ct).ConfigureAwait(false);
            return await op().ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is not DataAccessException)
        {
            _logger.LogError(ex, "Error de acceso a datos");
            throw new DataAccessException($"Error de acceso a datos: {ex.Message}", ex);
        }
        finally
        {
            _gate.Release();
        }
    }

    public Task<IReadOnlyList<T>> QueryAsync<T>(string sql, object? parameters = null, CancellationToken ct = default)
        => RunAsync<IReadOnlyList<T>>(async () =>
            (await Connection.QueryAsync<T>(Cmd(sql, parameters, null, ct)).ConfigureAwait(false)).AsList(), ct);

    public Task<T?> QueryFirstOrDefaultAsync<T>(string sql, object? parameters = null, CancellationToken ct = default)
        => RunAsync(() => Connection.QueryFirstOrDefaultAsync<T>(Cmd(sql, parameters, null, ct)), ct);

    public Task<int> ExecuteAsync(string sql, object? parameters = null, CancellationToken ct = default)
        => RunAsync(() => Connection.ExecuteAsync(Cmd(sql, parameters, null, ct)), ct);

    public Task<T?> ExecuteScalarAsync<T>(string sql, object? parameters = null, CancellationToken ct = default)
        => RunAsync(() => Connection.ExecuteScalarAsync<T>(Cmd(sql, parameters, null, ct)), ct);

    public Task<T> EnTransaccionAsync<T>(Func<ISqlComandos, Task<T>> trabajo, CancellationToken ct = default)
        => RunAsync(async () =>
        {
            await using var tx = await Connection.BeginTransactionAsync(ct).ConfigureAwait(false);
            try
            {
                var resultado = await trabajo(new ComandosEnTransaccion(Connection, tx, ct)).ConfigureAwait(false);
                await tx.CommitAsync(ct).ConfigureAwait(false);
                return resultado;
            }
            catch
            {
                try { await tx.RollbackAsync(ct).ConfigureAwait(false); } catch { /* ya revertida */ }
                throw;
            }
        }, ct);

    private static CommandDefinition Cmd(string sql, object? p, IDbTransaction? tx, CancellationToken ct)
        => new(sql, p, tx, TimeoutSegundos, CommandType.Text, cancellationToken: ct);

    /// <summary>Comandos ligados a una transacción abierta (no pasan por el semáforo: ya lo tiene la transacción).</summary>
    private sealed class ComandosEnTransaccion : ISqlComandos
    {
        private readonly DbConnection _cn;
        private readonly DbTransaction _tx;
        private readonly CancellationToken _ct;
        public ComandosEnTransaccion(DbConnection cn, DbTransaction tx, CancellationToken ct) { _cn = cn; _tx = tx; _ct = ct; }

        public async Task<IReadOnlyList<T>> QueryAsync<T>(string sql, object? parameters = null, CancellationToken ct = default)
            => (await _cn.QueryAsync<T>(Cmd(sql, parameters, _tx, Tok(ct))).ConfigureAwait(false)).AsList();
        public Task<T?> QueryFirstOrDefaultAsync<T>(string sql, object? parameters = null, CancellationToken ct = default)
            => _cn.QueryFirstOrDefaultAsync<T>(Cmd(sql, parameters, _tx, Tok(ct)));
        public Task<int> ExecuteAsync(string sql, object? parameters = null, CancellationToken ct = default)
            => _cn.ExecuteAsync(Cmd(sql, parameters, _tx, Tok(ct)));
        public Task<T?> ExecuteScalarAsync<T>(string sql, object? parameters = null, CancellationToken ct = default)
            => _cn.ExecuteScalarAsync<T>(Cmd(sql, parameters, _tx, Tok(ct)));
        private CancellationToken Tok(CancellationToken ct) => ct.CanBeCanceled ? ct : _ct;
    }
}

/// <summary>Excepción de la capa de datos (envuelve el error del proveedor con un mensaje legible).</summary>
public sealed class DataAccessException : Exception
{
    public DataAccessException(string message, Exception inner) : base(message, inner) { }
}
