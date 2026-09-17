using Gastos.Domain;
using Gastos.Domain.Abstractions;
using Gastos.Domain.Abstractions.Services;
using Gastos.Domain.Models;

namespace Gastos.Business.Services;

/// <summary>Roles y sociedades por usuario (GT_UsuariosRoles / GT_UsuariosSociedades).</summary>
public sealed class PermisosService : IPermisosService
{
    private readonly ISqlExecutor _sql;
    public PermisosService(ISqlExecutor sql) => _sql = sql;

    public async Task<PermisosUsuario> DamePermisosAsync(string email, CancellationToken ct = default)
    {
        var p = new PermisosUsuario { Email = email };
        if (string.IsNullOrWhiteSpace(email)) return p;
        p.Rol = await _sql.QueryFirstOrDefaultAsync<string>(
            "SELECT Rol FROM dbo.GT_UsuariosRoles WHERE Email = @email", new { email }, ct);
        if (p.TieneAcceso)
            p.Sociedades = (await _sql.QueryAsync<int>(
                "SELECT CodPais FROM dbo.GT_UsuariosSociedades WHERE Email = @email ORDER BY CodPais", new { email }, ct)).ToList();
        return p;
    }

    public async Task AsignarRolAsync(string email, string? rol, string usuarioAuditor, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(email)) throw new ArgumentException("Email obligatorio", nameof(email));
        if (string.IsNullOrWhiteSpace(rol))
        {
            // Retirar acceso: se elimina el rol y las sociedades asociadas.
            await _sql.ExecuteAsync(@"
                DELETE FROM dbo.GT_UsuariosSociedades WHERE Email = @email;
                DELETE FROM dbo.GT_UsuariosRoles WHERE Email = @email;", new { email }, ct);
            return;
        }
        if (!Roles.EsValido(rol)) throw new ArgumentException($"Rol no válido: {rol}", nameof(rol));
        rol = Roles.Todos.First(r => string.Equals(r, rol, StringComparison.OrdinalIgnoreCase));
        await _sql.ExecuteAsync(@"
            MERGE dbo.GT_UsuariosRoles AS t
            USING (SELECT @email AS Email) AS s ON t.Email = s.Email
            WHEN MATCHED THEN UPDATE SET Rol = @rol, FechaAlta = SYSDATETIME(), UsuarioAlta = @auditor
            WHEN NOT MATCHED THEN INSERT (Email, Rol, FechaAlta, UsuarioAlta) VALUES (@email, @rol, SYSDATETIME(), @auditor);",
            new { email, rol, auditor = usuarioAuditor }, ct);
    }

    public Task EstablecerSociedadAsync(string email, int codPais, bool asignada, CancellationToken ct = default)
        => asignada
            ? _sql.ExecuteAsync(@"
                IF NOT EXISTS (SELECT 1 FROM dbo.GT_UsuariosSociedades WHERE Email = @email AND CodPais = @codPais)
                    INSERT INTO dbo.GT_UsuariosSociedades (Email, CodPais) VALUES (@email, @codPais);", new { email, codPais }, ct)
            : _sql.ExecuteAsync("DELETE FROM dbo.GT_UsuariosSociedades WHERE Email = @email AND CodPais = @codPais", new { email, codPais }, ct);

    public Task EstablecerTodasSociedadesAsync(string email, bool asignar, CancellationToken ct = default)
        => asignar
            ? _sql.ExecuteAsync(@"
                INSERT INTO dbo.GT_UsuariosSociedades (Email, CodPais)
                SELECT @email, s.CodPais FROM dbo.GT_Sociedades s
                WHERE s.Activo = 1 AND NOT EXISTS (SELECT 1 FROM dbo.GT_UsuariosSociedades t WHERE t.Email = @email AND t.CodPais = s.CodPais);",
                new { email }, ct)
            : _sql.ExecuteAsync("DELETE FROM dbo.GT_UsuariosSociedades WHERE Email = @email", new { email }, ct);
}
