using Gastos.Domain.Abstractions;
using Gastos.Domain.Abstractions.Services;
using Gastos.Domain.Entities;

namespace Gastos.Business.Services;

/// <summary>Directorio de usuarios (MP_Usuarios, sincronizado desde Azure AD; solo lectura) + rol del portal (GT_UsuariosRoles).</summary>
public sealed class UsuarioService : IUsuarioService
{
    private readonly ISqlExecutor _sql;
    public UsuarioService(ISqlExecutor sql) => _sql = sql;

    private const string SelectBase = @"
        SELECT u.usId, u.usObjectId, u.usDisplayName, u.usFirstName, u.usLastName, u.usUserMail,
               u.usCompanyName, u.usDepartment, u.usJobTitle, u.usManagerId, u.usManagerDisplayName, u.usIsActive,
               r.Rol,
               (SELECT COUNT(1) FROM dbo.GT_UsuariosSociedades s WHERE s.Email = u.usId) AS NumSociedades
        FROM dbo.MP_Usuarios u
        LEFT JOIN dbo.GT_UsuariosRoles r ON r.Email = u.usId";

    // El correo del token de Entra varía entre entornos (@gicl.es = usId, o el UPN = usUserMail), por eso se
    // busca por las dos columnas y por el oid de Azure (inmutable). Mismo criterio que MasterPlan v2.
    public Task<MPUsuario?> DameUsuarioAsync(string email, string? objectId = null, CancellationToken ct = default)
        => _sql.QueryFirstOrDefaultAsync<MPUsuario>(SelectBase + @"
            WHERE u.usId = @email
               OR (@email <> '' AND u.usUserMail = @email)
               OR (@oid <> '' AND CONVERT(NVARCHAR(50), u.usObjectId) = @oid)
            ORDER BY CASE WHEN u.usId = @email THEN 0 ELSE 1 END",
            new { email = email ?? string.Empty, oid = objectId ?? string.Empty }, ct);

    public Task<IReadOnlyList<MPUsuario>> DameUsuariosAsync(CancellationToken ct = default)
        => _sql.QueryAsync<MPUsuario>(SelectBase + " ORDER BY u.usDisplayName", null, ct);
}
