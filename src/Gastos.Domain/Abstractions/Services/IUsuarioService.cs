using Gastos.Domain.Entities;

namespace Gastos.Domain.Abstractions.Services;

/// <summary>Directorio de usuarios de Azure (MP_Usuarios, solo lectura) con su rol en el portal.</summary>
public interface IUsuarioService
{
    /// <summary>Busca al usuario por correo (usId o usUserMail) o por el oid de Azure. Nulo si no está en el directorio.</summary>
    Task<MPUsuario?> DameUsuarioAsync(string email, string? objectId = null, CancellationToken ct = default);
    Task<IReadOnlyList<MPUsuario>> DameUsuariosAsync(CancellationToken ct = default);
}
