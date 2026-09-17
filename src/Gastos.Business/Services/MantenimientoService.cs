using Gastos.Domain.Abstractions;
using Gastos.Domain.Abstractions.Services;
using Gastos.Domain.Entities;
using Gastos.Domain.Models;

namespace Gastos.Business.Services;

/// <summary>Mantenimiento de departamentos, empleados y tarjetas. Altas y modificaciones; las bajas son lógicas (Activo / FechaBaja).</summary>
public sealed class MantenimientoService : IMantenimientoService
{
    private readonly ISqlExecutor _sql;
    public MantenimientoService(ISqlExecutor sql) => _sql = sql;

    // ───────────────────────────── Departamentos ─────────────────────────────

    public Task<IReadOnlyList<DepartamentoVista>> DameDepartamentosAsync(CancellationToken ct = default)
        => _sql.QueryAsync<DepartamentoVista>(@"
            SELECT d.CodPais, d.Codigo, d.Descripcion, d.CECO, d.Divisa, d.Activo, ISNULL(s.Nombre, CONVERT(NVARCHAR(10), d.CodPais)) AS SociedadNombre
            FROM dbo.GT_Departamentos d LEFT JOIN dbo.GT_Sociedades s ON s.CodPais = d.CodPais
            ORDER BY s.Nombre, d.Descripcion", null, ct);

    public async Task GuardarDepartamentoAsync(Departamento d, bool esAlta, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(d.Descripcion)) throw new ArgumentException("La descripción es obligatoria.");
        if (esAlta)
        {
            var existe = await _sql.ExecuteScalarAsync<int>("SELECT COUNT(1) FROM dbo.GT_Departamentos WHERE CodPais = @CodPais AND Codigo = @Codigo", new { d.CodPais, d.Codigo }, ct);
            if (existe > 0) throw new InvalidOperationException($"Ya existe el departamento {d.Codigo} en esa sociedad.");
            await _sql.ExecuteAsync(@"INSERT INTO dbo.GT_Departamentos (CodPais, Codigo, Descripcion, CECO, Divisa, Activo)
                                      VALUES (@CodPais, @Codigo, @Descripcion, @CECO, @Divisa, @Activo)", d, ct);
        }
        else
        {
            await _sql.ExecuteAsync(@"UPDATE dbo.GT_Departamentos SET Descripcion = @Descripcion, CECO = @CECO, Divisa = @Divisa, Activo = @Activo
                                      WHERE CodPais = @CodPais AND Codigo = @Codigo", d, ct);
        }
    }

    // ───────────────────────────── Empleados ─────────────────────────────

    public Task<IReadOnlyList<EmpleadoVista>> DameEmpleadosAsync(int? codPais = null, CancellationToken ct = default)
        => _sql.QueryAsync<EmpleadoVista>(@"
            SELECT u.CodPais, u.Departamento, u.Codigo, u.Nombre, u.DNI, u.Jerarquia, u.Id_Epsilon, u.Usuario, u.Cuenta_Bancaria, u.FechaBaja, u.Bloqueado_Por,
                   ISNULL(s.Nombre, CONVERT(NVARCHAR(10), u.CodPais)) AS SociedadNombre, ISNULL(d.Descripcion, CONVERT(NVARCHAR(10), u.Departamento)) AS DepartamentoNombre
            FROM dbo.GT_UsuariosPersonal u
            LEFT JOIN dbo.GT_Sociedades s ON s.CodPais = u.CodPais
            LEFT JOIN dbo.GT_Departamentos d ON d.CodPais = u.CodPais AND d.Codigo = u.Departamento
            WHERE (@codPais IS NULL OR u.CodPais = @codPais)
            ORDER BY s.Nombre, u.Nombre", new { codPais }, ct);

    public async Task GuardarEmpleadoAsync(UsuarioPersonal u, bool esAlta, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(u.Nombre)) throw new ArgumentException("El nombre es obligatorio.");
        u.Jerarquia ??= string.Empty;
        var p = new
        {
            u.CodPais, u.Departamento, u.Codigo, u.Nombre, u.DNI, u.Jerarquia, IdEpsilon = u.IdEpsilon, u.Usuario,
            CuentaBancaria = u.CuentaBancaria, u.FechaBaja, BloqueadoPor = u.BloqueadoPor
        };
        if (esAlta)
        {
            var existe = await _sql.ExecuteScalarAsync<int>(
                "SELECT COUNT(1) FROM dbo.GT_UsuariosPersonal WHERE CodPais = @CodPais AND Departamento = @Departamento AND Codigo = @Codigo", p, ct);
            if (existe > 0) throw new InvalidOperationException($"Ya existe el empleado {u.Codigo} en ese departamento.");
            await _sql.ExecuteAsync(@"
                INSERT INTO dbo.GT_UsuariosPersonal (CodPais, Departamento, Codigo, Nombre, DNI, Jerarquia, Id_Epsilon, Usuario, Cuenta_Bancaria, FechaBaja, Bloqueado_Por)
                VALUES (@CodPais, @Departamento, @Codigo, @Nombre, @DNI, @Jerarquia, @IdEpsilon, @Usuario, @CuentaBancaria, @FechaBaja, @BloqueadoPor)", p, ct);
        }
        else
        {
            await _sql.ExecuteAsync(@"
                UPDATE dbo.GT_UsuariosPersonal SET Nombre = @Nombre, DNI = @DNI, Jerarquia = @Jerarquia, Id_Epsilon = @IdEpsilon, Usuario = @Usuario,
                       Cuenta_Bancaria = @CuentaBancaria, FechaBaja = @FechaBaja
                WHERE CodPais = @CodPais AND Departamento = @Departamento AND Codigo = @Codigo", p, ct);
        }
    }

    // ───────────────────────────── Tarjetas ─────────────────────────────

    public Task<IReadOnlyList<TarjetaVista>> DameTarjetasAsync(CancellationToken ct = default)
        => _sql.QueryAsync<TarjetaVista>(@"
            SELECT t.NumTarjeta, t.Proveedor, t.CodPais, t.Titular, t.Empresa, t.CIF, t.Departamento, t.Codigo, t.Jerarquia, t.Ins_Gasto_Int, t.Activo,
                   ISNULL(s.Nombre, CONVERT(NVARCHAR(10), t.CodPais)) AS SociedadNombre,
                   CASE WHEN t.Departamento = -1 THEN 'Red comercial' ELSE ISNULL(d.Descripcion, CONVERT(NVARCHAR(10), t.Departamento)) END AS DepartamentoNombre,
                   ISNULL(u.Nombre, '') AS EmpleadoNombre
            FROM dbo.GT_TarjetasSolred t
            LEFT JOIN dbo.GT_Sociedades s ON s.CodPais = t.CodPais
            LEFT JOIN dbo.GT_Departamentos d ON d.CodPais = t.CodPais AND d.Codigo = t.Departamento
            LEFT JOIN dbo.GT_UsuariosPersonal u ON u.CodPais = t.CodPais AND u.Departamento = t.Departamento AND u.Codigo = t.Codigo
            ORDER BY t.Proveedor, s.Nombre, t.NumTarjeta", null, ct);

    public async Task GuardarTarjetaAsync(TarjetaCarburante t, bool esAlta, CancellationToken ct = default)
    {
        t.NumTarjeta = Cargas.ClasificadorCarburante.NormalizarTarjeta(t.NumTarjeta);
        if (t.NumTarjeta.Length == 0) throw new ArgumentException("El número de tarjeta es obligatorio.");
        if (t.Proveedor is not ("Repsol" or "Moeve")) throw new ArgumentException("Proveedor no válido (Repsol o Moeve).");
        t.Jerarquia ??= string.Empty;
        var p = new { t.NumTarjeta, t.Proveedor, t.CodPais, t.Titular, t.Empresa, t.CIF, t.Departamento, t.Codigo, t.Jerarquia, InsGastoInt = t.InsGastoInt, t.Activo };
        if (esAlta)
        {
            var existe = await _sql.ExecuteScalarAsync<int>("SELECT COUNT(1) FROM dbo.GT_TarjetasSolred WHERE NumTarjeta = @NumTarjeta", p, ct);
            if (existe > 0) throw new InvalidOperationException("La tarjeta ya existe en el sistema.");
            await _sql.ExecuteAsync(@"
                INSERT INTO dbo.GT_TarjetasSolred (NumTarjeta, Proveedor, CodPais, Titular, Empresa, CIF, Departamento, Codigo, Jerarquia, Ins_Gasto_Int, Activo)
                VALUES (@NumTarjeta, @Proveedor, @CodPais, @Titular, @Empresa, @CIF, @Departamento, @Codigo, @Jerarquia, @InsGastoInt, @Activo)", p, ct);
        }
        else
        {
            await _sql.ExecuteAsync(@"
                UPDATE dbo.GT_TarjetasSolred SET Proveedor = @Proveedor, CodPais = @CodPais, Titular = @Titular, Empresa = @Empresa, CIF = @CIF,
                       Departamento = @Departamento, Codigo = @Codigo, Jerarquia = @Jerarquia, Ins_Gasto_Int = @InsGastoInt, Activo = @Activo
                WHERE NumTarjeta = @NumTarjeta", p, ct);
        }
    }
}
