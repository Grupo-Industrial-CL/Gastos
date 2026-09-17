using Gastos.Domain.Abstractions;
using Gastos.Domain.Abstractions.Services;
using Gastos.Domain.Models;

namespace Gastos.Business.Services;

public sealed class GastoService : IGastoService
{
    private readonly ISqlExecutor _sql;
    public GastoService(ISqlExecutor sql) => _sql = sql;

    public async Task<IReadOnlyList<GastoVista>> DameGastosAsync(IReadOnlyCollection<int>? sociedades, DateTime? desde, DateTime? hasta, CancellationToken ct = default)
    {
        // Lista vacía = el usuario no tiene ninguna sociedad permitida: no se consulta nada.
        if (sociedades is { Count: 0 }) return Array.Empty<GastoVista>();

        // g.CodPais > 0: el VB6 "borraba" un gasto poniendo CodPais en negativo (CodPais * -1); se respeta esa marca.
        var sql = @"
            SELECT g.Contador, g.CodPais, ISNULL(s.Nombre, CONVERT(NVARCHAR(10), g.CodPais)) AS Sociedad,
                   g.Departamento, ISNULL(d.Descripcion, '') AS DepartamentoNombre,
                   g.Codigo, ISNULL(u.Nombre, '') AS Empleado, g.Jerarquia,
                   g.TipoGasto, ISNULL(t.NombreGasto, g.TipoGasto) AS TipoGastoNombre,
                   g.Fecha, g.Importe, g.Moneda, g.Numero, g.Donde, g.Factura, g.Origen,
                   g.Validado, g.Validacion, g.Pagado, g.Conta_SAP AS ContaSAP, g.Fecha_Insert AS FechaInsert
            FROM dbo.GT_GastosInternos g
            LEFT JOIN dbo.GT_Sociedades s ON s.CodPais = g.CodPais
            LEFT JOIN dbo.GT_Departamentos d ON d.CodPais = g.CodPais AND d.Codigo = g.Departamento
            LEFT JOIN dbo.GT_UsuariosPersonal u ON u.CodPais = g.CodPais AND u.Departamento = g.Departamento AND u.Codigo = g.Codigo
            OUTER APPLY (
                SELECT TOP 1 t.NombreGasto
                FROM dbo.GT_TiposGasto t
                LEFT JOIN dbo.GT_GastosSociedades gs ON gs.CodPais = g.CodPais
                WHERE t.Codigo = g.TipoGasto
                ORDER BY CASE WHEN t.Tipo = gs.Id THEN 0 WHEN t.Tipo = 'I' THEN 1 ELSE 2 END
            ) t
            WHERE g.CodPais > 0
              AND (@desde IS NULL OR g.Fecha >= @desde)
              AND (@hasta IS NULL OR g.Fecha <= @hasta)"
            + (sociedades is null ? string.Empty : " AND g.CodPais IN @socs")
            + " ORDER BY g.Fecha DESC, g.Contador DESC";

        return await _sql.QueryAsync<GastoVista>(sql, new { desde, hasta, socs = sociedades }, ct);
    }
}
