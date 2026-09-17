using Gastos.Domain.Abstractions;
using Gastos.Domain.Abstractions.Services;
using Gastos.Domain.Models;

namespace Gastos.Business.Services;

public sealed class GastoService : IGastoService
{
    private readonly ISqlExecutor _sql;
    public GastoService(ISqlExecutor sql) => _sql = sql;

    public async Task<GastoDetalle?> DameDetalleAsync(long contador, IReadOnlyCollection<int>? sociedades, CancellationToken ct = default)
    {
        if (sociedades is { Count: 0 }) return null;
        var vista = (await DameGastosAsync(sociedades, null, null, contador, ct)).FirstOrDefault();
        if (vista is null) return null;

        var det = new GastoDetalle { Gasto = vista };
        det.Registro = await _sql.QueryFirstOrDefaultAsync<Gastos.Domain.Entities.GastoInterno>(
            "SELECT * FROM dbo.GT_GastosInternos WHERE Contador = @contador", new { contador }, ct) ?? new();

        // Movimiento de tarjeta asociado (Repsol o Moeve).
        det.Movimiento = await _sql.QueryFirstOrDefaultAsync<Gastos.Domain.Entities.MovimientoCarburante>(
            "SELECT TOP 1 * FROM dbo.GT_GastosSolred WHERE Cont_GastoInterno = @contador", new { contador }, ct);
        if (det.Movimiento is not null) det.ProveedorMovimiento = "Repsol";
        else
        {
            det.Movimiento = await _sql.QueryFirstOrDefaultAsync<Gastos.Domain.Entities.MovimientoCarburante>(
                "SELECT TOP 1 * FROM dbo.GT_GastosMoeve WHERE Cont_GastoInterno = @contador", new { contador }, ct);
            if (det.Movimiento is not null) det.ProveedorMovimiento = "Moeve";
        }

        if (det.Registro.IdCarga is int idCarga)
            det.Carga = await _sql.QueryFirstOrDefaultAsync<Gastos.Domain.Entities.Carga>(
                "SELECT Id, Origen, Fecha, Usuario, Desde, Hasta, CodPais, NombreFichero, Filas, Insertados, Omitidos, Estado, Mensaje FROM dbo.GT_Cargas WHERE Id = @idCarga",
                new { idCarga }, ct);
        return det;
    }

    public Task<IReadOnlyList<GastoVista>> DameGastosAsync(IReadOnlyCollection<int>? sociedades, DateTime? desde, DateTime? hasta, CancellationToken ct = default)
        => DameGastosAsync(sociedades, desde, hasta, null, ct);

    private async Task<IReadOnlyList<GastoVista>> DameGastosAsync(IReadOnlyCollection<int>? sociedades, DateTime? desde, DateTime? hasta, long? contador, CancellationToken ct)
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
              AND (@hasta IS NULL OR g.Fecha <= @hasta)
              AND (@contador IS NULL OR g.Contador = @contador)"
            + (sociedades is null ? string.Empty : " AND g.CodPais IN @socs")
            + " ORDER BY g.Fecha DESC, g.Contador DESC";

        return await _sql.QueryAsync<GastoVista>(sql, new { desde, hasta, contador, socs = sociedades }, ct);
    }
}
