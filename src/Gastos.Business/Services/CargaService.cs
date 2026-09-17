using System.Text;
using Gastos.Business.Cargas;
using Gastos.Domain;
using Gastos.Domain.Abstractions;
using Gastos.Domain.Abstractions.Services;
using Gastos.Domain.Entities;
using Gastos.Domain.Models;
using Microsoft.Extensions.Logging;

namespace Gastos.Business.Services;

/// <summary>
/// Cargas de gastos internos. Port de las dos utilidades del VB6 que se conservan:
///  - LayTours (frmCargarFicheroORBIS): CSV subido por el administrador, para una sociedad.
///  - Solred (frmCargaFicheroSolred2): ahora los movimientos llegan de la API de Repsol (o de Moeve); se conserva
///    la clasificación por descripción de producto y el enlace gasto interno <-> detalle del movimiento.
/// Todas las cargas son "todo o nada" (transacción) y quedan registradas en GT_Cargas.
/// Solo se contemplan gastos INTERNOS: los movimientos de tarjetas de la red comercial (Departamento = -1 en la
/// tarjeta) se guardan como detalle pero no generan gasto (decisión de alcance del portal).
/// </summary>
public sealed class CargaService : ICargaService
{
    private static readonly int[] PaisesTipoInterno = { 34, 39, 51, 52, 56, 12, 16, 18 };

    private readonly ISqlExecutor _sql;
    private readonly IReadOnlyList<IProveedorTarjetaCarburante> _proveedores;
    private readonly ILogger<CargaService> _log;

    public CargaService(ISqlExecutor sql, IEnumerable<IProveedorTarjetaCarburante> proveedores, ILogger<CargaService> log)
    {
        _sql = sql;
        _proveedores = proveedores.ToList();
        _log = log;
    }

    public IReadOnlyList<(OrigenCarga Origen, bool Configurado, string Detalle)> EstadoProveedores()
        => _proveedores.Select(p => (p.Origen, p.Configurado, p.Detalle)).ToList();

    public Task<IReadOnlyList<Carga>> DameCargasAsync(int ultimas = 200, CancellationToken ct = default)
        => _sql.QueryAsync<Carga>(@"
            SELECT TOP (@n) Id, Origen, Fecha, Usuario, Desde, Hasta, CodPais, NombreFichero, Filas, Insertados, Omitidos, Estado, Mensaje
            FROM dbo.GT_Cargas ORDER BY Id DESC", new { n = ultimas }, ct);

    // ───────────────────────────── Comprobaciones previas (sin guardar nada) ─────────────────────────────

    public async Task<ResultadoComprobacion> ComprobarLayToursAsync(Stream csv, int codPais, CancellationToken ct = default)
    {
        ResultadoParseoLayTours parseo;
        using (var lector = new StreamReader(csv, Encoding.Latin1, detectEncodingFromByteOrderMarks: true))
            parseo = LayToursParser.Parsear(lector);
        var r = new ResultadoComprobacion { Ok = parseo.Ok, Filas = parseo.Lineas.Count };
        if (!parseo.Ok) { r.Mensaje = parseo.Error!; return r; }

        var ids = parseo.Lineas.Select(l => l.UsuarioEpsilon).Distinct().ToList();
        var existentes = (await _sql.QueryAsync<string>(
            "SELECT Id_Epsilon FROM dbo.GT_UsuariosPersonal WHERE CodPais = @codPais AND Id_Epsilon IN @ids", new { codPais, ids }, ct))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var g in parseo.Lineas.Where(l => !existentes.Contains(l.UsuarioEpsilon)).GroupBy(l => l.UsuarioEpsilon))
            r.Faltantes.Add(new ElementoFaltante
            {
                Clave = g.Key, Ocurrencias = g.Count(),
                Detalle = $"filas {string.Join(", ", g.Select(l => l.Fila).Take(8))}{(g.Count() > 8 ? "…" : string.Empty)} · {g.Sum(l => l.Importe):N2} €"
            });

        var tipo = "I";
        if (!PaisesTipoInterno.Contains(codPais))
            tipo = await _sql.QueryFirstOrDefaultAsync<string>("SELECT TOP 1 Id FROM dbo.GT_GastosSociedades WHERE CodPais = @codPais", new { codPais }, ct) ?? "I";
        var nombres = parseo.Lineas.Select(l => l.NombreGasto).Distinct().ToList();
        var tiposExistentes = (await _sql.QueryAsync<string>(
            "SELECT NombreGasto FROM dbo.GT_TiposGasto WHERE Tipo = @tipo AND NombreGasto IN @nombres", new { tipo, nombres }, ct)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var n in nombres.Where(n => !tiposExistentes.Contains(n)))
            r.Avisos.Add($"Falta el tipo de gasto '{n}' (tipo '{tipo}') en GT_TiposGasto: la carga fallará hasta darlo de alta.");

        r.Mensaje = r.Faltantes.Count == 0 && r.Avisos.Count == 0
            ? $"Fichero correcto: {r.Filas} fila(s), todos los usuarios Epsilon están asignados. Se puede cargar."
            : $"Fichero con formato correcto ({r.Filas} filas) pero hay {r.Faltantes.Count} usuario(s) Epsilon sin asignar en el sistema de gastos.";
        return r;
    }

    public async Task<ResultadoComprobacion> ComprobarProveedorAsync(OrigenCarga origen, DateTime desde, DateTime hasta, CancellationToken ct = default)
    {
        var r = new ResultadoComprobacion();
        var proveedor = _proveedores.FirstOrDefault(p => p.Origen == origen);
        if (proveedor is null || !proveedor.Configurado) { r.Mensaje = $"La API de {origen.Codigo()} no está configurada."; return r; }

        IReadOnlyList<MovimientoTarjeta> movimientos;
        try { movimientos = await proveedor.ObtenerMovimientosAsync(desde, hasta, ct); }
        catch (Exception ex) when (ex is not OperationCanceledException) { r.Mensaje = $"Error al llamar a la API de {origen.Codigo()}: {ex.Message}"; return r; }

        r.Ok = true;
        r.Filas = movimientos.Count;
        await ComprobarTarjetasAsync(_sql, movimientos, r, ct);
        r.Mensaje = r.Filas == 0 ? "La API no devuelve movimientos en el periodo."
            : r.Faltantes.Count == 0 ? $"{r.Filas} movimiento(s); todas las tarjetas están dadas de alta. Se puede cargar."
            : $"{r.Filas} movimiento(s); hay {r.Faltantes.Count} tarjeta(s) sin dar de alta en el sistema.";
        return r;
    }

    /// <summary>Rellena Faltantes con las tarjetas de los movimientos que no existen en GT_TarjetasSolred (y avisa de las de red comercial).</summary>
    private static async Task ComprobarTarjetasAsync(ISqlComandos sql, IReadOnlyList<MovimientoTarjeta> movimientos, ResultadoComprobacion r, CancellationToken ct)
    {
        if (movimientos.Count == 0) return;
        var numeros = movimientos.Select(m => ClasificadorCarburante.NormalizarTarjeta(m.NumTarjeta)).Distinct().ToList();
        var tarjetas = (await sql.QueryAsync<TarjetaCarburante>(
                "SELECT NumTarjeta, Proveedor, CodPais, Titular, Departamento, Codigo, Jerarquia, Ins_Gasto_Int, Activo FROM dbo.GT_TarjetasSolred WHERE NumTarjeta IN @numeros",
                new { numeros }, ct)).ToDictionary(t => t.NumTarjeta, StringComparer.OrdinalIgnoreCase);
        foreach (var g in movimientos.GroupBy(m => ClasificadorCarburante.NormalizarTarjeta(m.NumTarjeta)))
        {
            if (tarjetas.TryGetValue(g.Key, out var t))
            {
                if (t.Departamento == -1) r.Avisos.Add($"Tarjeta {g.Key} de la red comercial ({g.Count()} mov.): se guardará el detalle sin generar gasto.");
                else if (!string.Equals(t.InsGastoInt?.Trim(), "S", StringComparison.OrdinalIgnoreCase)) r.Avisos.Add($"Tarjeta {g.Key} sin 'insertar gasto interno' ({g.Count()} mov.): solo se guardará el detalle.");
                continue;
            }
            var ej = g.First();
            r.Faltantes.Add(new ElementoFaltante
            {
                Clave = g.Key, Ocurrencias = g.Count(),
                Detalle = $"{ej.Conductor} · {ej.Matricula} · {g.Sum(m => m.ImporteTotal):N2} €"
            });
        }
    }

    // ───────────────────────────── LayTours ─────────────────────────────

    public async Task<ResultadoCarga> CargarLayToursAsync(Stream csv, int codPais, string usuario, string nombreFichero, CancellationToken ct = default)
    {
        ResultadoParseoLayTours parseo;
        using (var lector = new StreamReader(csv, Encoding.Latin1, detectEncodingFromByteOrderMarks: true))
            parseo = LayToursParser.Parsear(lector);

        if (!parseo.Ok)
            return await RegistrarErrorAsync(OrigenCarga.LayTours, usuario, codPais, nombreFichero, null, null, parseo.Lineas.Count, parseo.Error!, ct);

        try
        {
            var resultado = await _sql.EnTransaccionAsync(async tx =>
            {
                var idCarga = await InsertarCargaAsync(tx, OrigenCarga.LayTours, usuario, codPais, nombreFichero, null, null, parseo.Lineas.Count, ct);

                // Tipo de gasto por sociedad: 'I' para las sociedades clásicas; para el resto el Id de GT_GastosSociedades (si existe).
                var tipo = "I";
                if (!PaisesTipoInterno.Contains(codPais))
                    tipo = await tx.QueryFirstOrDefaultAsync<string>("SELECT TOP 1 Id FROM dbo.GT_GastosSociedades WHERE CodPais = @codPais", new { codPais }, ct) ?? "I";

                var tiposGasto = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var empleados = new Dictionary<string, UsuarioPersonal>(StringComparer.OrdinalIgnoreCase);
                var insertados = 0;

                foreach (var l in parseo.Lineas)
                {
                    ct.ThrowIfCancellationRequested();

                    if (!empleados.TryGetValue(l.UsuarioEpsilon, out var emp))
                    {
                        emp = await tx.QueryFirstOrDefaultAsync<UsuarioPersonal>(
                            "SELECT TOP 1 CodPais, Departamento, Codigo, Nombre, Jerarquia, Id_Epsilon FROM dbo.GT_UsuariosPersonal WHERE CodPais = @codPais AND Id_Epsilon = @id",
                            new { codPais, id = l.UsuarioEpsilon }, ct)
                            ?? throw new CargaException($"Se ha encontrado un usuario de Epsilon ({l.UsuarioEpsilon}) sin asignar en el sistema de gastos en la fila {l.Fila}.");
                        empleados[l.UsuarioEpsilon] = emp;
                    }

                    if (!tiposGasto.TryGetValue(l.NombreGasto, out var tipoGasto))
                    {
                        tipoGasto = await tx.QueryFirstOrDefaultAsync<string>(
                            "SELECT TOP 1 Codigo FROM dbo.GT_TiposGasto WHERE Tipo = @tipo AND NombreGasto = @nombre",
                            new { tipo, nombre = l.NombreGasto }, ct)
                            ?? throw new CargaException($"No se ha encontrado el tipo de gasto '{l.NombreGasto}' (tipo '{tipo}') en la fila {l.Fila}.");
                        tiposGasto[l.NombreGasto] = tipoGasto;
                    }

                    await tx.ExecuteAsync(@"
                        INSERT INTO dbo.GT_GastosInternos
                            (CodPais, Codigo, Departamento, Jerarquia, TipoGasto, Fecha, Importe, Numero, Numero2, Donde, Tipo_Com_Inv, VISA,
                             Validado, Validacion, Pagado, Moneda, Factura, FamiliaGasto, Fecha_Insert, Origen, IdCarga)
                        VALUES (@CodPais, @Codigo, @Departamento, @Jerarquia, @TipoGasto, @Fecha, @Importe, @Numero, @Numero2, @Donde, '', 'F',
                                'S', 'Carga Orbis', 'S', 'EUR', @Factura, 'O', SYSDATETIME(), @Origen, @IdCarga)",
                        new
                        {
                            CodPais = codPais, emp.Codigo, emp.Departamento, emp.Jerarquia, TipoGasto = tipoGasto, l.Fecha, l.Importe,
                            Numero = l.EsHotel ? l.NumeroDias : 0, Numero2 = l.EsHotel ? 1 : 0, l.Donde, l.Factura,
                            Origen = OrigenCarga.LayTours.Codigo(), IdCarga = idCarga
                        }, ct);
                    insertados++;
                }

                var res = new ResultadoCarga { Ok = true, IdCarga = idCarga, Filas = parseo.Lineas.Count, Insertados = insertados, Mensaje = "Carga realizada correctamente." };
                await CerrarCargaAsync(tx, res, ct);
                return res;
            }, ct);
            return resultado;
        }
        catch (CargaException ex)
        {
            return await RegistrarErrorAsync(OrigenCarga.LayTours, usuario, codPais, nombreFichero, null, null, parseo.Lineas.Count, ex.Message, ct);
        }
    }

    // ───────────────────────────── Repsol / Moeve ─────────────────────────────

    public async Task<ResultadoCarga> CargarProveedorAsync(OrigenCarga origen, DateTime desde, DateTime hasta, string usuario, CancellationToken ct = default)
    {
        var proveedor = _proveedores.FirstOrDefault(p => p.Origen == origen)
            ?? throw new InvalidOperationException($"No hay proveedor registrado para {origen}.");
        if (!proveedor.Configurado)
            return await RegistrarErrorAsync(origen, usuario, null, null, desde, hasta, 0, $"La API de {origen.Codigo()} no está configurada: {proveedor.Detalle}", ct);

        IReadOnlyList<MovimientoTarjeta> movimientos;
        try
        {
            movimientos = await proveedor.ObtenerMovimientosAsync(desde, hasta, ct);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            _log.LogError(ex, "Error al obtener movimientos de {Origen}", origen);
            return await RegistrarErrorAsync(origen, usuario, null, null, desde, hasta, 0, $"Error al llamar a la API de {origen.Codigo()}: {ex.Message}", ct);
        }
        return await CargarMovimientosAsync(origen, movimientos, usuario, null, desde, hasta, ct);
    }

    public async Task<ResultadoCarga> CargarMovimientosAsync(OrigenCarga origen, IReadOnlyList<MovimientoTarjeta> movimientos, string usuario,
        string? nombreFichero, DateTime? desde, DateTime? hasta, CancellationToken ct = default)
    {
        if (origen == OrigenCarga.LayTours) throw new ArgumentException("LayTours se carga con CargarLayToursAsync.", nameof(origen));
        var tablaDetalle = origen == OrigenCarga.Moeve ? "dbo.GT_GastosMoeve" : "dbo.GT_GastosSolred";
        var validacion = origen == OrigenCarga.Moeve ? "Carga Moeve" : "Carga SolRed";

        if (movimientos.Count == 0)
            return await RegistrarErrorAsync(origen, usuario, null, nombreFichero, desde, hasta, 0, "No hay movimientos que cargar en el periodo indicado.", ct);

        try
        {
            return await _sql.EnTransaccionAsync(async tx =>
            {
                var idCarga = await InsertarCargaAsync(tx, origen, usuario, null, nombreFichero, desde, hasta, movimientos.Count, ct);
                var res = new ResultadoCarga { Ok = true, IdCarga = idCarga, Filas = movimientos.Count };

                // 1) Comprobación previa de tarjetas (equivale al botón "Comprobar tarjetas" del VB6): si falta alguna, se aborta con la lista completa.
                var numeros = movimientos.Select(m => ClasificadorCarburante.NormalizarTarjeta(m.NumTarjeta)).Distinct().ToList();
                var tarjetas = (await tx.QueryAsync<TarjetaCarburante>(
                        "SELECT NumTarjeta, Proveedor, CodPais, Titular, Departamento, Codigo, Jerarquia, Ins_Gasto_Int, Activo FROM dbo.GT_TarjetasSolred WHERE NumTarjeta IN @numeros",
                        new { numeros }, ct))
                    .ToDictionary(t => t.NumTarjeta, StringComparer.OrdinalIgnoreCase);
                var sinAsignar = numeros.Where(n => !tarjetas.ContainsKey(n)).ToList();
                if (sinAsignar.Count > 0)
                    throw new CargaException($"Se han encontrado {sinAsignar.Count} tarjeta(s) sin asignar en el sistema de gastos. No se ha guardado ningún dato. Tarjetas: {string.Join(", ", sinAsignar.Take(50))}{(sinAsignar.Count > 50 ? "…" : string.Empty)}");

                foreach (var m in movimientos)
                {
                    ct.ThrowIfCancellationRequested();
                    var numTarjeta = ClasificadorCarburante.NormalizarTarjeta(m.NumTarjeta);
                    var tarjeta = tarjetas[numTarjeta];

                    // 2) Idempotencia: la API puede devolver movimientos ya cargados en una ejecución anterior; se omiten.
                    var existe = await tx.ExecuteScalarAsync<int>($@"
                        SELECT COUNT(1) FROM {tablaDetalle}
                        WHERE Num_Tarjeta = @numTarjeta AND Factura = @Factura AND Fec_Operacion = @Fecha AND Hor_Operacion = @Hora AND Imp_Total = @Total",
                        new { numTarjeta, m.Factura, Fecha = m.FechaOperacion.Date, Hora = m.HoraOperacion, Total = m.ImporteTotal }, ct);
                    if (existe > 0) { res.Omitidos++; continue; }

                    long? contGasto = null;
                    if (tarjeta.Departamento == -1)
                    {
                        // Red comercial: fuera del alcance del portal (solo gastos internos). Se guarda el detalle sin gasto.
                        res.Avisos.Add($"Tarjeta {numTarjeta} de la red comercial: movimiento guardado sin generar gasto interno.");
                    }
                    else if (string.Equals(tarjeta.InsGastoInt?.Trim(), "S", StringComparison.OrdinalIgnoreCase) && m.ImporteTotal > 0)
                    {
                        var g = ClasificadorCarburante.Clasificar(m);
                        contGasto = await tx.ExecuteScalarAsync<long>(@"
                            INSERT INTO dbo.GT_GastosInternos
                                (CodPais, Codigo, Departamento, Jerarquia, TipoGasto, Fecha, Importe, Numero, Numero2, Donde, Tipo_Com_Inv, VISA,
                                 Validado, Validacion, Pagado, Moneda, FamiliaGasto, Fecha_Insert, Origen, IdCarga)
                            OUTPUT INSERTED.Contador
                            VALUES (@CodPais, @Codigo, @Departamento, @Jerarquia, @TipoGasto, @Fecha, @Importe, @Numero, 0, @Donde, '', 'R',
                                    'S', @Validacion, 'S', 'EUR', 'O', SYSDATETIME(), @Origen, @IdCarga)",
                            new
                            {
                                tarjeta.CodPais, tarjeta.Codigo, tarjeta.Departamento, Jerarquia = tarjeta.Jerarquia.Trim(), g.TipoGasto,
                                Fecha = m.FechaOperacion.Date, Importe = g.ImporteSinIVA, g.Numero, g.Donde, Validacion = validacion,
                                Origen = origen.Codigo(), IdCarga = idCarga
                            }, ct);
                        res.Insertados++;
                    }
                    else
                    {
                        res.Omitidos++;   // importe negativo (descuento) o tarjeta sin "insertar gasto interno"
                    }

                    // 3) Detalle del movimiento (siempre). Km > 9999 se guarda como 0, igual que el VB6.
                    await tx.ExecuteAsync($@"
                        INSERT INTO {tablaDetalle}
                            (Cont_GastoInterno, Num_Tarjeta, Anyo_Fact, Factura, Matricula, Conductor, Fec_Operacion, Hor_Operacion, Poblacion, Km,
                             Des_Prod, Num_Litros, IVA, Imp_Total, Imp_Total_SinIVA, Precio_Litro, Info_Aux, IdCarga, FechaCarga)
                        VALUES (@ContGasto, @NumTarjeta, @Anyo, @Factura, @Matricula, @Conductor, @Fecha, @Hora, @Poblacion, @Km,
                                @DesProd, @Litros, @IVA, @Total, @TotalSinIVA, @PrecioLitro, @InfoAux, @IdCarga, SYSDATETIME())",
                        new
                        {
                            ContGasto = contGasto, NumTarjeta = numTarjeta, Anyo = m.AnyoFactura, m.Factura, m.Matricula, m.Conductor,
                            Fecha = m.FechaOperacion.Date, Hora = m.HoraOperacion, m.Poblacion, Km = m.Km > 9999 ? 0 : m.Km,
                            DesProd = m.DescripcionProducto, Litros = m.Litros, m.IVA, Total = m.ImporteTotal,
                            TotalSinIVA = ClasificadorCarburante.ImporteSinIVA(m.ImporteTotal, m.IVA), m.PrecioLitro,
                            InfoAux = (m.InfoAux ?? string.Empty).Replace('\'', ' '), IdCarga = idCarga
                        }, ct);
                }

                res.Mensaje = $"Carga realizada correctamente: {res.Insertados} gasto(s) generado(s), {res.Omitidos} movimiento(s) sin gasto u omitido(s).";
                await CerrarCargaAsync(tx, res, ct);
                return res;
            }, ct);
        }
        catch (CargaException ex)
        {
            return await RegistrarErrorAsync(origen, usuario, null, nombreFichero, desde, hasta, movimientos.Count, ex.Message, ct);
        }
    }

    // ───────────────────────────── Auditoría (GT_Cargas) ─────────────────────────────

    private static Task<int> InsertarCargaAsync(ISqlComandos sql, OrigenCarga origen, string usuario, int? codPais, string? fichero, DateTime? desde, DateTime? hasta, int filas, CancellationToken ct)
        => sql.ExecuteScalarAsync<int>(@"
            INSERT INTO dbo.GT_Cargas (Origen, Fecha, Usuario, Desde, Hasta, CodPais, NombreFichero, Filas, Insertados, Omitidos, Estado, Mensaje)
            OUTPUT INSERTED.Id
            VALUES (@Origen, SYSDATETIME(), @Usuario, @Desde, @Hasta, @CodPais, @Fichero, @Filas, 0, 0, 'En curso', NULL)",
            new { Origen = origen.Codigo(), Usuario = usuario, Desde = desde, Hasta = hasta, CodPais = codPais, Fichero = fichero, Filas = filas }, ct);

    private static Task<int> CerrarCargaAsync(ISqlComandos sql, ResultadoCarga r, CancellationToken ct)
        => sql.ExecuteAsync("UPDATE dbo.GT_Cargas SET Insertados = @Insertados, Omitidos = @Omitidos, Estado = 'OK', Mensaje = @Mensaje WHERE Id = @IdCarga",
            new { r.Insertados, r.Omitidos, Mensaje = Recortar(r.Mensaje), r.IdCarga }, ct);

    // El error se registra FUERA de la transacción revertida (para que quede constancia de la carga fallida).
    private async Task<ResultadoCarga> RegistrarErrorAsync(OrigenCarga origen, string usuario, int? codPais, string? fichero, DateTime? desde, DateTime? hasta, int filas, string mensaje, CancellationToken ct)
    {
        _log.LogWarning("Carga {Origen} fallida: {Mensaje}", origen, mensaje);
        var res = ResultadoCarga.Error(mensaje, filas);
        try
        {
            res.IdCarga = await _sql.ExecuteScalarAsync<int>(@"
                INSERT INTO dbo.GT_Cargas (Origen, Fecha, Usuario, Desde, Hasta, CodPais, NombreFichero, Filas, Insertados, Omitidos, Estado, Mensaje)
                OUTPUT INSERTED.Id
                VALUES (@Origen, SYSDATETIME(), @Usuario, @Desde, @Hasta, @CodPais, @Fichero, @Filas, 0, 0, 'Error', @Mensaje)",
                new { Origen = origen.Codigo(), Usuario = usuario, Desde = desde, Hasta = hasta, CodPais = codPais, Fichero = fichero, Filas = filas, Mensaje = Recortar(mensaje) }, ct);
        }
        catch (Exception ex) { _log.LogError(ex, "No se pudo registrar la carga fallida en GT_Cargas"); }
        return res;
    }

    private static string Recortar(string s) => s.Length > 1000 ? s[..1000] : s;
}

/// <summary>Error de negocio en una carga (fila incorrecta, tarjeta o usuario desconocido...). Provoca el rollback y se muestra al usuario.</summary>
public sealed class CargaException : Exception
{
    public CargaException(string message) : base(message) { }
}
