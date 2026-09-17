/* =====================================================================================================
   Migración del histórico desde Informix (BD crislaym, DSN GALILEO) a las tablas GT_* de SQL Server
   -----------------------------------------------------------------------------------------------------
   PENDIENTE de ajustar cuando se tenga acceso a Informix (columnas reales y tipos). Dos vías posibles:

   A) SERVIDOR VINCULADO (recomendada si el servidor SQL puede ver Informix por ODBC):
        EXEC sp_addlinkedserver @server = N'GALILEO', @srvproduct = N'Informix', @provider = N'MSDASQL', @datasrc = N'GALILEO';
        EXEC sp_addlinkedsrvlogin @rmtsrvname = N'GALILEO', @useself = 'False', @rmtuser = N'<usuario>', @rmtpassword = N'<clave>';
      y ejecutar los INSERT ... SELECT * FROM OPENQUERY(GALILEO, '...') de abajo.

   B) FICHEROS: exportar cada tabla desde Informix (UNLOAD ... DELIMITER ';') o desde DBeaver a CSV y cargar con
        BULK INSERT dbo.GT_xxx FROM 'C:\migracion\xxx.csv' WITH (FIRSTROW = 2, FIELDTERMINATOR = ';', ROWTERMINATOR = '\n', CODEPAGE = '1252');
      (o con el asistente de importación de DBeaver / SSMS) respetando el orden de tablas de este script.

   Orden: sociedades -> empresas de gasto -> departamentos -> tipos de gasto -> usuarios de personal -> tarjetas
          -> gastos internos -> detalle Solred.  Idempotente: cada bloque solo inserta lo que no existe.
   Los nombres de columna de Informix son los que usa el código VB6; verificar contra el esquema real.
   ===================================================================================================== */

-- 1) Sociedades (paises). La semilla del script 01 ya trae las del VB6; aquí se completan las que falten.
INSERT INTO dbo.GT_Sociedades (CodPais, Nombre, NombreIngles, Activo)
SELECT q.codpais, RTRIM(q.nombre), NULL, 1
FROM OPENQUERY(GALILEO, 'SELECT codpais, nombre FROM paises') q
WHERE NOT EXISTS (SELECT 1 FROM dbo.GT_Sociedades s WHERE s.CodPais = q.codpais);

-- 2) Empresas de gasto (gastos_empresas)
INSERT INTO dbo.GT_GastosSociedades (Id, CodPais, Empresa, Orden)
SELECT RTRIM(q.id), q.codpais, RTRIM(q.empresa), ISNULL(q.orden, 0)
FROM OPENQUERY(GALILEO, 'SELECT id, codpais, empresa, orden FROM gastos_empresas') q
WHERE NOT EXISTS (SELECT 1 FROM dbo.GT_GastosSociedades g WHERE g.Id = RTRIM(q.id));

-- 3) Departamentos (departamentoscl)
INSERT INTO dbo.GT_Departamentos (CodPais, Codigo, Descripcion, CECO, Divisa, Activo)
SELECT q.codpais, q.codigo, RTRIM(q.descripcion), RTRIM(q.ceco), RTRIM(q.divisa), CASE WHEN q.activo = 'S' THEN 1 ELSE 0 END
FROM OPENQUERY(GALILEO, 'SELECT codpais, codigo, descripcion, ceco, divisa, activo FROM departamentoscl') q
WHERE NOT EXISTS (SELECT 1 FROM dbo.GT_Departamentos d WHERE d.CodPais = q.codpais AND d.Codigo = q.codigo);

-- 4) Tipos de gasto (tiposgasto): necesarios para la carga LayTours ("Billetes", "Hoteles", "Otros gastos").
INSERT INTO dbo.GT_TiposGasto (Tipo, Codigo, NombreGasto, Orden)
SELECT RTRIM(q.tipo), RTRIM(CAST(q.tipogasto AS VARCHAR(5))), RTRIM(q.nombregasto), ISNULL(q.idgasto, 0)
FROM OPENQUERY(GALILEO, 'SELECT tipo, tipogasto, nombregasto, idgasto FROM tiposgasto') q
WHERE NOT EXISTS (SELECT 1 FROM dbo.GT_TiposGasto t WHERE t.Tipo = RTRIM(q.tipo) AND t.Codigo = RTRIM(CAST(q.tipogasto AS VARCHAR(5))));

-- 5) Usuarios de personal (usuarios_personal). No se migra la columna Clave (contraseña del sistema antiguo).
INSERT INTO dbo.GT_UsuariosPersonal (CodPais, Departamento, Codigo, Nombre, DNI, Jerarquia, Id_Epsilon, Usuario, Cuenta_Bancaria, FechaBaja, Bloqueado_Por)
SELECT q.codpais, q.departamento, q.codigo, RTRIM(q.nombre), RTRIM(q.dni), ISNULL(RTRIM(q.jerarquia), ''), RTRIM(q.id_epsilon), RTRIM(q.usuario),
       RTRIM(q.cuenta_bancaria), q.fechabaja, RTRIM(q.bloqueado_por)
FROM OPENQUERY(GALILEO, 'SELECT codpais, departamento, codigo, nombre, dni, jerarquia, id_epsilon, usuario, cuenta_bancaria, fechabaja, bloqueado_por FROM usuarios_personal') q
WHERE NOT EXISTS (SELECT 1 FROM dbo.GT_UsuariosPersonal u WHERE u.CodPais = q.codpais AND u.Departamento = q.departamento AND u.Codigo = q.codigo);

-- 6) Tarjetas (tarjeta_solred). Todas las antiguas son de Repsol.
INSERT INTO dbo.GT_TarjetasSolred (NumTarjeta, Proveedor, CodPais, Titular, Empresa, CIF, Departamento, Codigo, Jerarquia, Ins_Gasto_Int, Activo)
SELECT RTRIM(q.numtarjeta), 'Repsol', q.codpais, RTRIM(q.titular), RTRIM(q.empresa), RTRIM(q.cif), q.departamento, q.codigo, ISNULL(RTRIM(q.jerarquia), ''), RTRIM(q.ins_gasto_int), 1
FROM OPENQUERY(GALILEO, 'SELECT numtarjeta, codpais, titular, empresa, cif, departamento, codigo, jerarquia, ins_gasto_int FROM tarjeta_solred') q
WHERE NOT EXISTS (SELECT 1 FROM dbo.GT_TarjetasSolred t WHERE t.NumTarjeta = RTRIM(q.numtarjeta));

-- 7) Gastos internos (gastos_internos). Se conserva el Contador original para poder enlazar el detalle Solred.
--    Origen se deduce de la columna Validacion del VB6 ('Carga Orbis' -> LayTours, 'Carga SolRed' -> Repsol).
SET IDENTITY_INSERT dbo.GT_GastosInternos ON;
INSERT INTO dbo.GT_GastosInternos (Contador, CodPais, Codigo, Departamento, Jerarquia, TipoGasto, Fecha, Importe, Importe_Ori, Moneda, Numero, Numero2, Donde,
                                   Tipo_Com_Inv, VISA, Validado, Validacion, Pagado, Factura, FamiliaGasto, Proveedor, Personas, PEP_CECO, Conta_SAP, Documento_SAP,
                                   Archivado_En, Fecha_Insert, Origen, IdCarga)
SELECT q.contador, q.codpais, q.codigo, q.departamento, ISNULL(RTRIM(q.jerarquia), ''), RTRIM(CAST(q.tipogasto AS VARCHAR(5))), q.fecha, q.importe, q.importe_ori,
       ISNULL(RTRIM(q.moneda), 'EUR'), ISNULL(q.numero, 0), ISNULL(q.numero2, 0), ISNULL(RTRIM(q.donde), ''),
       RTRIM(q.tipo_com_inv), RTRIM(q.visa), RTRIM(q.validado), RTRIM(q.validacion), RTRIM(q.pagado), RTRIM(q.factura), RTRIM(q.familiagasto), RTRIM(q.proveedor),
       q.personas, RTRIM(q.pep_ceco), RTRIM(q.conta_sap), RTRIM(q.documento_sap), RTRIM(q.archivado_en), q.fecha_insert,
       CASE RTRIM(q.validacion) WHEN 'Carga Orbis' THEN 'LayTours' WHEN 'Carga SolRed' THEN 'Repsol' ELSE NULL END, NULL
FROM OPENQUERY(GALILEO, 'SELECT contador, codpais, codigo, departamento, jerarquia, tipogasto, fecha, importe, importe_ori, moneda, numero, numero2, donde,
                                tipo_com_inv, visa, validado, validacion, pagado, factura, familiagasto, proveedor, personas, pep_ceco, conta_sap, documento_sap,
                                archivado_en, fecha_insert FROM gastos_internos') q
WHERE NOT EXISTS (SELECT 1 FROM dbo.GT_GastosInternos g WHERE g.Contador = q.contador);
SET IDENTITY_INSERT dbo.GT_GastosInternos OFF;
-- Reposiciona la secuencia por encima del mayor Contador migrado.
DBCC CHECKIDENT ('dbo.GT_GastosInternos', RESEED);

-- 8) Detalle Solred (gastos_solred). En el VB6 Cont_GastoRed apuntaba a "gastos" (red comercial) o a "gastos_internos";
--    solo se enlaza cuando el contador existe en GT_GastosInternos (el resto queda con Cont_GastoInterno NULL).
INSERT INTO dbo.GT_GastosSolred (Cont_GastoInterno, Num_Tarjeta, Anyo_Fact, Factura, Matricula, Conductor, Fec_Operacion, Hor_Operacion, Poblacion, Km, Des_Prod,
                                 Num_Litros, IVA, Imp_Total, Imp_Total_SinIVA, Precio_Litro, Info_Aux, IdCarga, FechaCarga)
SELECT CASE WHEN g.Contador IS NOT NULL THEN q.cont_gastored END, RTRIM(q.num_tarjeta), q.anyo_fact, RTRIM(q.factura), RTRIM(q.matricula), RTRIM(q.conductor),
       q.fec_operacion, RTRIM(q.hor_operacion), RTRIM(q.poblacion), ISNULL(q.km, 0), RTRIM(q.des_prod), ISNULL(q.num_litros, 0), ISNULL(q.iva, 0),
       q.imp_total, q.imp_total_sinIVA, ISNULL(q.precio_litro, 0), RTRIM(q.info_aux), NULL, NULL
FROM OPENQUERY(GALILEO, 'SELECT cont_gastored, num_tarjeta, anyo_fact, factura, matricula, conductor, fec_operacion, hor_operacion, poblacion, km, des_prod,
                                num_litros, iva, imp_total, imp_total_sinIVA, precio_litro, info_aux FROM gastos_solred') q
LEFT JOIN dbo.GT_GastosInternos g ON g.Contador = q.cont_gastored
WHERE NOT EXISTS (SELECT 1 FROM dbo.GT_GastosSolred s
                  WHERE s.Num_Tarjeta = RTRIM(q.num_tarjeta) AND s.Fec_Operacion = q.fec_operacion AND s.Hor_Operacion = RTRIM(q.hor_operacion)
                    AND s.Factura = RTRIM(q.factura) AND s.Imp_Total = q.imp_total);

-- Comprobación rápida
SELECT 'GT_Sociedades' AS Tabla, COUNT(*) AS Filas FROM dbo.GT_Sociedades
UNION ALL SELECT 'GT_GastosSociedades', COUNT(*) FROM dbo.GT_GastosSociedades
UNION ALL SELECT 'GT_Departamentos', COUNT(*) FROM dbo.GT_Departamentos
UNION ALL SELECT 'GT_TiposGasto', COUNT(*) FROM dbo.GT_TiposGasto
UNION ALL SELECT 'GT_UsuariosPersonal', COUNT(*) FROM dbo.GT_UsuariosPersonal
UNION ALL SELECT 'GT_TarjetasSolred', COUNT(*) FROM dbo.GT_TarjetasSolred
UNION ALL SELECT 'GT_GastosInternos', COUNT(*) FROM dbo.GT_GastosInternos
UNION ALL SELECT 'GT_GastosSolred', COUNT(*) FROM dbo.GT_GastosSolred;
