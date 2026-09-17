# Scripts de BD — Portal de Gastos de viaje

Pensados para **DBeaver** (sin `GO`) contra `MasterPlan_DES` (desarrollo) y después contra la BD de producción.
Son **idempotentes**: se pueden relanzar sin duplicar nada. Ejecutar en orden.

| # | Script | Qué hace |
|---|--------|----------|
| 1 | `01_esquema.sql` | **Solo estructura.** Crea las tablas `GT_*` (sociedades, empresas de gasto, departamentos, empleados, tipos de gasto, gastos internos, detalle Repsol/Moeve, tarjetas, roles y sociedades por usuario, auditoría de cargas), claves e índices. Sin datos. |
| 2 | `02_datos.sql` | **Datos iniciales.** A) sociedades sacadas de `Constantes.bas` del VB6, **pendientes de validar contra Informix** antes de cargar históricos; B) tipos de gasto de carburante fijos del VB6; C) alta del **primer Administrador** (bloque comentado: poner el correo y ejecutar). |
| 3 | `03_migracion_informix.sql` | Vuelca el histórico desde Informix (`crislaym`, DSN GALILEO) por servidor vinculado u `OPENQUERY`; alternativa por ficheros CSV documentada dentro. **Pendiente de validar** columnas reales cuando se tenga acceso a Informix. |

Notas:
- `GT_GastosInternos.CodPais` negativo = borrado lógico (herencia del VB6). El portal solo muestra `CodPais > 0`.
- `GT_TarjetasSolred.Departamento = -1` = tarjeta de la red comercial: sus movimientos se guardan pero no generan gasto interno (fuera del alcance del portal).
- El detalle de movimientos va a `GT_GastosSolred` (Repsol) o `GT_GastosMoeve` (Moeve), con la misma estructura, enlazado al gasto por `Cont_GastoInterno`.
