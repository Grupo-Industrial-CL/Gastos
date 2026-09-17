# Scripts de BD — Portal de Gastos de viaje

Pensados para **DBeaver** (sin `GO`) contra `MasterPlan_DES` (desarrollo) y después contra la BD de producción.
Son **idempotentes**: se pueden relanzar sin duplicar nada.

| # | Script | Qué hace |
|---|--------|----------|
| 1 | `01_creacion_GT.sql` | Crea todas las tablas `GT_*` (sociedades, departamentos, empleados, tipos de gasto, gastos internos, detalle Repsol/Moeve, tarjetas, roles y sociedades por usuario, auditoría de cargas) y siembra las sociedades del VB6 y los tipos de gasto de carburante. Al final trae, comentado, el `MERGE` para dar de alta al **primer Administrador**: sin él nadie puede entrar. |
| 2 | `02_migracion_informix.sql` | Vuelca el histórico desde Informix (`crislaym`, DSN GALILEO) por servidor vinculado u `OPENQUERY`; alternativa por ficheros CSV documentada dentro. **Pendiente de validar** las columnas reales cuando se tenga acceso a Informix. |

Notas:
- `GT_GastosInternos.CodPais` negativo = borrado lógico (herencia del VB6). El portal solo muestra `CodPais > 0`.
- `GT_TarjetasSolred.Departamento = -1` = tarjeta de la red comercial: sus movimientos se guardan pero no generan gasto interno (fuera del alcance del portal).
- El detalle de movimientos va a `GT_GastosSolred` (Repsol) o `GT_GastosMoeve` (Moeve), con la misma estructura, enlazado al gasto por `Cont_GastoInterno`.
