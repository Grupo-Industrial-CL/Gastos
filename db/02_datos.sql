/* =====================================================================================================
   Portal de Gastos de viaje — 02_datos.sql: DATOS INICIALES (catálogos y primer administrador)
   -----------------------------------------------------------------------------------------------------
   - Requiere haber ejecutado antes 01_esquema.sql. Pensado para DBeaver (sin GO). IDEMPOTENTE (MERGE).
   - Bloque A: sociedades. La lista sale de Constantes.bas del VB6 y está PENDIENTE DE VALIDAR contra la
     tabla "paises" de Informix cuando se tenga acceso (códigos, nombres y cuáles siguen activas). No pisa
     filas que ya existan.
   - Bloque B: tipos de gasto que el VB6 usa de forma fija en la carga de tarjetas. El resto de tipos
     ("Billetes", "Hoteles", "Otros gastos"...) llega con la migración desde Informix.
   - Bloque C: primer Administrador. Sin él nadie puede entrar en el portal.
   ===================================================================================================== */

-- ───────────────────────────── A) Sociedades (PENDIENTE DE VALIDAR contra Informix "paises") ─────────────────────────────
MERGE dbo.GT_Sociedades AS t
USING (VALUES
    (34, N'España', N'Spain'),
    (35, N'España2', N'Spain2'),
    (51, N'Portugal', N'Portugal'),
    (39, N'Italia', N'Italy'),
    (55, N'Brasil', N'Brazil'),
    (1, N'U.S.A.', N'U.S.A.'),
    (9, N'P.Rico', N'P.Rico'),
    (52, N'Mexico', N'Mexico'),
    (56, N'Chile', N'Chile'),
    (48, N'Polonia', N'Poland'),
    (44, N'Reino Unido', N'United Kingdom'),
    (85, N'Croacia', N'Croatia'),
    (43, N'Austria', N'Austria'),
    (49, N'Alemania', N'Germany'),
    (7, N'Panama', N'Panama'),
    (6, N'Costa Rica', N'Costa Rica'),
    (12, N'Marruecos', N'Marruecos'),
    (42, N'Republica Checa', N'Republica Checa'),
    (58, N'Venezuela', N'Venezuela'),
    (16, N'Tunez', N'Tunez'),
    (54, N'Argentina', N'Argentina'),
    (57, N'Colombia', N'Colombia'),
    (24, N'Angola', N'Angola'),
    (13, N'Argelia', N'Argelia'),
    (23, N'Nigeria', N'Nigeria'),
    (80, N'Hong Kong', N'Hong Kong'),
    (40, N'Rumania', N'Rumania'),
    (90, N'Turquia', N'Turkey'),
    (8, N'Republica Dominicana', N'Republica Dominicana'),
    (59, N'Ecuador', N'Ecuador'),
    (18, N'Libia', N'Libia'),
    (50, N'Peru', N'Peru'),
    (2, N'Guatemala', N'Guatemala'),
    (25, N'Costa de Marfil', N'Costa de Marfil'),
    (10, N'Perseida', N'Perseida'),
    (86, N'Gas Extremadura', N'Gas Extremadura'),
    (87, N'Ondupack', N'Ondupack'),
    (88, N'Alter Enersun', N'Alter Enersun'),
    (89, N'Plastiverd', N'Plastiverd'),
    (91, N'Iqoxe', N'Iqoxe'),
    (92, N'GES', N'GES'),
    (93, N'Lilo Technology', N'Lilo Technology'),
    (83, N'Ondupet', N'Ondupet'),
    (81, N'Murcia Cartón', N'Murcia Cartón'),
    (82, N'Matías Gomá', N'Matías Gomá'),
    (78, N'AG Siderúrgica Balboa', N'AG Siderúrgica Balboa'),
    (77, N'Grupo Gallardo Balboa', N'Grupo Gallardo Balboa'),
    (76, N'Alfonso Gallardo', N'Alfonso Gallardo'),
    (75, N'Corrugados Lasao', N'Corrugados Lasao'),
    (74, N'Marceliano Martín', N'Marceliano Martín'),
    (73, N'Ferromallas', N'Ferromallas'),
    (72, N'Corrugados Getafe', N'Corrugados Getafe'),
    (71, N'Iqlit', N'Iqlit'),
    (70, N'Galvacolor', N'Galvacolor'),
    (69, N'Camal Ecosystems', N'Camal Ecosystems'),
    (68, N'Papeleras Arlanzón', N'Papeleras Arlanzón'),
    (67, N'Fundación RL', N'Fundación RL'),
    (66, N'Gasiluz', N'Gasiluz'),
    (65, N'Cristian LAY DSS', N'Cristian LAY DSS'),
    (64, N'Industrias CL', N'Industrias CL'),
    (63, N'Solar Steel', N'Solar Steel'),
    (62, N'Cartonajes Extremadura', N'Cartonajes Extremadura')
) AS s (CodPais, Nombre, NombreIngles) ON t.CodPais = s.CodPais
WHEN NOT MATCHED THEN INSERT (CodPais, Nombre, NombreIngles, Activo) VALUES (s.CodPais, s.Nombre, s.NombreIngles, 1);

-- ───────────────────────────── B) Tipos de gasto de carburante (códigos fijos del VB6) ─────────────────────────────
MERGE dbo.GT_TiposGasto AS t
USING (VALUES ('I', '1', N'Gasolina', 1), ('I', '9', N'Peaje', 9), ('I', '10', N'Otros gastos', 10), ('I', '11', N'Parking', 11)) AS s (Tipo, Codigo, NombreGasto, Orden)
   ON t.Tipo = s.Tipo AND t.Codigo = s.Codigo
WHEN NOT MATCHED THEN INSERT (Tipo, Codigo, NombreGasto, Orden) VALUES (s.Tipo, s.Codigo, s.NombreGasto, s.Orden);

-- ───────────────────────────── C) Primer Administrador ─────────────────────────────
-- Sustituye el correo (usId de MP_Usuarios) por el del primer administrador y ejecuta:
MERGE dbo.GT_UsuariosRoles AS t USING (SELECT N'itsasor@gicl.es' AS Email) AS s ON t.Email = s.Email
WHEN NOT MATCHED THEN INSERT (Email, Rol, UsuarioAlta) VALUES (s.Email, 'Administrador', 'script inicial');
