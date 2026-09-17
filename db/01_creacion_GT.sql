/* =====================================================================================================
   Portal de Gastos de viaje — creación del esquema GT_* en SQL Server (MasterPlan_DES / MasterPlan)
   -----------------------------------------------------------------------------------------------------
   - Pensado para DBeaver (sin GO). IDEMPOTENTE: se puede relanzar; solo crea lo que falta y actualiza semillas.
   - Requiere que exista dbo.MP_Usuarios (directorio de Azure AD compartido con MasterPlan).
   - Equivalencia con Informix (aplicación VB6 "Gastos"):
       paises            -> GT_Sociedades          gastos_empresas -> GT_GastosSociedades
       departamentoscl   -> GT_Departamentos       usuarios_personal -> GT_UsuariosPersonal
       gastos_internos   -> GT_GastosInternos      tiposgasto -> GT_TiposGasto
       gastos_solred     -> GT_GastosSolred (Repsol) y GT_GastosMoeve (Moeve, nueva)
       tarjeta_solred    -> GT_TarjetasSolred
     Nuevas del portal: GT_UsuariosRoles, GT_UsuariosSociedades, GT_Cargas.
   ===================================================================================================== */

-- ───────────────────────────── Sociedades ─────────────────────────────
IF OBJECT_ID('dbo.GT_Sociedades') IS NULL
CREATE TABLE dbo.GT_Sociedades
(
    CodPais       INT            NOT NULL CONSTRAINT PK_GT_Sociedades PRIMARY KEY,
    Nombre        NVARCHAR(100)  NOT NULL,
    NombreIngles  NVARCHAR(100)  NULL,
    Activo        BIT            NOT NULL CONSTRAINT DF_GT_Sociedades_Activo DEFAULT (1)
);

-- Semilla: la lista de sociedades ("países") de Constantes.bas del VB6. No pisa el nombre si ya existe la fila.
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

-- ───────────────────────────── Empresas de gasto ─────────────────────────────
IF OBJECT_ID('dbo.GT_GastosSociedades') IS NULL
CREATE TABLE dbo.GT_GastosSociedades
(
    Id       NVARCHAR(10)   NOT NULL CONSTRAINT PK_GT_GastosSociedades PRIMARY KEY,   -- "Tipo" de GT_TiposGasto para esa sociedad
    CodPais  INT            NOT NULL,
    Empresa  NVARCHAR(100)  NOT NULL,
    Orden    INT            NOT NULL CONSTRAINT DF_GT_GastosSociedades_Orden DEFAULT (0)
);

-- ───────────────────────────── Departamentos ─────────────────────────────
IF OBJECT_ID('dbo.GT_Departamentos') IS NULL
CREATE TABLE dbo.GT_Departamentos
(
    CodPais      INT            NOT NULL,
    Codigo       INT            NOT NULL,
    Descripcion  NVARCHAR(100)  NOT NULL,
    CECO         NVARCHAR(30)   NULL,
    Divisa       NVARCHAR(3)    NULL,
    Activo       BIT            NOT NULL CONSTRAINT DF_GT_Departamentos_Activo DEFAULT (1),
    CONSTRAINT PK_GT_Departamentos PRIMARY KEY (CodPais, Codigo)
);

-- ───────────────────────────── Empleados (usuarios de personal) ─────────────────────────────
IF OBJECT_ID('dbo.GT_UsuariosPersonal') IS NULL
CREATE TABLE dbo.GT_UsuariosPersonal
(
    CodPais          INT            NOT NULL,
    Departamento     INT            NOT NULL,
    Codigo           BIGINT         NOT NULL,
    Nombre           NVARCHAR(100)  NOT NULL,
    DNI              NVARCHAR(20)   NULL,
    Jerarquia        NVARCHAR(10)   NOT NULL CONSTRAINT DF_GT_UsuariosPersonal_Jerarquia DEFAULT (''),
    Id_Epsilon       NVARCHAR(20)   NULL,        -- "User ID" del CSV de LayTours
    Usuario          NVARCHAR(50)   NULL,
    Cuenta_Bancaria  NVARCHAR(34)   NULL,
    FechaBaja        DATE           NULL,
    Bloqueado_Por    NVARCHAR(50)   NULL,
    CONSTRAINT PK_GT_UsuariosPersonal PRIMARY KEY (CodPais, Departamento, Codigo)
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_GT_UsuariosPersonal_Epsilon')
    CREATE INDEX IX_GT_UsuariosPersonal_Epsilon ON dbo.GT_UsuariosPersonal (CodPais, Id_Epsilon);

-- ───────────────────────────── Tipos de gasto ─────────────────────────────
IF OBJECT_ID('dbo.GT_TiposGasto') IS NULL
CREATE TABLE dbo.GT_TiposGasto
(
    Tipo         NVARCHAR(10)  NOT NULL,   -- 'I' (interno estándar) o GT_GastosSociedades.Id
    Codigo       NVARCHAR(5)   NOT NULL,   -- Informix: TipoGasto
    NombreGasto  NVARCHAR(50)  NOT NULL,
    Orden        INT           NOT NULL CONSTRAINT DF_GT_TiposGasto_Orden DEFAULT (0),
    CONSTRAINT PK_GT_TiposGasto PRIMARY KEY (Tipo, Codigo)
);
-- Semilla mínima: los códigos que el VB6 usa de forma fija en la carga de tarjetas. El resto ("Billetes", "Hoteles",
-- "Otros gastos", etc.) se completa con la migración desde Informix (db/02_migracion_informix.sql).
MERGE dbo.GT_TiposGasto AS t
USING (VALUES ('I', '1', N'Gasolina', 1), ('I', '9', N'Peaje', 9), ('I', '10', N'Otros gastos', 10), ('I', '11', N'Parking', 11)) AS s (Tipo, Codigo, NombreGasto, Orden)
   ON t.Tipo = s.Tipo AND t.Codigo = s.Codigo
WHEN NOT MATCHED THEN INSERT (Tipo, Codigo, NombreGasto, Orden) VALUES (s.Tipo, s.Codigo, s.NombreGasto, s.Orden);

-- ───────────────────────────── Auditoría de cargas ─────────────────────────────
IF OBJECT_ID('dbo.GT_Cargas') IS NULL
CREATE TABLE dbo.GT_Cargas
(
    Id             INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_GT_Cargas PRIMARY KEY,
    Origen         NVARCHAR(20)   NOT NULL,   -- LayTours | Repsol | Moeve
    Fecha          DATETIME2(0)   NOT NULL CONSTRAINT DF_GT_Cargas_Fecha DEFAULT (SYSDATETIME()),
    Usuario        NVARCHAR(255)  NOT NULL,   -- correo del administrador que ejecuta la carga
    Desde          DATE           NULL,
    Hasta          DATE           NULL,
    CodPais        INT            NULL,       -- sociedad (LayTours)
    NombreFichero  NVARCHAR(255)  NULL,
    Filas          INT            NOT NULL CONSTRAINT DF_GT_Cargas_Filas DEFAULT (0),
    Insertados     INT            NOT NULL CONSTRAINT DF_GT_Cargas_Insertados DEFAULT (0),
    Omitidos       INT            NOT NULL CONSTRAINT DF_GT_Cargas_Omitidos DEFAULT (0),
    Estado         NVARCHAR(20)   NOT NULL,   -- En curso | OK | Error
    Mensaje        NVARCHAR(1000) NULL
);

-- ───────────────────────────── Gastos internos ─────────────────────────────
IF OBJECT_ID('dbo.GT_GastosInternos') IS NULL
CREATE TABLE dbo.GT_GastosInternos
(
    Contador       BIGINT IDENTITY(1,1) NOT NULL CONSTRAINT PK_GT_GastosInternos PRIMARY KEY,
    CodPais        INT            NOT NULL,   -- negativo = borrado lógico (herencia del VB6: CodPais * -1)
    Codigo         BIGINT         NOT NULL,
    Departamento   INT            NOT NULL,
    Jerarquia      NVARCHAR(10)   NOT NULL CONSTRAINT DF_GT_GastosInternos_Jerarquia DEFAULT (''),
    TipoGasto      NVARCHAR(5)    NOT NULL,
    Fecha          DATE           NOT NULL,
    Importe        DECIMAL(18,2)  NOT NULL,
    Importe_Ori    DECIMAL(18,2)  NULL,
    Moneda         NVARCHAR(3)    NOT NULL CONSTRAINT DF_GT_GastosInternos_Moneda DEFAULT ('EUR'),
    Numero         INT            NOT NULL CONSTRAINT DF_GT_GastosInternos_Numero DEFAULT (0),
    Numero2        INT            NOT NULL CONSTRAINT DF_GT_GastosInternos_Numero2 DEFAULT (0),
    Donde          NVARCHAR(50)   NOT NULL CONSTRAINT DF_GT_GastosInternos_Donde DEFAULT (''),
    Tipo_Com_Inv   NVARCHAR(1)    NULL,
    VISA           NVARCHAR(1)    NULL,       -- F (LayTours) | R (tarjeta) ...
    Validado       NVARCHAR(1)    NULL,
    Validacion     NVARCHAR(30)   NULL,       -- 'Carga Orbis' | 'Carga SolRed' | 'Carga Moeve' | usuario validador
    Pagado         NVARCHAR(1)    NULL,
    Factura        NVARCHAR(30)   NULL,
    FamiliaGasto   NVARCHAR(1)    NULL,
    Proveedor      NVARCHAR(50)   NULL,
    Personas       INT            NULL,
    PEP_CECO       NVARCHAR(30)   NULL,
    Conta_SAP      NVARCHAR(1)    NULL,
    Documento_SAP  NVARCHAR(20)   NULL,
    Archivado_En   NVARCHAR(100)  NULL,
    Fecha_Insert   DATETIME2(0)   NULL,
    Origen         NVARCHAR(20)   NULL,       -- LayTours | Repsol | Moeve | NULL (migrado)
    IdCarga        INT            NULL CONSTRAINT FK_GT_GastosInternos_Carga REFERENCES dbo.GT_Cargas (Id)
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_GT_GastosInternos_PaisFecha')
    CREATE INDEX IX_GT_GastosInternos_PaisFecha ON dbo.GT_GastosInternos (CodPais, Fecha) INCLUDE (Departamento, Codigo, Importe);

-- ───────────────────────────── Tarjetas de carburante ─────────────────────────────
IF OBJECT_ID('dbo.GT_TarjetasSolred') IS NULL
CREATE TABLE dbo.GT_TarjetasSolred
(
    NumTarjeta    NVARCHAR(20)  NOT NULL CONSTRAINT PK_GT_TarjetasSolred PRIMARY KEY,
    Proveedor     NVARCHAR(10)  NOT NULL CONSTRAINT DF_GT_TarjetasSolred_Proveedor DEFAULT ('Repsol'),   -- Repsol | Moeve
    CodPais       INT           NOT NULL,
    Titular       NVARCHAR(100) NULL,
    Empresa       NVARCHAR(100) NULL,
    CIF           NVARCHAR(20)  NULL,
    Departamento  INT           NOT NULL,   -- -1 = red comercial (no genera gasto interno en el portal)
    Codigo        BIGINT        NOT NULL,
    Jerarquia     NVARCHAR(10)  NOT NULL CONSTRAINT DF_GT_TarjetasSolred_Jerarquia DEFAULT (''),
    Ins_Gasto_Int NVARCHAR(1)   NULL,       -- 'S' = sus movimientos generan gasto interno
    Activo        BIT           NOT NULL CONSTRAINT DF_GT_TarjetasSolred_Activo DEFAULT (1)
);

-- ───────────────────────────── Detalle de movimientos (Repsol y Moeve) ─────────────────────────────
IF OBJECT_ID('dbo.GT_GastosSolred') IS NULL
CREATE TABLE dbo.GT_GastosSolred
(
    Id                BIGINT IDENTITY(1,1) NOT NULL CONSTRAINT PK_GT_GastosSolred PRIMARY KEY,
    Cont_GastoInterno BIGINT         NULL CONSTRAINT FK_GT_GastosSolred_Gasto REFERENCES dbo.GT_GastosInternos (Contador),
    Num_Tarjeta       NVARCHAR(20)   NOT NULL,
    Anyo_Fact         INT            NOT NULL,
    Factura           NVARCHAR(30)   NOT NULL,
    Matricula         NVARCHAR(15)   NULL,
    Conductor         NVARCHAR(60)   NULL,
    Fec_Operacion     DATE           NOT NULL,
    Hor_Operacion     NVARCHAR(8)    NOT NULL,
    Poblacion         NVARCHAR(60)   NULL,
    Km                DECIMAL(10,2)  NOT NULL CONSTRAINT DF_GT_GastosSolred_Km DEFAULT (0),
    Des_Prod          NVARCHAR(60)   NULL,
    Num_Litros        DECIMAL(10,3)  NOT NULL CONSTRAINT DF_GT_GastosSolred_Litros DEFAULT (0),
    IVA               DECIMAL(5,2)   NOT NULL CONSTRAINT DF_GT_GastosSolred_IVA DEFAULT (0),
    Imp_Total         DECIMAL(12,2)  NOT NULL,
    Imp_Total_SinIVA  DECIMAL(12,2)  NOT NULL,
    Precio_Litro      DECIMAL(10,4)  NOT NULL CONSTRAINT DF_GT_GastosSolred_PrecioLitro DEFAULT (0),
    Info_Aux          NVARCHAR(200)  NULL,
    IdCarga           INT            NULL CONSTRAINT FK_GT_GastosSolred_Carga REFERENCES dbo.GT_Cargas (Id),
    FechaCarga        DATETIME2(0)   NULL
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_GT_GastosSolred_Dup')
    CREATE INDEX IX_GT_GastosSolred_Dup ON dbo.GT_GastosSolred (Num_Tarjeta, Fec_Operacion, Hor_Operacion, Factura, Imp_Total);

IF OBJECT_ID('dbo.GT_GastosMoeve') IS NULL
CREATE TABLE dbo.GT_GastosMoeve
(
    Id                BIGINT IDENTITY(1,1) NOT NULL CONSTRAINT PK_GT_GastosMoeve PRIMARY KEY,
    Cont_GastoInterno BIGINT         NULL CONSTRAINT FK_GT_GastosMoeve_Gasto REFERENCES dbo.GT_GastosInternos (Contador),
    Num_Tarjeta       NVARCHAR(20)   NOT NULL,
    Anyo_Fact         INT            NOT NULL,
    Factura           NVARCHAR(30)   NOT NULL,
    Matricula         NVARCHAR(15)   NULL,
    Conductor         NVARCHAR(60)   NULL,
    Fec_Operacion     DATE           NOT NULL,
    Hor_Operacion     NVARCHAR(8)    NOT NULL,
    Poblacion         NVARCHAR(60)   NULL,
    Km                DECIMAL(10,2)  NOT NULL CONSTRAINT DF_GT_GastosMoeve_Km DEFAULT (0),
    Des_Prod          NVARCHAR(60)   NULL,
    Num_Litros        DECIMAL(10,3)  NOT NULL CONSTRAINT DF_GT_GastosMoeve_Litros DEFAULT (0),
    IVA               DECIMAL(5,2)   NOT NULL CONSTRAINT DF_GT_GastosMoeve_IVA DEFAULT (0),
    Imp_Total         DECIMAL(12,2)  NOT NULL,
    Imp_Total_SinIVA  DECIMAL(12,2)  NOT NULL,
    Precio_Litro      DECIMAL(10,4)  NOT NULL CONSTRAINT DF_GT_GastosMoeve_PrecioLitro DEFAULT (0),
    Info_Aux          NVARCHAR(200)  NULL,
    IdCarga           INT            NULL CONSTRAINT FK_GT_GastosMoeve_Carga REFERENCES dbo.GT_Cargas (Id),
    FechaCarga        DATETIME2(0)   NULL
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_GT_GastosMoeve_Dup')
    CREATE INDEX IX_GT_GastosMoeve_Dup ON dbo.GT_GastosMoeve (Num_Tarjeta, Fec_Operacion, Hor_Operacion, Factura, Imp_Total);

-- ───────────────────────────── Acceso al portal: roles y sociedades por usuario ─────────────────────────────
IF OBJECT_ID('dbo.GT_UsuariosRoles') IS NULL
CREATE TABLE dbo.GT_UsuariosRoles
(
    Email        NVARCHAR(255) NOT NULL CONSTRAINT PK_GT_UsuariosRoles PRIMARY KEY,   -- = MP_Usuarios.usId
    Rol          NVARCHAR(20)  NOT NULL CONSTRAINT CK_GT_UsuariosRoles_Rol CHECK (Rol IN ('Administrador', 'Consultor')),
    FechaAlta    DATETIME2(0)  NOT NULL CONSTRAINT DF_GT_UsuariosRoles_FechaAlta DEFAULT (SYSDATETIME()),
    UsuarioAlta  NVARCHAR(255) NULL
);

IF OBJECT_ID('dbo.GT_UsuariosSociedades') IS NULL
CREATE TABLE dbo.GT_UsuariosSociedades
(
    Email    NVARCHAR(255) NOT NULL CONSTRAINT FK_GT_UsuariosSociedades_Rol REFERENCES dbo.GT_UsuariosRoles (Email) ON DELETE CASCADE,
    CodPais  INT           NOT NULL CONSTRAINT FK_GT_UsuariosSociedades_Sociedad REFERENCES dbo.GT_Sociedades (CodPais),
    CONSTRAINT PK_GT_UsuariosSociedades PRIMARY KEY (Email, CodPais)
);

/* Primer administrador: sin él nadie puede entrar al portal. Sustituye el correo (usId de MP_Usuarios) y ejecuta:
MERGE dbo.GT_UsuariosRoles AS t USING (SELECT N'itsasor@gicl.es' AS Email) AS s ON t.Email = s.Email
WHEN NOT MATCHED THEN INSERT (Email, Rol, UsuarioAlta) VALUES (s.Email, 'Administrador', 'script inicial');
*/
