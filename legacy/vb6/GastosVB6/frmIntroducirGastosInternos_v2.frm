VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIntroducirGastosInternos_v2 
   BorderStyle     =   0  'None
   Caption         =   "Alta / Modificación de gastos y anticipos del personal interno"
   ClientHeight    =   11940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19485
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIntroducirGastosInternos_v2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11940
   ScaleWidth      =   19485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEjecucionRemota 
      Caption         =   "Ejecución remota"
      Height          =   255
      Left            =   17280
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdEncript 
      BackColor       =   &H008080FF&
      Caption         =   "Encr"
      Height          =   375
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   600
      Width           =   1335
   End
   Begin VB.Frame fraHistorico 
      Height          =   9975
      Left            =   19080
      TabIndex        =   83
      Top             =   11760
      Visible         =   0   'False
      Width           =   19320
      Begin VB.CommandButton cmdCerrarHistorico 
         Caption         =   "Cerrar"
         Height          =   735
         Left            =   6360
         Picture         =   "frmIntroducirGastosInternos_v2.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   5880
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid mfgHistoricoModificaciones 
         Height          =   4215
         Left            =   120
         TabIndex        =   84
         Top             =   960
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   7435
         _Version        =   393216
         AllowUserResizing=   1
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Histórico de modificaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   29
         Left            =   120
         TabIndex        =   85
         Top             =   600
         Width           =   2685
      End
   End
   Begin VB.CommandButton cmdLogonGESDesarrollo 
      BackColor       =   &H008080FF&
      Caption         =   "Logon GES Desarrollo"
      Height          =   375
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeslogon 
      BackColor       =   &H008080FF&
      Caption         =   "DesLogon"
      Height          =   375
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraPersonal 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   240
      TabIndex        =   36
      Top             =   720
      Width           =   8415
      Begin VB.CheckBox chkTodosPreValidacion 
         Caption         =   "Usuarios con pendientes prevalidar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CheckBox chkBajas 
         Caption         =   "Listar usuarios baja"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbEmpleados 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   5775
      End
      Begin VB.ComboBox cmbDepartamentos 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   200
         Width           =   5775
      End
      Begin VB.Image imgPassword 
         Height          =   480
         Left            =   7920
         Picture         =   "frmIntroducirGastosInternos_v2.frx":1434
         ToolTipText     =   "Consultar password empleado"
         Top             =   960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblEmpleadoSinNumAcreedor 
         Caption         =   "EMPLEADO SIN NÚMERO DE ACREEDOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3840
         TabIndex        =   107
         Top             =   1200
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Image imgAceptarEmpleado 
         Height          =   480
         Left            =   7920
         Picture         =   "frmIntroducirGastosInternos_v2.frx":1CFE
         ToolTipText     =   "Introducir gasto del empleado seleccionado"
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Empleado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   38
         Top             =   765
         Width           =   930
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame fraMensaje 
      Height          =   9975
      Left            =   120
      TabIndex        =   103
      Top             =   11760
      Visible         =   0   'False
      Width           =   19320
      Begin VB.Label lblMensaje3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mensaje3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   480
         TabIndex        =   106
         Top             =   4320
         Width           =   18135
      End
      Begin VB.Label lblMensaje2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mensaje2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   480
         TabIndex        =   105
         Top             =   3840
         Width           =   18135
      End
      Begin VB.Label lblMensaje1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mensaje1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   104
         Top             =   2880
         Width           =   18135
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   2895
         Left            =   360
         Top             =   2400
         Width           =   18495
      End
   End
   Begin VB.CommandButton cmdContaAnticipo 
      BackColor       =   &H008080FF&
      Caption         =   "Contabilizar Anticipo"
      Height          =   375
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogonGESCalidad 
      BackColor       =   &H008080FF&
      Caption         =   "Logon GES Calidad"
      Height          =   375
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogonGESProduccion 
      BackColor       =   &H008080FF&
      Caption         =   "Logon GES Producción"
      Height          =   375
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraDatos 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9975
      Left            =   120
      TabIndex        =   39
      Top             =   1920
      Width           =   19320
      Begin VB.TextBox txtTipoIVA2 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   14760
         MaxLength       =   2
         TabIndex        =   32
         Top             =   7200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtBaseIVA2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   12240
         TabIndex        =   31
         Top             =   7200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtImporteIVA2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   9720
         TabIndex        =   30
         Top             =   7200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtTipoIVA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   14760
         MaxLength       =   2
         TabIndex        =   29
         Top             =   6915
         Width           =   615
      End
      Begin VB.TextBox txtBaseIVA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   12240
         TabIndex        =   28
         Top             =   6915
         Width           =   1095
      End
      Begin VB.CheckBox chkDesgloseIVA 
         Caption         =   "Desglose IVA"
         Height          =   195
         Left            =   6480
         TabIndex        =   26
         Top             =   6960
         Width           =   1575
      End
      Begin VB.TextBox txtImporteIVA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9720
         TabIndex        =   27
         Top             =   6915
         Width           =   1095
      End
      Begin VB.CheckBox chkVerTicketAutomaticamente 
         Height          =   495
         Left            =   18000
         Picture         =   "frmIntroducirGastosInternos_v2.frx":25C8
         Style           =   1  'Graphical
         TabIndex        =   112
         TabStop         =   0   'False
         ToolTipText     =   "Ver ticket automáticamente"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtFactura 
         Height          =   285
         Left            =   11640
         TabIndex        =   19
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox txtProveedor 
         Height          =   285
         Left            =   10320
         TabIndex        =   18
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "Rechazar gasto"
         Height          =   735
         Left            =   17520
         Picture         =   "frmIntroducirGastosInternos_v2.frx":2CB2
         Style           =   1  'Graphical
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   4560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkPedirFechaContaSAP 
         Caption         =   "Pedir fecha contabilización"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9360
         TabIndex        =   96
         Top             =   8400
         Width           =   2775
      End
      Begin VB.Frame etAyudaConceptoGasto 
         Height          =   1905
         Left            =   360
         TabIndex        =   60
         Top             =   7920
         Visible         =   0   'False
         Width           =   4815
         Begin VB.Label Label4 
            Caption         =   " 8->Taxis          9->Peajes  10->Otros gastos 11->Parking 12->Muestras     13->Formación"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1260
            Index           =   1
            Left            =   2760
            TabIndex        =   62
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label4 
            Caption         =   $"frmIntroducirGastosInternos_v2.frx":357C
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1500
            Index           =   0
            Left            =   135
            TabIndex        =   61
            Top             =   225
            Width           =   2280
         End
      End
      Begin VB.ComboBox cmbCuentasContables 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7560
         TabIndex        =   25
         Top             =   6465
         Width           =   6135
      End
      Begin VB.TextBox txtCIF_Proveedor 
         Height          =   285
         Left            =   7560
         TabIndex        =   24
         Top             =   6080
         Width           =   2175
      End
      Begin VB.Frame etAyudaTipoPago 
         Height          =   1185
         Left            =   360
         TabIndex        =   67
         Top             =   7920
         Visible         =   0   'False
         Width           =   2175
         Begin VB.Label Label4 
            Caption         =   "1 ó N->Metálico    2 ó S->VISA       3 ó F->Facturado   4 ó R->SolRed"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   900
            Index           =   2
            Left            =   135
            TabIndex        =   68
            Top             =   225
            Width           =   1920
         End
      End
      Begin VB.TextBox txtPEPCECO 
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   6465
         Width           =   2175
      End
      Begin VB.TextBox txtCambio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         TabIndex        =   23
         Top             =   6795
         Width           =   855
      End
      Begin VB.ComboBox cmbDivisa 
         Height          =   315
         ItemData        =   "frmIntroducirGastosInternos_v2.frx":3605
         Left            =   1440
         List            =   "frmIntroducirGastosInternos_v2.frx":3607
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   6795
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid mfgTiposGasto 
         Height          =   1815
         Left            =   360
         TabIndex        =   73
         Top             =   7920
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3201
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         ForeColor       =   255
         BackColorFixed  =   -2147483643
         ForeColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin VB.ComboBox cmbEmpresas 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   6075
         Width           =   3495
      End
      Begin VB.TextBox txtDonde 
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   4800
         Width           =   3855
      End
      Begin VB.Frame etAyudaFamiliaGasto 
         Height          =   945
         Left            =   360
         TabIndex        =   65
         Top             =   7920
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label Label4 
            Caption         =   "1 ó O->Ordinario    2 ó F->Formación     3 ó P->Presentación"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   660
            Index           =   4
            Left            =   135
            TabIndex        =   66
            Top             =   225
            Width           =   2160
         End
      End
      Begin VB.CommandButton cmdBorrarLinea 
         Height          =   495
         Left            =   12360
         MaskColor       =   &H8000000F&
         Picture         =   "frmIntroducirGastosInternos_v2.frx":3609
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Eliminar línea"
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton cmdSubirLinea 
         Height          =   495
         Left            =   11640
         Picture         =   "frmIntroducirGastosInternos_v2.frx":3B15
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Guardar línea"
         Top             =   4560
         Width           =   495
      End
      Begin VB.TextBox txtPersonas 
         Height          =   285
         Left            =   5760
         TabIndex        =   17
         Top             =   5520
         Width           =   4575
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8160
         TabIndex        =   11
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtNumPersonas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4680
         TabIndex        =   16
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox txtNumTickets 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   15
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox txtKms 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Top             =   5520
         Width           =   855
      End
      Begin VB.TextBox txtNumNoches 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox txtTipoComida 
         Height          =   285
         Left            =   360
         TabIndex        =   12
         Top             =   5520
         Width           =   1335
      End
      Begin VB.TextBox txtTipoPago 
         Height          =   285
         Left            =   3360
         TabIndex        =   9
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtFamiliaGasto 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtConceptoGasto 
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   4800
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   345
         Left            =   3000
         TabIndex        =   3
         Top             =   510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yy"
         Format          =   110952451
         CurrentDate     =   42384
      End
      Begin MSFlexGridLib.MSFlexGrid mfgParrilla 
         Height          =   3255
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   19140
         _ExtentX        =   33761
         _ExtentY        =   5741
         _Version        =   393216
         AllowUserResizing=   1
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   345
         Left            =   3000
         TabIndex        =   4
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yy"
         Format          =   110952451
         CurrentDate     =   42384
      End
      Begin VB.Frame etAyudaTipoComida 
         Height          =   825
         Left            =   360
         TabIndex        =   69
         Top             =   8040
         Visible         =   0   'False
         Width           =   1815
         Begin VB.Label Label4 
            Caption         =   "1 ó N->Negocio 2 ó I->Interno"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   420
            Index           =   5
            Left            =   135
            TabIndex        =   70
            Top             =   225
            Width           =   1560
         End
      End
      Begin VB.Frame etFechaAyuda 
         Height          =   705
         Left            =   360
         TabIndex        =   63
         Top             =   7920
         Visible         =   0   'False
         Width           =   3735
         Begin VB.Label Label4 
            Caption         =   "Fecha en Formato ""ddmmaaaa"""
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   3
            Left            =   135
            TabIndex        =   64
            Top             =   225
            Width           =   2880
         End
      End
      Begin VB.Image imgCancelarContabilizarSAP 
         Height          =   360
         Left            =   10440
         Picture         =   "frmIntroducirGastosInternos_v2.frx":43DF
         ToolTipText     =   "Cancelar contabilización en SAP"
         Top             =   7920
         Width           =   720
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "Tipo IVA"
         Height          =   195
         Index           =   41
         Left            =   13800
         TabIndex        =   118
         Top             =   7245
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "Base IVA"
         Height          =   195
         Index           =   40
         Left            =   11280
         TabIndex        =   117
         Top             =   7245
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "Importe IVA "
         Height          =   195
         Index           =   39
         Left            =   8520
         TabIndex        =   116
         Top             =   7245
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo IVA"
         Height          =   195
         Index           =   38
         Left            =   13800
         TabIndex        =   115
         Top             =   6960
         Width           =   735
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Base IVA"
         Height          =   195
         Index           =   37
         Left            =   11280
         TabIndex        =   114
         Top             =   6960
         Width           =   795
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Importe IVA"
         Height          =   195
         Index           =   36
         Left            =   8520
         TabIndex        =   113
         Top             =   6960
         Width           =   1065
      End
      Begin VB.Image imgExportar 
         Height          =   480
         Left            =   18600
         Picture         =   "frmIntroducirGastosInternos_v2.frx":4D69
         ToolTipText     =   "Copiar al Portapapeles"
         Top             =   720
         Width           =   480
      End
      Begin VB.Shape shCambioUsuario 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   255
         Left            =   3150
         Top             =   6825
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Image imgSimuladorCambios 
         Height          =   360
         Left            =   5040
         Picture         =   "frmIntroducirGastosInternos_v2.frx":5633
         ToolTipText     =   "Simulador cambio"
         Top             =   6720
         Width           =   360
      End
      Begin VB.Label lblCPD_CIF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10800
         TabIndex        =   101
         Top             =   6120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblCPD_Pais 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   100
         Top             =   6120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblCPD_CP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10440
         TabIndex        =   99
         Top             =   6120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblCPD_Ciudad 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10560
         TabIndex        =   98
         Top             =   6120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblCPD_Nombre 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10320
         TabIndex        =   97
         Top             =   6120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imgDatos_ProveedorCPD 
         Height          =   360
         Left            =   9840
         Picture         =   "frmIntroducirGastosInternos_v2.frx":5D1D
         ToolTipText     =   "Datos vendor CPD"
         Top             =   6015
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblValidado 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   95
         Top             =   5880
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Validado y NO pagado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   33
         Left            =   17760
         TabIndex        =   92
         ToolTipText     =   "Filtrar por validados y no pagados"
         Top             =   0
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Contable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   31
         Left            =   6480
         TabIndex        =   91
         Top             =   6525
         Width           =   1005
      End
      Begin VB.Label lblFichero 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   14400
         TabIndex        =   89
         Top             =   4560
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imgAdjuntar 
         Height          =   480
         Left            =   13920
         Picture         =   "frmIntroducirGastosInternos_v2.frx":6407
         ToolTipText     =   "Adjuntar justificante"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "IF Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   30
         Left            =   6480
         TabIndex        =   87
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Image imgModificado 
         Height          =   240
         Left            =   0
         Picture         =   "frmIntroducirGastosInternos_v2.frx":6CD1
         ToolTipText     =   "Modifcado desde Administración, click para ver historial de cambios"
         Top             =   4800
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblContabilizadoSAP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   5640
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblPagado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   80
         Top             =   5640
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblImporteEUR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   79
         Top             =   7155
         Width           =   1335
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Imp.cambio"
         Height          =   195
         Index           =   27
         Left            =   360
         TabIndex        =   78
         Top             =   7200
         Width           =   1020
      End
      Begin VB.Image imgContabilizarSAP 
         Height          =   360
         Left            =   9360
         Picture         =   "frmIntroducirGastosInternos_v2.frx":725B
         ToolTipText     =   "Contabilizar en SAP"
         Top             =   7920
         Width           =   720
      End
      Begin VB.Image imgArchivarEn 
         Height          =   720
         Left            =   8520
         Picture         =   "frmIntroducirGastosInternos_v2.frx":7BE5
         ToolTipText     =   "Justificantes archivados en"
         Top             =   7800
         Width           =   720
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cambio"
         Height          =   195
         Index           =   25
         Left            =   3240
         TabIndex        =   76
         Top             =   6840
         Width           =   660
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Divisa"
         Height          =   195
         Index           =   24
         Left            =   360
         TabIndex        =   75
         Top             =   6840
         Width           =   885
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "PEP/CECO/IO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   23
         Left            =   360
         TabIndex        =   74
         Top             =   6525
         Width           =   960
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Index           =   22
         Left            =   360
         TabIndex        =   72
         Top             =   6120
         Width           =   1110
      End
      Begin VB.Image imgDesmarcarTodas 
         Height          =   360
         Left            =   600
         Picture         =   "frmIntroducirGastosInternos_v2.frx":8AAF
         ToolTipText     =   "Desmarcar todas"
         Top             =   840
         Width           =   360
      End
      Begin VB.Image imgMarcarTodas 
         Height          =   360
         Left            =   240
         Picture         =   "frmIntroducirGastosInternos_v2.frx":9199
         ToolTipText     =   "Marcar todas"
         Top             =   840
         Width           =   360
      End
      Begin VB.Image imgImprimirPagos 
         Height          =   720
         Left            =   7560
         Picture         =   "frmIntroducirGastosInternos_v2.frx":9883
         ToolTipText     =   "Imprimir gastos"
         Top             =   7800
         Width           =   720
      End
      Begin VB.Image imgPagar 
         Height          =   720
         Left            =   6480
         Picture         =   "frmIntroducirGastosInternos_v2.frx":A74D
         ToolTipText     =   "Pagar gastos seleccionados"
         Top             =   7800
         Width           =   720
      End
      Begin VB.Image imgValidar 
         Height          =   720
         Left            =   5520
         Picture         =   "frmIntroducirGastosInternos_v2.frx":B617
         ToolTipText     =   "Validar gastos seleccionados"
         Top             =   7800
         Width           =   720
      End
      Begin VB.Image imgConsultar 
         Height          =   480
         Left            =   13320
         Picture         =   "frmIntroducirGastosInternos_v2.frx":C4E1
         ToolTipText     =   "Ver ticket asociado al gasto"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Personas"
         Height          =   195
         Index           =   21
         Left            =   5760
         TabIndex        =   71
         Top             =   5280
         Width           =   1140
      End
      Begin VB.Label lblRuta 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   5400
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblContador 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   58
         Top             =   5400
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imgBorrarLinea 
         Height          =   360
         Left            =   12360
         Picture         =   "frmIntroducirGastosInternos_v2.frx":CDAB
         ToolTipText     =   "Eliminar línea"
         Top             =   4605
         Width           =   360
      End
      Begin VB.Image imgAceptarLinea 
         Height          =   480
         Left            =   11640
         Picture         =   "frmIntroducirGastosInternos_v2.frx":D2B7
         ToolTipText     =   "Guardar línea"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         Height          =   195
         Index           =   20
         Left            =   11640
         TabIndex        =   57
         Top             =   5280
         Width           =   990
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   19
         Left            =   10320
         TabIndex        =   56
         Top             =   5280
         Width           =   1245
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Index           =   18
         Left            =   8160
         TabIndex        =   55
         Top             =   4560
         Width           =   1050
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "NºPersonas"
         Height          =   195
         Index           =   17
         Left            =   4680
         TabIndex        =   54
         Top             =   5280
         Width           =   1350
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nº Tickets"
         Height          =   195
         Index           =   16
         Left            =   3600
         TabIndex        =   53
         Top             =   5280
         Width           =   1230
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Kms"
         Height          =   195
         Index           =   15
         Left            =   2760
         TabIndex        =   52
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nº Noches"
         Height          =   195
         Index           =   14
         Left            =   1680
         TabIndex        =   51
         Top             =   5280
         Width           =   885
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "T. Comida"
         Height          =   195
         Index           =   12
         Left            =   360
         TabIndex        =   50
         Top             =   5280
         Width           =   885
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Donde / Itinerario / Concepto"
         Height          =   195
         Index           =   13
         Left            =   4320
         TabIndex        =   49
         Top             =   4560
         Width           =   2550
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "T. Pago"
         Height          =   195
         Index           =   11
         Left            =   3360
         TabIndex        =   48
         Top             =   4560
         Width           =   645
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "T. Gasto"
         Height          =   195
         Index           =   10
         Left            =   2400
         TabIndex        =   47
         Top             =   4560
         Width           =   720
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Index           =   9
         Left            =   1320
         TabIndex        =   46
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "C. Gasto"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   45
         Top             =   4560
         Width           =   750
      End
      Begin VB.Image imgConsultaGastosFecha 
         Height          =   480
         Left            =   4560
         Picture         =   "frmIntroducirGastosInternos_v2.frx":DB81
         ToolTipText     =   "Introducir/Consultar gastos del empleado seleccionado"
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "al"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   1320
         TabIndex        =   41
         Top             =   960
         Width           =   165
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Buscar desde el "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   1320
         TabIndex        =   40
         Top             =   600
         Width           =   1560
      End
      Begin VB.Image imgPreValidar 
         Height          =   720
         Left            =   5640
         Picture         =   "frmIntroducirGastosInternos_v2.frx":E44B
         ToolTipText     =   "Prevalidar gastos seleccionados"
         Top             =   7800
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblTexto 
         Alignment       =   2  'Center
         Caption         =   "Puede filtrar los resultados de la búsqueda picando en el estado que desea visualizar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   28
         Left            =   11400
         TabIndex        =   82
         ToolTipText     =   "Filtrar por contabilizados SAP"
         Top             =   0
         Width           =   6195
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C0FFC0&
         Height          =   135
         Left            =   11280
         Top             =   600
         Width           =   135
      End
      Begin VB.Label lblTexto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Pte. prevalidar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   42
         Left            =   11490
         TabIndex        =   120
         ToolTipText     =   "Filtrar por pendientes de prevalidar"
         Top             =   585
         Width           =   1110
      End
      Begin VB.Shape shFiltro_PendientePreValidar 
         BorderColor     =   &H000000FF&
         Height          =   255
         Left            =   11160
         Top             =   540
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C0FFC0&
         Height          =   135
         Left            =   11280
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Solo validado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   34
         Left            =   11520
         TabIndex        =   108
         ToolTipText     =   "Filtrar por validados"
         Top             =   945
         Width           =   1005
      End
      Begin VB.Shape shFiltro_SoloValidado 
         BorderColor     =   &H000000FF&
         Height          =   255
         Left            =   11160
         Top             =   900
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Validado y pagado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   35
         Left            =   17760
         TabIndex        =   93
         ToolTipText     =   "Filtrar por validados y pagados"
         Top             =   240
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C0FFC0&
         Height          =   135
         Left            =   14640
         Top             =   600
         Width           =   135
      End
      Begin VB.Label lblTexto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Rechazado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   32
         Left            =   14865
         TabIndex        =   88
         ToolTipText     =   "Filtrar por pendientes de validar"
         Top             =   585
         Width           =   810
      End
      Begin VB.Shape shFiltro_Rechazado 
         BorderColor     =   &H000000FF&
         Height          =   255
         Left            =   14520
         Top             =   540
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shFiltro_ContabilizadoSAP 
         BorderColor     =   &H000000FF&
         Height          =   255
         Left            =   14520
         Top             =   900
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shFiltro_ValidadoPagado 
         BorderColor     =   &H000000FF&
         Height          =   255
         Left            =   15960
         Top             =   900
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Shape shFiltro_ValidadoNoPagado 
         BorderColor     =   &H000000FF&
         Height          =   255
         Left            =   12840
         Top             =   900
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Shape shFiltro_PendienteValidar 
         BorderColor     =   &H000000FF&
         Height          =   255
         Left            =   12840
         Top             =   540
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C0FFC0&
         Height          =   135
         Index           =   1
         Left            =   14640
         Top             =   975
         Width           =   135
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Conta. SAP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   26
         Left            =   14880
         TabIndex        =   77
         ToolTipText     =   "Filtrar por contabilizados SAP"
         Top             =   945
         Width           =   870
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Validado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   13200
         TabIndex        =   44
         ToolTipText     =   "Filtrar por validados"
         Top             =   945
         Width           =   660
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Pagado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   16320
         TabIndex        =   43
         ToolTipText     =   "Filtrar por pagados"
         Top             =   945
         Width           =   555
      End
      Begin VB.Label lblTexto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Pndte. validar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   13200
         TabIndex        =   42
         ToolTipText     =   "Filtrar por pendientes de validar"
         Top             =   585
         Width           =   1050
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C0FFC0&
         Height          =   135
         Left            =   12960
         Top             =   600
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C0FFC0&
         Height          =   135
         Index           =   0
         Left            =   16080
         Top             =   960
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C0FFC0&
         Height          =   135
         Left            =   12960
         Top             =   960
         Width           =   135
      End
   End
   Begin VB.Image imgInforme 
      Height          =   720
      Left            =   12240
      Picture         =   "frmIntroducirGastosInternos_v2.frx":F315
      ToolTipText     =   "Ir a pantalla de informe"
      Top             =   960
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuscarUsuariosConGastosValidados 
      Height          =   720
      Left            =   13320
      Picture         =   "frmIntroducirGastosInternos_v2.frx":101DF
      ToolTipText     =   "Buscar empleados"
      Top             =   960
      Width           =   720
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   150
   End
   Begin VB.Image imgNuevoEmpleado 
      Height          =   480
      Left            =   9000
      Picture         =   "frmIntroducirGastosInternos_v2.frx":110A9
      ToolTipText     =   "Nuevo empleado"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmIntroducirGastosInternos_v2.frx":116C5
      ToolTipText     =   "Pica y mueve el formulario"
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   12000
      Picture         =   "frmIntroducirGastosInternos_v2.frx":1238F
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Alta / Modificación de gastos y anticipos del personal interno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   8
      Left            =   720
      TabIndex        =   35
      Top             =   75
      Width           =   10920
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   18840
      Picture         =   "frmIntroducirGastosInternos_v2.frx":12699
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E2C07A&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   19500
   End
End
Attribute VB_Name = "frmIntroducirGastosInternos_v2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ZbapiContabilizarAnticipos

Dim mafoRegistro As ADODB.Recordset
Dim mafoRegEmpleados As ADODB.Recordset
Dim mafoRegCuentasContables As ADODB.Recordset

Dim mbEmpleados As Boolean
Dim mbCuentasContables As Boolean

Dim moParrilla As New cFlexGrid2
Dim moParrillaHistorico As New CFlexGrid
Dim moTiposGasto As New CFlexGrid

Dim miCodDepartamento As Integer
Dim miCodEmpleado As Long
Dim msJerEmpleado As String
Dim msAcreedorEmpleado As String
Dim msMonedaDepartamento As String
Dim msDivision_PersonalEmpleado As String
Dim msClave_OrganizacionEmpleado As String
Dim miManagerEmpleado As Long
Dim miOfi_PreValidacionEmpleado As Integer

Dim msImporteKms As String

Dim msTipo As String

Dim msFiltroTipoPago As String

Dim mbConexionSAP As Boolean
Dim msFechaContaSAP As String

Dim mid_Ofi_PreValidacion As Integer

Dim miColumnaOrdenacion As Integer
Dim miPosicionarEnFila As Long

Dim cMouseW As New cMouseWheel

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub chkBajas_Click()
    cmbDepartamentos_Click
End Sub

Private Sub chkDesgloseIVA_Click()

    If chkDesgloseIVA.Tag <> "" Then Exit Sub
    
    If chkDesgloseIVA.value = vbChecked Then
        txtImporteIVA.Enabled = True
        txtBaseIVA.Enabled = True
        txtTipoIVA.Enabled = True
'        txtImporteIVA2.Enabled = True
'        txtBaseIVA2.Enabled = True
'        txtTipoIVA2.Enabled = True
        txtImporteIVA.SetFocus
    Else
        txtImporteIVA.Text = ""
        txtImporteIVA.Enabled = False
        txtBaseIVA.Text = ""
        txtBaseIVA.Enabled = False
        txtTipoIVA.Text = ""
        txtTipoIVA.Enabled = False
'        txtImporteIVA2.Text = ""
'        txtImporteIVA2.Enabled = False
'        txtBaseIVA2.Text = ""
'        txtBaseIVA2.Enabled = False
'        txtTipoIVA2.Text = ""
'        txtTipoIVA2.Enabled = False
    End If

End Sub

Private Sub chkTodosPreValidacion_Click()
    cmbDepartamentos_Click
End Sub

Private Sub cmbCuentasContables_Change()

    If mbCuentasContables = False Then Exit Sub

    Dim iCont As Integer
    
    For iCont = 0 To cmbCuentasContables.ListCount - 1
        cmbCuentasContables.RemoveItem 0
    Next iCont
    
    If cmbCuentasContables.Text = "" Then
    
        While Not mafoRegCuentasContables.EOF
            cmbCuentasContables.AddItem Trim(mafoRegCuentasContables("Cuenta_Contable")) & " - " & Trim(mafoRegCuentasContables("DescCuenta_Contable"))
            mafoRegCuentasContables.MoveNext
        Wend
    
    Else
    
        While Not mafoRegCuentasContables.EOF
            If InStr(1, UCase(Trim(mafoRegCuentasContables("Cuenta_Contable"))), UCase(cmbCuentasContables.Text)) Then cmbCuentasContables.AddItem Trim(mafoRegCuentasContables("Cuenta_Contable")) & " - " & Trim(mafoRegCuentasContables("DescCuenta_Contable"))
            mafoRegCuentasContables.MoveNext
        Wend
    
    End If
    mafoRegCuentasContables.MoveFirst

End Sub

Private Sub cmbDepartamentos_Click()

    Dim sAux As String
    Dim sBajas As String
    Dim sPendientesPreValidar As String

    If cmbDepartamentos.Tag = "C" Then Exit Sub
    VaciarBox cmbEmpleados
    If cmbDepartamentos.Text = "" Then Exit Sub
    
    sBajas = "N"
    If chkBajas.value = vbChecked Then sBajas = "S"
    sPendientesPreValidar = "N"
    If mid_Ofi_PreValidacion > 0 And chkTodosPreValidacion.value = vbChecked Then sPendientesPreValidar = "S"
    
    Set mafoRegEmpleados = Nothing
    mbEmpleados = False
    If bDameEmpleadosDepartamento(CodPaisTrabajo, cmbDepartamentos.ItemData(cmbDepartamentos.ListIndex), mafoRegEmpleados, sBajas, mid_Ofi_PreValidacion, sPendientesPreValidar) Then
        mbEmpleados = True
        While Not mafoRegEmpleados.EOF
            sAux = Trim(mafoRegEmpleados("Nombre"))
            If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegEmpleados("Codigo") & ") " & sAux
            cmbEmpleados.AddItem sAux
            cmbEmpleados.ItemData(cmbEmpleados.NewIndex) = mafoRegEmpleados("Codigo")
            mafoRegEmpleados.MoveNext
        Wend
        mafoRegEmpleados.MoveFirst
    End If
    'Set mafoRegistro = Nothing
    
End Sub

Private Sub cmbDepartamentos_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then imgAceptarEmpleado_Click
End Sub

Private Sub cmbDivisa_Change()
    txtCambio = ""
    txtCambio.Tag = ""
End Sub

Private Sub cmbDivisa_Click()
    
    If cmbDivisa.Tag <> "" Then Exit Sub
    'If lblContador <> "" Then Exit Sub
    If cmbDivisa.Text = "" Then Exit Sub
    'If Trim(txtCambio) <> "" Then Exit Sub
    If Trim(txtFecha) = "" Then Exit Sub
    
    txtCambio.Text = dDameCambioDivisa(CodPaisTrabajo, cmbDivisa.Text, txtFecha, msMonedaDepartamento)
    ReCalcularImporteCambio
    
End Sub

Private Sub cmbDivisa_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub cmbDivisa_LostFocus()
    cmbDivisa_Click
End Sub

Private Sub cmbEmpleados_Change()

    Dim sAux As String

    If mbEmpleados = False Then Exit Sub

    'VaciarBox cmbEmpleados
    Dim iCont As Integer
    
    For iCont = 0 To cmbEmpleados.ListCount - 1
        cmbEmpleados.RemoveItem 0
    Next iCont
    
    If cmbEmpleados.Text = "" Then
    
        While Not mafoRegEmpleados.EOF
        
            sAux = Trim(mafoRegEmpleados("Nombre"))
            If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegEmpleados("Codigo") & ") " & sAux
            cmbEmpleados.AddItem sAux
            cmbEmpleados.ItemData(cmbEmpleados.NewIndex) = mafoRegEmpleados("Codigo")
            mafoRegEmpleados.MoveNext
        
        Wend
    
    Else
    
        While Not mafoRegEmpleados.EOF
        
            If InStr(1, UCase(Trim(mafoRegEmpleados("Nombre"))), UCase(cmbEmpleados.Text)) Then
                sAux = Trim(mafoRegEmpleados("Nombre"))
                If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegEmpleados("Codigo") & ") " & sAux
                cmbEmpleados.AddItem sAux
                cmbEmpleados.ItemData(cmbEmpleados.NewIndex) = mafoRegEmpleados("Codigo")
            End If
            
            mafoRegEmpleados.MoveNext
        
        Wend
    
    End If
    mafoRegEmpleados.MoveFirst
    
End Sub

Private Sub cmbEmpleados_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then imgAceptarEmpleado_Click
End Sub

Private Sub cmbEmpresas_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub cmdBorrarLinea_Click()
    If cmdBorrarLinea.Tag <> "" Then Exit Sub
    cmdBorrarLinea.Tag = "A"
    imgBorrarLinea_Click
    cmdBorrarLinea.Tag = ""
End Sub

Private Sub cmdCerrarHistorico_Click()
    fraHistorico.Visible = False
    fraDatos.Enabled = True
End Sub

Private Sub cmdContaAnticipo_Click()

    Dim oZSAP As Object
    Dim sSociedad As String
    Dim sFechaContaSAP As String
    Dim sEjercicio As String
    Dim sAcreedor As String
    Dim sMoneda As String
    Dim sImporte As String
    Dim sTexto As String
    
    Dim oRetorno As Object
    Dim sSociedadSAP As String
    Dim sNumDocumentoSAP As String
    Dim sEjercicioSAP As String
    
    Dim sAux As String
    Dim bErrores As Boolean
    
    Exit Sub

    
    
    sSociedad = "0610"
    sFechaContaSAP = Format(Now, "yyyy-MM-dd")
    sEjercicio = Year(Now)
    sAcreedor = "0000301761"
    sMoneda = "EUR"
    sImporte = "-25"
    sTexto = "ANTICIPO GASTOS 20-08-19"
    
    Set oZSAP = oBapiControl.GetSAPObject("ZTRAVELEXP")
    Set oRetorno = Nothing
    
    oZSAP.ZbapiContabilizarAnticipos AWTYP:="BKPF", _
                            GLVOR:="RFST", _
                            BKTXT:="ANTICIPO GASTOS", _
                            POSNR:="0000000001", _
                            BUKRS:=sSociedad, _
                            GJAHR:=sEjercicio, _
                            BUZEI:="001", _
                            BLDAT:=sFechaContaSAP, _
                            BUDAT:=sFechaContaSAP, _
                            BLART:="KA", _
                            XBLNR:="ANTICIPO GASTOS", _
                            ZUMSK:="L", _
                            UMSKZ:="F", _
                            BSTAT:="S", _
                            BSCHL:="39", _
                            LIFNR:=sAcreedor, _
                            SGTXT:=sTexto, _
                            WAERS:=sMoneda, _
                            WRBTR:=sImporte, _
                            RETURN:=oRetorno, _
                            EBUKRS:=sSociedadSAP, _
                            EBELNR:=sNumDocumentoSAP, _
                            EGJAHR:=sEjercicioSAP
                            
    If oRetorno Is Nothing Then
        bErrores = True
    Else
        If Trim(oRetorno("TYPE")) = "E" Then
            bErrores = True
            MsgBox "Error al contabilizar el gasto '" & sTexto & "'." & vbCr & "ERROR: " & Trim(oRetorno("MESSAGE"))
            sAux = "ERROR/TYPE:" & Trim(oRetorno("TYPE")) & "/ID:" & Trim(oRetorno("ID")) & _
                   "/NUMBER:" & Trim(oRetorno("NUMBER")) & "/MESSAGE:" & Trim(oRetorno("MESSAGE")) & "/LOG_NO:" & Trim(oRetorno("LOG_NO")) & _
                   "/LOG_MSG_NO:" & Trim(oRetorno("LOG_MSG_NO")) & "/MESSAGE_V1:" & Trim(oRetorno("MESSAGE_V1")) & _
                   "/MESSAGE_V2:" & Trim(oRetorno("MESSAGE_V2")) & "/MESSAGE_V3:" & Trim(oRetorno("MESSAGE_V3")) & _
                   "/MESSAGE_V4:" & Trim(oRetorno("MESSAGE_V4")) & "/PARAMETER:" & Trim(oRetorno("PARAMETER")) & _
                   "/ROW:" & Trim(oRetorno("ROW")) & "/FIELD:" & Trim(oRetorno("FIELD")) & "/SYSTEM:" & Trim(oRetorno("SYSTEM"))
        Else
            sAux = "OK/TYPE:" & Trim(oRetorno("TYPE")) & "/ID:" & Trim(oRetorno("ID")) & _
                   "/NUMBER:" & Trim(oRetorno("NUMBER")) & "/MESSAGE:" & Trim(oRetorno("MESSAGE")) & "/LOG_NO:" & Trim(oRetorno("LOG_NO")) & _
                   "/LOG_MSG_NO:" & Trim(oRetorno("LOG_MSG_NO")) & "/MESSAGE_V1:" & Trim(oRetorno("MESSAGE_V1")) & _
                   "/MESSAGE_V2:" & Trim(oRetorno("MESSAGE_V2")) & "/MESSAGE_V3:" & Trim(oRetorno("MESSAGE_V3")) & _
                   "/MESSAGE_V4:" & Trim(oRetorno("MESSAGE_V4")) & "/PARAMETER:" & Trim(oRetorno("PARAMETER")) & _
                   "/ROW:" & Trim(oRetorno("ROW")) & "/FIELD:" & Trim(oRetorno("FIELD")) & "/SYSTEM:" & Trim(oRetorno("SYSTEM"))
        End If
    
    End If


End Sub

Private Sub cmdDeslogon_Click()
    DesLogon
End Sub


Private Sub cmdEncript_Click()
    ListarEncriptacion
End Sub

Private Sub cmdLogonGESCalidad_Click()
    Logon "ges_calidad"
End Sub

Private Sub cmdLogonGESDesarrollo_Click()
    Logon "ges_desarrollo"
End Sub

Private Sub cmdLogonGESProduccion_Click()
    Logon "ges_produccion"
End Sub

Private Sub cmdRechazar_Click()

    Dim sMotivoRechazo As String
    
    If lblContador = "" Then Exit Sub
    If lblPagado = "S" Then
        MsgBox "El gasto ya está pagado y no puede ser rechazado.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    If lblContabilizadoSAP = "S" Then
        MsgBox "El gasto ya está contabilizado y no puede ser rechazado.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    If lblValidado = "R" Then
        MsgBox "El gasto ya está rechazado.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    If MsgBox("Va a rechazar el gasto, ¿está seguro?", vbApplicationModal + vbQuestion + vbYesNo, "Confirmación") = vbNo Then Exit Sub
    
    frmPedirMotivoRechazo.Show MODAL
    sMotivoRechazo = frmPedirMotivoRechazo.MotivoRechazo
    If sMotivoRechazo = "@@CANCELADO##" Then
        MsgBox "NO se ha rechazo el gasto.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    
    'Crear función para registrar el rechazo.
    If Not bRechazarGasto(lblContador, sMotivoRechazo) Then
        MsgBox "Se ha producido un error al intentar realizar la acción.", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    RellenarDatosGastos

End Sub

Private Sub cmdSubirLinea_Click()
    If cmdSubirLinea.Tag <> "" Then Exit Sub
    cmdSubirLinea.Tag = "A"
    imgAceptarLinea_Click
    cmdSubirLinea.Tag = ""
End Sub

Private Sub Form_Load()

    Dim sAux1 As String
    Dim sAux2 As String
    
    Dim iAux As Integer
    Dim iCont As Integer
    
    miColumnaOrdenacion = -1
    miPosicionarEnFila = -1
    
    mfgTiposGasto.BorderStyle = flexBorderNone
    
    imgPassword.Visible = False
    'Por indicaciones de Filo no se pueden consultar password
    'If goUsuario.IdUsuario = "lanchi" Or goUsuario.IdUsuario = "mmarsanchez" Or goUsuario.IdUsuario = "agonzalez" Or goUsuario.IdUsuario = "ivazquez" Or goUsuario.IdUsuario = "janaharro" Or goUsuario.IdUsuario = "fsanchezc" Then imgPassword.Visible = True
    
'    iAux = 0
'    If goUsuario.IdUsuario = "mmarsanchez" Then iAux = 1000
'    If DisplayMonitorCount = 1 Then
'        Me.Left = (Screen.Width / 2) - (Me.Width / 2)
'        Me.Top = (Screen.Height / 2) - (Me.Height / 2)
'    Else
'        'Me.Move Screen.Width + (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'        'Me.Move gAnchoPantallaPrincipal + (((Screen.Width - gAnchoPantallaPrincipal) - Me.Width) \ 2), (Screen.Height - Me.Height) \ 2
'        If MDIPrincipal.Left + 120 < Screen.Width Then
'        'El MDI está en la pantalla 1, lo centro en la pantalla 1
'            Me.Left = (Screen.Width / 2) - (Me.Width / 2)
'            Me.Top = (Screen.Height / 2) - (Me.Height / 2) - iAux
'        Else
'            Me.Left = gAnchoPantallaPrincipal + (((gAnchoPantallaVirtual - gAnchoPantallaPrincipal) - Me.Width) \ 2)
'            Me.Top = ((Screen.Height - Me.Height) \ 2) - iAux
'            'Me.Move gAnchoPantallaPrincipal + (((gAnchoPantallaVirtual - gAnchoPantallaPrincipal) - Me.Width) \ 2), (Screen.Height - Me.Height) \ 2
'            'MsgBox "Ancho mdi: " & gAnchoPantallaPrincipal & " - Ancho virtual: " & gAnchoPantallaVirtual & " - Ancho form: " & Me.Width & " - posicion izq form: " & Me.Left
'        End If
'    End If
    
    EmpleadoDesBloquear_UsuGasto
    mid_Ofi_PreValidacion = 0
    bEsUsuarioPreValidacion CodPaisTrabajo, Trim(goUsuario.IdUsuario), mid_Ofi_PreValidacion

    miCodDepartamento = -1
    miCodEmpleado = -1
    msJerEmpleado = ""
    msAcreedorEmpleado = ""
    msMonedaDepartamento = ""
    msDivision_PersonalEmpleado = ""
    msClave_OrganizacionEmpleado = ""
    miManagerEmpleado = -1
    miOfi_PreValidacionEmpleado = -1
    
    shRecuadroForm.Width = Me.Width
    shRecuadroForm.Height = Me.Height
    shRecuadroForm.Left = 0
    shRecuadroForm.Top = 0
    
    fraHistorico.BorderStyle = vbBSNone
    fraHistorico.Visible = False
    fraHistorico.Top = 1920
    fraHistorico.Left = 120
    
    fraMensaje.BorderStyle = vbBSNone
    fraMensaje.Visible = False
    fraMensaje.Top = 1920
    fraMensaje.Left = 120
    lblMensaje1 = ""
    lblMensaje2 = ""
    lblMensaje3 = ""
    
    cmdLogonGESProduccion.Visible = False
    cmdLogonGESCalidad.Visible = False
    cmdLogonGESDesarrollo.Visible = False
    cmdDeslogon.Visible = False
    cmdContaAnticipo.Visible = False
    cmdEncript.Visible = False
    If goUsuario.IdUsuario = "lanchi" Then
        cmdLogonGESProduccion.Visible = True
        cmdLogonGESCalidad.Visible = True
        cmdLogonGESDesarrollo.Visible = True
        cmdDeslogon.Visible = True
        cmdContaAnticipo.Visible = True
        chkTodosPreValidacion.Visible = True
        cmdEncript.Visible = True
        cmdRechazar.Visible = True
    End If
    
    mbConexionSAP = False
    
    msTipo = sDameTipoGastoPais(CodPaisTrabajo)
    
    If bDameDivisas(mafoRegistro) Then
        cmbDivisa.Tag = "C"
        While Not mafoRegistro.EOF
            cmbDivisa.AddItem Trim(mafoRegistro("Valor"))
            mafoRegistro.MoveNext
        Wend
        cmbDivisa.Tag = ""
    End If
    Set mafoRegistro = Nothing
    
    cmbDepartamentos.Tag = "C"
    If bDameDepartamentosCL(CodPaisTrabajo, mafoRegistro, mid_Ofi_PreValidacion) Then
        While Not mafoRegistro.EOF
            If goUsuario.IdUsuario = "ldelgado" And CodPaisTrabajo <> giPAIS_PERSEIDA Then
                If mafoRegistro("CodDepartamento") = 1201 And mafoRegistro("Activo") = "S" Then
                    cmbDepartamentos.AddItem Trim(mafoRegistro("Departamento"))
                    cmbDepartamentos.ItemData(cmbDepartamentos.NewIndex) = mafoRegistro("CodDepartamento")
                End If
            Else
                If mafoRegistro("Activo") = "S" Then
                    cmbDepartamentos.AddItem Trim(mafoRegistro("Departamento"))
                    cmbDepartamentos.ItemData(cmbDepartamentos.NewIndex) = mafoRegistro("CodDepartamento")
                End If
            End If
            mafoRegistro.MoveNext
        Wend
    End If
    cmbDepartamentos.Tag = ""
    Set mafoRegistro = Nothing
    
    If bDameEmpresasGasto(mafoRegistro) Then
        While Not mafoRegistro.EOF
            cmbEmpresas.AddItem Trim(mafoRegistro("Empresa"))
            cmbEmpresas.ItemData(cmbEmpresas.NewIndex) = mafoRegistro("id_Empresa")
            mafoRegistro.MoveNext
        Wend
        For iCont = 0 To cmbEmpresas.ListCount - 1
            If CInt(cmbEmpresas.ItemData(iCont)) = Empresa_CodPaisTrabajo Then
                cmbEmpresas.ListIndex = iCont
                Exit For
            Else
                cmbEmpresas.ListIndex = 0
            End If
        Next
        'cmbEmpresas.ListIndex = 0
    End If
    Set mafoRegistro = Nothing
    
    mbCuentasContables = False
    If CodPaisTrabajo = giPAIS_GES Then
        If bDameCuentasContablesGES(mafoRegCuentasContables) Then
            mbCuentasContables = True
            While Not mafoRegCuentasContables.EOF
                cmbCuentasContables.AddItem Trim(mafoRegCuentasContables("Cuenta_Contable")) & " - " & Trim(mafoRegCuentasContables("DescCuenta_Contable"))
                mafoRegCuentasContables.MoveNext
            Wend
        End If
        mafoRegCuentasContables.MoveFirst
    End If
    
    dtpFechaDesde.value = Now - 30
    dtpFechaHasta.value = Now
    
    'Si se añaden nuevas columnas, tener en cuenta que hay que modificar el doble click de la parrilla para los indicadores de selección y buscar los sitios donde se pregunte por
    'las columnas Validado, Pagado e Importe2
    moParrilla.Create mfgParrilla, 51, 2
    moParrilla.Formatear 0, 0, "Contador"
    moParrilla.Formatear 1, 0, "RutaTicket"
    moParrilla.Formatear 2, 400, "", 4, 4
    moParrilla.Formatear 3, 1250, "Fecha", 4, 4, True
    moParrilla.Formatear 4, 1250, "C. Gasto", 4, , True
    moParrilla.Formatear 5, 1250, "T. Gasto", 4
    moParrilla.Formatear 6, 1250, "T. Pago", 4
    moParrilla.Formatear 7, 4000, "Donde / Itinerario / Concepto", 4
    moParrilla.Formatear 8, 1250, "T. Comida", 4
    moParrilla.Formatear 9, 1250, "Nº Noches", 4, 4
    moParrilla.Formatear 10, 1000, "Kms", 4, 7
    moParrilla.Formatear 11, 1250, "Nº Tickets", 4, 4
    moParrilla.Formatear 12, 1250, "Nº Personas", 4, 4
    moParrilla.Formatear 13, 1250, "Importe", 4, 7, True                 'se va a cargar el importe_ori
    moParrilla.Formatear 14, 3000, "Proveedor", 4
    moParrilla.Formatear 15, 1500, "Factura", 4
    moParrilla.Formatear 16, 3000, "Personas", 4
    moParrilla.Formatear 17, 2000, "Empresa", 4
    moParrilla.Formatear 18, 1500, "PEP/CECO/IO", 4, , True
    moParrilla.Formatear 19, 1000, "Divisa", 4, 4
    moParrilla.Formatear 20, 1000, "Cambio", 4, 7
    moParrilla.Formatear 21, 2000, "Archivado en", 4
    moParrilla.Formatear 22, 1500, "IF Proveedor", 4
    moParrilla.Formatear 23, 0, "Validado"
    moParrilla.Formatear 24, 0, "Pagado"
    moParrilla.Formatear 25, 0, "Importe2"
    moParrilla.Formatear 26, 0, "Conta SAP"
    moParrilla.Formatear 27, 0, "Importe_C"                         'importe tras aplicarle el cambio
    moParrilla.Formatear 28, 0, "Modificado_Adm"
    moParrilla.Formatear 29, 1500, "F. Val/Rec", 4, 4, True
    moParrilla.Formatear 30, 2500, "Motivo rechazo", 4
    moParrilla.Formatear 31, 2500, "Cuenta contable", 4
    moParrilla.Formatear 32, 2500, "Fecha grabación gasto", 4, 7
    moParrilla.Formatear 33, 3250, "Documento SAP / Fecha Conta", 4, 7, True
    moParrilla.Formatear 34, 4000, "Validador", 4, , True
    moParrilla.Formatear 35, 0, "Cod_Cta_Con"
    moParrilla.Formatear 36, 0, "Cambio_Usu"
    moParrilla.Formatear 37, 0, "CPD_Nombre"
    moParrilla.Formatear 38, 0, "CPD_CP"
    moParrilla.Formatear 39, 0, "CPD_Ciudad"
    moParrilla.Formatear 40, 0, "CPD_Pais"
    moParrilla.Formatear 41, 0, "CPD_CIF"
    moParrilla.Formatear 42, 1500, "Desglose IVA", 4, 4
    moParrilla.Formatear 43, 0, "Importe IVA", 4, 7
    moParrilla.Formatear 44, 0, "Base IVA", 4, 7
    moParrilla.Formatear 45, 0, "Tipo IVA", 4, 4
    moParrilla.Formatear 46, 0, "Importe IVA 2", 4, 7
    moParrilla.Formatear 47, 0, "Base IVA 2", 4, 7
    moParrilla.Formatear 48, 0, "Tipo IVA 2", 4, 4
    moParrilla.Formatear 49, 0, "Contador asociado", 4, 4
    moParrilla.Formatear 50, 400, "", 4, 4
    moParrilla.SituarEn 0, 1
    
    moParrillaHistorico.Create mfgHistoricoModificaciones, 38, 2
    moParrillaHistorico.Formatear 0, 0, ""
    moParrillaHistorico.Formatear 1, 1750, "Fecha modi.", 4, 4
    moParrillaHistorico.Formatear 2, 1500, "Usuario modi.", 4, 4
    moParrillaHistorico.Formatear 3, 1250, "Fecha", 4, 4
    moParrillaHistorico.Formatear 4, 1250, "C. Gasto", 4
    moParrillaHistorico.Formatear 5, 1250, "T. Gasto", 4
    moParrillaHistorico.Formatear 6, 1250, "T. Pago", 4
    moParrillaHistorico.Formatear 7, 4000, "Donde / Itinerario / Concepto", 4
    moParrillaHistorico.Formatear 8, 1250, "T. Comida", 4
    moParrillaHistorico.Formatear 9, 1250, "Nº Noches", 4, 4
    moParrillaHistorico.Formatear 10, 1000, "Kms", 4, 7
    moParrillaHistorico.Formatear 11, 1250, "Nº Tickets", 4, 4
    moParrillaHistorico.Formatear 12, 1250, "Nº Personas", 4, 4
    moParrillaHistorico.Formatear 13, 1250, "Importe", 4, 7                  'se va a cargar el importe_ori
    moParrillaHistorico.Formatear 14, 3000, "Proveedor", 4
    moParrillaHistorico.Formatear 15, 1500, "Factura", 4
    moParrillaHistorico.Formatear 16, 3000, "Personas", 4
    moParrillaHistorico.Formatear 17, 2000, "Empresa", 4
    moParrillaHistorico.Formatear 18, 1500, "PEP/CECO/IO", 4
    moParrillaHistorico.Formatear 19, 1000, "Divisa", 4, 4
    moParrillaHistorico.Formatear 20, 1000, "Cambio", 4, 7
    moParrillaHistorico.Formatear 21, 1500, "Imp. Cambio", 4, 7              'importe tras aplicarle el cambio
    moParrillaHistorico.Formatear 22, 2000, "Archivado en", 4
    moParrillaHistorico.Formatear 23, 1500, "IF Proveedor", 4
    moParrillaHistorico.Formatear 24, 1250, "Validado", 4, 4
    moParrillaHistorico.Formatear 25, 1250, "Pagado", 4, 4
    moParrillaHistorico.Formatear 26, 1250, "Conta. SAP", 4, 4
    moParrillaHistorico.Formatear 27, 1500, "F. Val/Rec", 4, 4
    moParrillaHistorico.Formatear 28, 2500, "Motivo rechazo", 4
    moParrillaHistorico.Formatear 29, 2500, "Cuenta contable", 4
    moParrillaHistorico.Formatear 30, 1500, "Desglose IVA", 4, 4
    moParrillaHistorico.Formatear 31, 1500, "Importe IVA", 4, 7
    moParrillaHistorico.Formatear 32, 1500, "Base IVA", 4, 7
    moParrillaHistorico.Formatear 33, 1500, "Tipo IVA", 4, 4
    moParrillaHistorico.Formatear 34, 1500, "Importe IVA 2", 4, 7
    moParrillaHistorico.Formatear 35, 1500, "Base IVA 2", 4, 7
    moParrillaHistorico.Formatear 36, 1500, "Tipo IVA 2", 4, 4
    moParrillaHistorico.Formatear 37, 6000, "Motivo modificación", 4
    moParrillaHistorico.SituarEn 0, 1
    
    moTiposGasto.Create mfgTiposGasto, 4, 1, 0, 0
    moTiposGasto.Formatear 0, 0, ""
    moTiposGasto.Formatear 1, 2300, ""
    moTiposGasto.Formatear 2, 100, ""
    moTiposGasto.Formatear 3, 2300, ""
    If bDameTiposGasto(msTipo, mafoRegistro) Then
        While Not mafoRegistro.EOF
            
            sAux1 = ""
            sAux2 = ""
            
            sAux1 = Trim(mafoRegistro("TipoGasto")) & "->" & Trim(mafoRegistro("NombreGasto"))
            If Not mafoRegistro.EOF Then mafoRegistro.MoveNext
            If Not mafoRegistro.EOF Then sAux2 = Trim(mafoRegistro("TipoGasto")) & "->" & Trim(mafoRegistro("NombreGasto"))
            
            If moTiposGasto.LineaLibre >= 1 Then moTiposGasto.Grid.Rows = moTiposGasto.Grid.Rows + 1
            
            moTiposGasto.TextoFila moTiposGasto.LineaLibre, "", sAux1, "", sAux2
            
            If Not mafoRegistro.EOF Then mafoRegistro.MoveNext
            
        Wend
    End If
    Set mafoRegistro = Nothing
    moTiposGasto.SituarEn 0, 1
        
    txtPEPCECO.Enabled = False
    cmbDivisa.Enabled = False
    txtCambio.Enabled = False
    shCambioUsuario.Visible = False
    txtCIF_Proveedor.Enabled = False
    cmbCuentasContables.Enabled = False
    Label4(2).Caption = "1 ó N->Metálico    2 ó S->VISA       3 ó F->Facturado   4 ó R->SolRed"
    msFiltroTipoPago = "1nN2sS3fF4rR"
    imgContabilizarSAP.Visible = False
    imgCancelarContabilizarSAP.Visible = False
    chkPedirFechaContaSAP.Visible = False
    imgArchivarEn.Visible = False
    chkDesgloseIVA.Enabled = False
    txtImporteIVA.Enabled = False
    txtBaseIVA.Enabled = False
    txtTipoIVA.Enabled = False
'    txtImporteIVA2.Enabled = False
'    txtBaseIVA2.Enabled = False
'    txtTipoIVA2.Enabled = False
    '18-06-18 Por indicaciones desde GES, sólo vamos a tener tipo de pago en Metálico
    If CodPaisTrabajo = giPAIS_GES Then
        
        txtFamiliaGasto.Enabled = False
        txtFamiliaGasto = "Ordinario"
        txtProveedor.Enabled = False
        txtFactura.Enabled = False
        cmbEmpresas.Enabled = False
        cmbEmpresas = "GES"
        txtPEPCECO.Enabled = True
        cmbDivisa.Enabled = True
        txtCambio.Enabled = True
        
        '14-04-21 Solicitan que en GES también van a empezar a utilizar tarjetas de credito, por lo que comentamos la parte que teníamos especial en esta empresa
        ''Label4(2).Caption = "1 ó N->Metálico    2 ó S->VISA"
        ''msFiltroTipoPago = "1nN2sS"
        'txtTipoPago.Enabled = False
        'txtTipoPago = "Metálico"
        'Solo vamos a permitir que selecciones Metálico o Visa
        Label4(2).Caption = "1 ó N->Metálico    2 ó S->VISA"
        msFiltroTipoPago = "1nN2sS"
        '14-04-21 Fin
        
        imgContabilizarSAP.Visible = True
        If goUsuario.IdUsuario = "lanchi" Or goUsuario.IdUsuario = "agonzalez" Or goUsuario.IdUsuario = "ivazquez" Or goUsuario.IdUsuario = "luciagil" Or goUsuario.IdUsuario = "inmamoya" Or goUsuario.IdUsuario = "ialcal" Or goUsuario.IdUsuario = "lvazquezmo" Then imgCancelarContabilizarSAP.Visible = True
        chkPedirFechaContaSAP.Visible = True
        imgArchivarEn.Visible = True
    
        If mid_Ofi_PreValidacion = 0 Then
        'Comprobamos sí hay gastos del tipo Vendor/Proveedor que no tengan el CIF validado. Siempre que no sea prevalidador
            If bCIF_Proveedores_SinConciliar(CodPaisTrabajo) Then
                frmConciliarProveedores.Create CodPaisTrabajo
                frmConciliarProveedores.Show MODAL
            End If
        End If
        
        cmdRechazar.Visible = True
        imgInforme.Visible = True
        
        chkDesgloseIVA.Enabled = True
        'El resto de campos de IVA no hay que activarlos porque los maneja el check desglose iva
        
    End If
    
    imgPreValidar.Visible = False
    If mid_Ofi_PreValidacion > 0 Then
    
        '19-11-19 Piden que los prevalidadores puedan ver todos los gastos
'        Shape3.Visible = False
'        Shape7.Visible = False
'        Shape1.Visible = False
'        Shape2(1).Visible = False
'        Shape2(0).Visible = False
'        lblTexto(4).Visible = False
'        lblTexto(32).Visible = False
'        lblTexto(6).Visible = False
'        lblTexto(26).Visible = False
'        lblTexto(5).Visible = False
'        lblTexto(28).Visible = False
'        shFiltro_PendienteValidar.Visible = False
'        shFiltro_Rechazado.Visible = False
'        shFiltro_ContabilizadoSAP.Visible = False
'        shFiltro_ValidadoPagado.Visible = False
        '19-11-19 Fin
        imgValidar.Visible = False
        imgPagar.Visible = False
        imgImprimirPagos.Visible = False
        imgArchivarEn.Visible = False
        imgContabilizarSAP.Visible = False
        imgCancelarContabilizarSAP.Visible = False
        chkPedirFechaContaSAP.Visible = False
        imgBuscarUsuariosConGastosValidados.Visible = False
        imgInforme.Visible = False
        lblTexto(8) = "Prevalidación gastos personal interno"
        
        chkTodosPreValidacion.Visible = True
        imgPreValidar.Visible = True
        
    End If
    
    cMouseW.InitMouseWheel mfgParrilla.hwnd
    cMouseW.InitMouseWheel mfgHistoricoModificaciones.hwnd
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub fraDatos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub fraPersonal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub Image1_Click()

End Sub

Private Sub imgAceptarEmpleado_Click()
    
    Dim sUsuBloqueo As String
    
    Me.MousePointer = DEFAULT

    If cmbDepartamentos.Text = "" Then
        cmbDepartamentos.SetFocus
        Exit Sub
    End If
    If cmbEmpleados.Text = "" Then
        cmbEmpleados.SetFocus
        Exit Sub
    End If
    If cmbEmpleados.ListIndex = -1 Then
        cmbEmpleados.SetFocus
        Exit Sub
    End If
    
    miCodDepartamento = cmbDepartamentos.ItemData(cmbDepartamentos.ListIndex)
    miCodEmpleado = cmbEmpleados.ItemData(cmbEmpleados.ListIndex)
    msJerEmpleado = sDameJerarquiaEmpleado(CodPaisTrabajo, miCodDepartamento, miCodEmpleado)
    msMonedaDepartamento = sDameMonedaDepartamento(CodPaisTrabajo, miCodDepartamento)
    DameDatosAuxEmpleado CodPaisTrabajo, miCodDepartamento, miCodEmpleado, msJerEmpleado, msAcreedorEmpleado, msDivision_PersonalEmpleado, msClave_OrganizacionEmpleado, miManagerEmpleado, miOfi_PreValidacionEmpleado
    'msAcreedorEmpleado = sDameAcreedorEmpleado(CodPaisTrabajo, miCodDepartamento, miCodEmpleado, msJerEmpleado)
    If CodPaisTrabajo = giPAIS_GES Then If msAcreedorEmpleado = "" Then lblEmpleadoSinNumAcreedor.Visible = True
    
    Select Case CodPaisTrabajo
    Case giPAIS_MARRUECOS
        msImporteKms = "0.33"
    Case giPAIS_ITALIA
        msImporteKms = "0.08"
    Case giPAIS_PORTUGAL
        msImporteKms = "0.1"
    Case giPAIS_CHILE
        msImporteKms = "15"
    Case giPAIS_COLOMBIA
        msImporteKms = "75"
    Case giPAIS_ESPAÑA
        msImporteKms = "0.03"
        '02-04-19 Según indicaciones de Rafael Carrascal, para el usuario de Xavier Guillem el importe del Km se le va a pagar a 0.3 en lugar de 0.03
        If miCodDepartamento = 21 And msJerEmpleado = "JD" And miCodEmpleado = 1 Then msImporteKms = "0.3"
    Case giPAIS_ONDUPACK, giPAIS_ONDUPET, giPAIS_PLASTIVERD, giPAIS_MURCIACARTON, giPAIS_CARTONAJESEXTREMADURA
    '18-09-19 Amparo solicita que el importe del KM en Alter Enersun sea 0.19
    '27-09-19 Amparo solicita que el importe del KM en Ondupack/Ondupet sea 0.19
    '13-11-20 Amparo solicita que el importe del KM en Plastiverd sea de 0.19
    '26-10-22 Amparo solicita que el importe del KM en Murcia Cartón sea de 0.19
        msImporteKms = "0.19"
    Case giPAIS_MATIASGOMA
    '24-02-20 Amparo solicita que el importe del KM en Matías Gomá sea 0.24
        msImporteKms = "0.24"
    Case giPAIS_AGSIDERURGICABALBOA, giPAIS_GRUPOGALLARDOBALBOA, giPAIS_ALFONSOGALLARDO, giPAIS_CORRUGADOSLASAO, giPAIS_MARCELIANOMARTIN, giPAIS_FERROMALLAS, giPAIS_CORRUGADOSGETAFE, giPAIS_GALVACOLOR, giPAIS_SOLARSTEEL
    '14-10-20 Solicitan que para las empresas del Grupo Gallardo los Kms se paguen a 0.24
        msImporteKms = "0.24"
    Case giPAIS_IQOXE, giPAIS_IQLIT
    '09-10-20 Amparo solicita que el importe del KM en Iqoxe sea 0.19
        msImporteKms = "0.19"
    Case giPAIS_ALTERENERSUN, giPAIS_GASILUZ
    '29-03-23 Jesus Ramos me pasa mail de Angela Perez Acosta en la que indica que el gasto de KMs en Alter y Gasiluz por convenio debe ser 0.28
        msImporteKms = 0.28
    Case giPAIS_CRISTIANLAYDSS
        msImporteKms = 0.03
        'Red comercial de Italia
        If miCodDepartamento = 30 Then msImporteKms = 0.08
        'Red comercial de Portugal
        If miCodDepartamento = 31 Then msImporteKms = 0.15
    Case giPAIS_FUNDACIONRL
        msImporteKms = 0.3
    Case Else
        msImporteKms = "0.03"
    End Select
    
'    If bEmpleadoBloqueado(CodPaisTrabajo, miCodDepartamento, miCodEmpleado, msJerEmpleado) Then
'        miCodEmpleado = -1
'        msJerEmpleado = ""
'        MsgBox "El trabajador está bloqueado por otro usuario y no puede trabajar con él.", vbApplicationModal + vbExclamation, "Información"
'        Exit Sub
'    End If
    sUsuBloqueo = sEmpleadoBloqueado_Por(CodPaisTrabajo, miCodDepartamento, miCodEmpleado, msJerEmpleado)
    If sUsuBloqueo <> "" Then
        miCodEmpleado = -1
        msJerEmpleado = ""
        MsgBox "El trabajador está bloqueado por el usuario " & sUsuBloqueo & " y no puede trabajar con él.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    EmpleadoBloquear CodPaisTrabajo, miCodDepartamento, miCodEmpleado, msJerEmpleado, goUsuario.IdUsuario
    
    fraPersonal.Enabled = False
    fraDatos.Enabled = True
    
    '19-12-19 MariMar pide que no nos vayamos una año atrás en las búsquedas, lo dejamos como está por defecto, un mes atrás
'    If CodPaisTrabajo = giPAIS_GES Then
'        dtpFechaDesde.Value = Now - 365
'        dtpFechaHasta.Value = Now
'    End If
    
    RellenarDatosGastos
    
End Sub

Private Sub imgAceptarEmpleado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgAceptarLinea_Click()
     
    Dim sSqlAux As String
    Dim afoRegAux As ADODB.Recordset
    
    Dim iNumero As Double
    Dim iNumero2 As Double
    Dim sCadena As String
    Dim sPagado As String
    Dim sValidado As String
    Dim sValidacion As String
    Dim sEnv_Mail_Validacion As String
    Dim iFila As Integer
    
    Dim sPEPCECO As String
    Dim sCambio As String
    Dim sMoneda As String
    Dim dImporte As Double
    Dim dImporte_Ori As Double
    Dim sCIF_Proveedor As String
    Dim sid_Proveedor As String
    Dim sProveedor_OK As String
    Dim sCuenta_Contable As String
    Dim sDesglose_IVA As String
    Dim dImporte_IVA As Double
    Dim dBase_IVA As Double
    Dim sTipo_IVA As String
    Dim dImporte_IVA2 As Double
    Dim dBase_IVA2 As Double
    Dim sTipo_IVA2 As String
    Dim dAux As Double
    
    Dim sJustificante As String
    Dim sExtension As String
    Dim sRutaNAS As String
    Dim sRutaMarte As String
    
    Dim sCambioGuardadoPorUsuario As String
    
    Dim sCPD_Nombre As String
    Dim sCPD_CP As String
    Dim sCPD_Ciudad As String
    Dim sCPD_Pais As String
    Dim sCPD_CIF As String
    
    Dim sMotivoModificacion As String
    
    Dim sValidador As String
    Dim iOfi_PreValidacion As Integer
    Dim lResponsable_PEPCECO As Long
    Dim lDirectorProyecto_PEPCECO As Long
    Dim sPreValidador As String
    
    Dim lContadorAsociado As Long
    
    On Error GoTo Error_Rutina
    
    Me.MousePointer = DEFAULT
    
    If imgAceptarLinea.Tag <> "" Then Exit Sub
    imgAceptarLinea.Tag = "A"
    imgAceptarLinea.Tag = ""
    
    If lblContador = "" And mid_Ofi_PreValidacion > 0 Then
        imgAceptarLinea.Tag = ""
        MsgBox "No tiene permiso para insertar gastos.", vbApplicationModal + vbExclamation, "Información"
        LimpiarCampos
        Exit Sub
    End If
        
    '18-06-18
    'Para GES una vez contabilizado en SAP o pagado, no permitimos hacer cambios
    If CodPaisTrabajo = giPAIS_GES Then
        If lblPagado = "S" Or lblContabilizadoSAP = "S" Then
            imgAceptarLinea.Tag = ""
            MsgBox "El gasto ya está contabilizado/pagado y no se puede modificar.", vbApplicationModal + vbInformation, "Información"
            Exit Sub
        End If
        If lblRuta = "" And lblFichero = "" And Trim(txtConceptoGasto) <> "Anticipo" Then
            imgAceptarLinea.Tag = ""
            MsgBox "No hay un justificante para el gasto, debe adjuntarse un justificante antes de guardar.", vbApplicationModal + vbInformation, "Información"
            Exit Sub
        End If
    End If
    
    txtConceptoGasto = Trim(txtConceptoGasto)
    If txtConceptoGasto = "" Then
        imgAceptarLinea.Tag = ""
        txtConceptoGasto.SetFocus
        Exit Sub
    End If
    txtFecha = Trim(txtFecha)
    If txtFecha = "" Then
        imgAceptarLinea.Tag = ""
        txtFecha.SetFocus
        Exit Sub
    End If
    txtFamiliaGasto = Trim(txtFamiliaGasto)
    If txtFamiliaGasto = "" Then
        imgAceptarLinea.Tag = ""
        txtFamiliaGasto.SetFocus
        Exit Sub
    End If
    txtTipoPago = Trim(txtTipoPago)
    If txtTipoPago = "" Then
        imgAceptarLinea.Tag = ""
        txtTipoPago.SetFocus
        Exit Sub
    End If
    txtDonde = Trim(Replace(txtDonde, "'", " "))
    txtTipoComida = Trim(txtTipoComida)
    txtNumNoches = Trim(txtNumNoches)
    txtKms = Trim(txtKms)
    txtNumTickets = Trim(txtNumTickets)
    txtNumPersonas = Trim(txtNumPersonas)
    txtPersonas = Trim(Replace(txtPersonas, "'", " "))
    txtProveedor = Trim(Replace(txtProveedor, "'", " "))
    txtFactura = Trim(Replace(txtFactura, "'", " "))
    txtImporte = Trim(txtImporte)
    If txtImporte = "" Then
        imgAceptarLinea.Tag = ""
        txtImporte.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtImporte) Then
        imgAceptarLinea.Tag = ""
        txtImporte = ""
        txtImporte.SetFocus
        Exit Sub
    End If
    txtCIF_Proveedor = Trim(txtCIF_Proveedor)
    txtCIF_Proveedor = Replace(txtCIF_Proveedor, "'", "")
    'txtCIF_Proveedor = Replace(txtCIF_Proveedor, "-", "")
    sCIF_Proveedor = ""
    sid_Proveedor = ""
    sProveedor_OK = "N"
    
    iFila = moParrilla.Grid.Row
    
    iNumero = 0
    iNumero2 = 0
    sCadena = ""
    sCPD_Nombre = ""
    sCPD_CP = ""
    sCPD_Ciudad = ""
    sCPD_Pais = ""
    sCPD_CIF = ""
    Select Case txtConceptoGasto.Text
    Case "Gasolina"
    
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        sCadena = txtDonde
        
    Case "Kilómetros"
    
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        If txtKms = "" Then
            imgAceptarLinea.Tag = ""
            txtKms.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtKms) Then
            imgAceptarLinea.Tag = ""
            txtKms = ""
            txtKms.SetFocus
            Exit Sub
        End If
        
        iNumero = txtKms
        sCadena = txtDonde
        
    Case "Comidas-Invitaciones", "Comidas/Invitaciones"
    
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        If txtTipoComida = "" Then
            imgAceptarLinea.Tag = ""
            txtTipoComida.SetFocus
            Exit Sub
        End If
        If txtNumPersonas = "" Then
            imgAceptarLinea.Tag = ""
            txtNumPersonas.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtNumPersonas) Then
            imgAceptarLinea.Tag = ""
            txtNumPersonas = ""
            txtNumPersonas.SetFocus
            Exit Sub
        End If
        
        iNumero = txtNumPersonas
        sCadena = txtDonde
        
        If iNumero > 1 Then
            If txtPersonas = "" Then
                imgAceptarLinea.Tag = ""
                txtPersonas.SetFocus
                Exit Sub
            End If
        End If
        
    Case "Salones"
    
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        
        sCadena = txtDonde
        
    Case "Hoteles"
    
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        If txtNumNoches = "" Then
            imgAceptarLinea.Tag = ""
            txtNumNoches.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtNumNoches) Then
            imgAceptarLinea.Tag = ""
            txtNumNoches = ""
            txtNumNoches.SetFocus
            Exit Sub
        End If
        If txtNumPersonas = "" Then
            imgAceptarLinea.Tag = ""
            txtNumPersonas.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtNumPersonas) Then
            imgAceptarLinea.Tag = ""
            txtNumPersonas = ""
            txtNumPersonas.SetFocus
            Exit Sub
        End If
    
        sCadena = txtDonde
        iNumero = txtNumNoches
        iNumero2 = txtNumPersonas
        
    Case "Teléfono"
                
    
    Case "Billetes", "Taxis"
        
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        
        sCadena = txtDonde
        
    Case "Peaje", "Peajes"
    
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        If txtNumTickets = "" Then
            imgAceptarLinea.Tag = ""
            txtNumTickets.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtNumTickets) Then
            imgAceptarLinea.Tag = ""
            txtNumTickets = ""
            txtNumTickets.SetFocus
            Exit Sub
        End If
        
        sCadena = txtDonde
        iNumero = txtNumTickets
        
    Case "Parking"
    
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        
        sCadena = txtDonde
        
    Case "Otros gastos", "Otros Gastos"
        
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        
        sCadena = txtDonde
        
        'No vamos a hacer obligatoria la cuenta contable
'        If CodPaisTrabajo = giPAIS_GES Then
'            If cmbCuentasContables.Text = "" Then
'                imgAceptarLinea.Tag = ""
'                cmbCuentasContables.SetFocus
'                Exit Sub
'            End If
'        End If
        
    Case "Muestras"
        
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        
        sCadena = txtDonde
    
    Case "Formación"
        
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        
        sCadena = txtDonde
    
    Case "Vendor/Proveedor"
    
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        
        If txtCIF_Proveedor = "" Then
            imgAceptarLinea.Tag = ""
            txtCIF_Proveedor.SetFocus
            Exit Sub
        End If
        If Not bComprobarCIF_Proveedor_GES(txtCIF_Proveedor, mafoRegistro) Then
            imgAceptarLinea.Tag = ""
            MsgBox "El ID Fiscal no se encuentra en el sistema.", vbApplicationModal + vbExclamation, "Información"
            txtCIF_Proveedor.SetFocus
            Exit Sub
        End If
        If LCase(txtCIF_Proveedor) = "cpd" Then
            If Not bComprobarDatosProveedorCPD() Then
                imgAceptarLinea.Tag = ""
                MsgBox "Debe rellenar los datos necesarios para la contabilización del proveedor CPD."
                txtCIF_Proveedor.SetFocus
                Exit Sub
            End If
            sCPD_Nombre = lblCPD_Nombre
            sCPD_CP = lblCPD_CP
            sCPD_Ciudad = lblCPD_Ciudad
            sCPD_Pais = lblCPD_Pais
            sCPD_CIF = lblCPD_CIF
        End If
        sCIF_Proveedor = UCase(Trim(txtCIF_Proveedor))
        sid_Proveedor = Trim(mafoRegistro("id"))
        sProveedor_OK = "S"
        Set mafoRegistro = Nothing
        
        sCadena = txtDonde
        
    Case "Billetes Tren", "Billetes Avión"
        
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        
        sCadena = txtDonde
        
    Case "Anticipo"
        
        If txtDonde = "" Then
            imgAceptarLinea.Tag = ""
            txtDonde.SetFocus
            Exit Sub
        End If
        
        sCadena = txtDonde

    Case Else
        imgAceptarLinea.Tag = ""
        Exit Sub
    End Select
    
    If txtTipoPago = "Facturado" And txtProveedor = "" Then
        imgAceptarLinea.Tag = ""
        txtProveedor.SetFocus
        Exit Sub
    End If
    If txtTipoPago = "Facturado" And txtFactura = "" Then
        imgAceptarLinea.Tag = ""
        txtFactura.SetFocus
        Exit Sub
    End If
    If txtTipoPago <> "Facturado" Then
        txtProveedor = ""
        txtFactura = ""
    End If
    
    lResponsable_PEPCECO = 0
    lDirectorProyecto_PEPCECO = 0
    sPreValidador = ""
    '10-12-19 Para los anticipos también tenemos que pedir pep_ceco
    'If CodPaisTrabajo = giPAIS_GES And Trim(txtConceptoGasto.Text) <> "Anticipo" Then
    If CodPaisTrabajo = giPAIS_GES Then
    
        txtPEPCECO = Trim(txtPEPCECO)
        If txtPEPCECO = "" Then
            imgAceptarLinea.Tag = ""
            txtPEPCECO.SetFocus
            Exit Sub
        Else
            If Not bComprobarPEPCECO(txtPEPCECO, CStr(miCodDepartamento), txtFecha, lResponsable_PEPCECO, lDirectorProyecto_PEPCECO, sPreValidador) Then
                imgAceptarLinea.Tag = ""
                MsgBox "El PEP/CECO/IO no existe o no es válido para este empleado.", vbApplicationModal + vbExclamation, "E R R O R"
                txtPEPCECO.SetFocus
                Exit Sub
            End If
        End If
    
    End If
    sCambioGuardadoPorUsuario = ""
    If txtCambio.Enabled And cmbDivisa.Text <> msMonedaDepartamento And cmbDivisa.Text <> "" Then
        txtCambio = Trim(txtCambio)
        If txtCambio = "" Then txtCambio = dDameCambioDivisa(CodPaisTrabajo, cmbDivisa.Text, txtFecha, msMonedaDepartamento)
        If txtCambio = "" Then
            imgAceptarLinea.Tag = ""
            txtCambio.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtCambio) Then
            imgAceptarLinea.Tag = ""
            txtCambio = ""
            txtCambio.SetFocus
            Exit Sub
        End If
        If txtCambio.Text <> txtCambio.Tag Then sCambioGuardadoPorUsuario = "N"
    End If
    sPEPCECO = UCase(Trim(txtPEPCECO))
    sCambio = Trim(txtCambio)
    If sCambio = "" Then sCambio = "1"
    sMoneda = cmbDivisa.Text
    If sMoneda = "" Then sMoneda = msMonedaDepartamento
    If sMoneda = msMonedaDepartamento Then sCambio = "1"
    
    sCuenta_Contable = ""
    If cmbCuentasContables.Text <> "" Then
        If cmbCuentasContables.ListIndex <> -1 Then
            sCuenta_Contable = Trim(Mid(cmbCuentasContables.Text, 1, InStr(1, cmbCuentasContables.Text, "-") - 1))
        Else
            If bCuentaContableValida(Trim(cmbCuentasContables.Text)) Then
                sCuenta_Contable = Trim(cmbCuentasContables.Text)
            Else
                imgAceptarLinea.Tag = ""
                cmbCuentasContables.ListIndex = -1
                cmbCuentasContables = ""
                MsgBox "La cuenta contable no es válida.", vbApplicationModal + vbExclamation, "Información"
                cmbCuentasContables.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    dImporte_Ori = txtImporte
    If sCambio = "1" Then
        dImporte = txtImporte
    Else
        dImporte = Round(CDbl(txtImporte) / CDbl(sCambio), 2)
    End If
    
    sDesglose_IVA = "N"
    dImporte_IVA = 0
    dBase_IVA = 0
    sTipo_IVA = ""
    If CodPaisTrabajo = giPAIS_GES And Trim(txtConceptoGasto.Text) <> "Anticipo" Then
    
        If chkDesgloseIVA.value = vbChecked Then
        
            sDesglose_IVA = "S"
            
            txtImporteIVA.Text = Trim(txtImporteIVA)
            If txtImporteIVA.Text = "" Then
                imgAceptarLinea.Tag = ""
                MsgBox "Introduzca un valor válido para el Importe IVA.", vbApplicationModal + vbExclamation, "E R R O R"
                txtImporteIVA.SetFocus
                Exit Sub
            End If
            If Not IsNumeric(txtImporteIVA.Text) Then
                imgAceptarLinea.Tag = ""
                MsgBox "Introduzca un valor válido para el Importe IVA.", vbApplicationModal + vbExclamation, "E R R O R"
                txtImporteIVA.Text = ""
                txtImporteIVA.SetFocus
                Exit Sub
            End If
'            If CDbl(txtImporteIVA.Text) > dImporte Then
'                imgAceptarLinea.Tag = ""
'                MsgBox "El Importe IVA no puede ser mayor que el Importe del gasto.", vbApplicationModal + vbExclamation, "E R R O R"
'                txtImporteIVA.Text = ""
'                txtImporteIVA.SetFocus
'                Exit Sub
'            End If
            dImporte_IVA = txtImporteIVA
            
            txtBaseIVA.Text = Trim(txtBaseIVA)
            If txtBaseIVA.Text = "" Then
                imgAceptarLinea.Tag = ""
                MsgBox "Introduzca un valor válido para la Base IVA.", vbApplicationModal + vbExclamation, "E R R O R"
                txtBaseIVA.SetFocus
                Exit Sub
            End If
            If Not IsNumeric(txtBaseIVA.Text) Then
                imgAceptarLinea.Tag = ""
                MsgBox "Introduzca un valor válido para la Base IVA.", vbApplicationModal + vbExclamation, "E R R O R"
                txtBaseIVA.Text = ""
                txtBaseIVA.SetFocus
                Exit Sub
            End If
            dBase_IVA = txtBaseIVA
            
            txtTipoIVA.Text = Trim(txtTipoIVA)
            If txtTipoIVA.Text = "" Then
                imgAceptarLinea.Tag = ""
                MsgBox "Introduzca un valor válido para el Tipo IVA.", vbApplicationModal + vbExclamation, "E R R O R"
                txtTipoIVA.SetFocus
                Exit Sub
            End If
            sTipo_IVA = txtTipoIVA
            
            '18-11-19 Indican desde GES que solo se podrá introducir un tipo de IVA para cada gasto.
'            dImporte_IVA2 = -1
'            dBase_IVA2 = -1
'            sTipo_IVA2 = ""
'            txtImporteIVA2 = Trim(txtImporteIVA2)
'            txtBaseIVA2 = Trim(txtBaseIVA2)
'            txtTipoIVA2 = Trim(txtTipoIVA2)
'            If txtImporteIVA2 <> "" Or txtBaseIVA2 <> "" Or txtTipoIVA2 <> "" Then
'
'                If txtImporteIVA2 <> "" And txtBaseIVA2 = "" Then
'                    imgAceptarLinea.Tag = ""
'                    MsgBox "Debe rellenar todos los campos IVA.", vbApplicationModal + vbExclamation, "E R R O R"
'                    txtBaseIVA2.SetFocus
'                    Exit Sub
'                End If
'                If txtImporteIVA2 <> "" And txtTipoIVA2 = "" Then
'                    imgAceptarLinea.Tag = ""
'                    MsgBox "Debe rellenar todos los campos IVA.", vbApplicationModal + vbExclamation, "E R R O R"
'                    txtTipoIVA2.SetFocus
'                    Exit Sub
'                End If
'                If txtBaseIVA2 <> "" And txtImporteIVA2 = "" Then
'                    imgAceptarLinea.Tag = ""
'                    MsgBox "Debe rellenar todos los campos IVA.", vbApplicationModal + vbExclamation, "E R R O R"
'                    txtImporteIVA2.SetFocus
'                    Exit Sub
'                End If
'                If txtBaseIVA2 <> "" And txtTipoIVA2 = "" Then
'                    imgAceptarLinea.Tag = ""
'                    MsgBox "Debe rellenar todos los campos IVA.", vbApplicationModal + vbExclamation, "E R R O R"
'                    txtTipoIVA2.SetFocus
'                    Exit Sub
'                End If
'                If txtTipoIVA2 <> "" And txtImporteIVA2 = "" Then
'                    imgAceptarLinea.Tag = ""
'                    MsgBox "Debe rellenar todos los campos IVA.", vbApplicationModal + vbExclamation, "E R R O R"
'                    txtImporteIVA2.SetFocus
'                    Exit Sub
'                End If
'                If txtTipoIVA2 <> "" And txtBaseIVA2 = "" Then
'                    imgAceptarLinea.Tag = ""
'                    MsgBox "Debe rellenar todos los campos IVA.", vbApplicationModal + vbExclamation, "E R R O R"
'                    txtBaseIVA2.SetFocus
'                    Exit Sub
'                End If
'
'                If Not IsNumeric(txtImporteIVA2.Text) Then
'                    imgAceptarLinea.Tag = ""
'                    MsgBox "Introduzca un valor válido para el Importe IVA.", vbApplicationModal + vbExclamation, "E R R O R"
'                    txtImporteIVA2.Text = ""
'                    txtImporteIVA2.SetFocus
'                    Exit Sub
'                End If
'                dImporte_IVA2 = txtImporteIVA2
'
'                If Not IsNumeric(txtBaseIVA2.Text) Then
'                    imgAceptarLinea.Tag = ""
'                    MsgBox "Introduzca un valor válido para la Base IVA.", vbApplicationModal + vbExclamation, "E R R O R"
'                    txtBaseIVA2.Text = ""
'                    txtBaseIVA2.SetFocus
'                    Exit Sub
'                End If
'                dBase_IVA2 = txtBaseIVA2
'
'                sTipo_IVA2 = txtTipoIVA2
'
'            End If
            
            '14-11-19 Tras comentarlo con MariMar, en principio no vamos a hacer esta comprobación
'            dAux = dImporte_IVA + dBase_IVA
'            If dImporte_IVA2 <> -1 Then
'                dAux = dAux + dImporte_IVA2 + dBase_IVA2
'            Else
'                dImporte_IVA2 = 0
'                dBase_IVA2 = 0
'                dTipo_IVA2 = 0
'            End If
'            If dAux <> dImporte_Ori Then
'                imgAceptarLinea.Tag = ""
'                MsgBox "El Importe del gasto no coincide con el desglose IVA.", vbApplicationModal + vbExclamation, "E R R O R"
'                txtImporteIVA.SetFocus
'                Exit Sub
'            End If
            
        End If
    
    End If
    
    sCadena = Replace(sCadena, "'", " ")
    sCadena = Replace(sCadena, """", " ")
    
    '19-06-18 En GES piden que si el gasto se inserta o modifica desde la aplicación no se marque como validado el gasto para que el responsable tenga que validar las modificaciones.
    sValidado = "S"
    sValidacion = "Aplic GastosInt"
    sEnv_Mail_Validacion = "N"
    If CodPaisTrabajo = giPAIS_GES Then
        sValidado = "N"
        sValidacion = ""
        sEnv_Mail_Validacion = "N"
    End If
    
    sJustificante = ""
    If lblFichero <> "" Then
        
        sExtension = lblFichero
        While InStr(1, sExtension, ".")
            sExtension = Mid(sExtension, InStr(1, sExtension, ".") + 1)
        Wend
        
        sJustificante = Format(Now, "yyyyMMddHHmmss")
        sRutaNAS = "\\nas\L\Intercambio\FicherosGV\"
        sRutaMarte = "C:\FicherosGV\" & Mid(sJustificante, 1, 4) & "\" & Mid(sJustificante, 5, 2) & "\"
        sJustificante = CodPaisTrabajo & "_" & miCodEmpleado & "_" & miCodDepartamento & "_" & sJustificante & "." & sExtension
        
        FileCopy lblFichero.Tag, sRutaNAS & sJustificante
        
    End If
    
    sValidador = ""
    iOfi_PreValidacion = 0
    'Para GES tenemos que calcular los campos Validador y Ofi_PreValidacion
    If CodPaisTrabajo = giPAIS_GES Then
    
        'Buscamos quien debe validar el gasto siguiendo las siguientes condiciones:
        '1 - Sí el usuario que introduce el gasto no es responsable del PEP/CECO al que se inputa, el validador será el responsable del PEP/CECO
        '2 - Sí es el responsable del PEP/CECO y no es el Director de Proyecto del PEP/CECO, el validador será el Director de Proyecto del PEP/CECO, siempre que tenga Director de Proyecto.
        '3 - Sí es el responsable y Director de Proyecto del PEP/CECO o el PEP/CECO no tiene director de proyecto, el validador será su superior jerárquico, su Manager.
        'El responsable y director proyecto del PEP/CECO lo hemos obtenido al comprobar la validez del PEP/CECO
        If miCodEmpleado = lResponsable_PEPCECO Then
            
            If miCodEmpleado = lDirectorProyecto_PEPCECO Then
                sValidador = miManagerEmpleado
            Else
                If lDirectorProyecto_PEPCECO > 0 Then
                    sValidador = lDirectorProyecto_PEPCECO
                Else
                    sValidador = miManagerEmpleado
                End If
            End If
            
        Else
            sValidador = lResponsable_PEPCECO
        End If
        
        If miOfi_PreValidacionEmpleado > 0 Then iOfi_PreValidacion = miOfi_PreValidacionEmpleado
        
        If sPreValidador <> "" Then iOfi_PreValidacion = -9999
    
    End If
    
    If lblContador = "" Then
    
        If mid_Ofi_PreValidacion > 0 Then
            imgAceptarLinea.Tag = ""
            MsgBox "No tiene permiso para insertar gastos.", vbApplicationModal + vbExclamation, "Información"
            RellenarDatosGastos
            Exit Sub
        End If
    
        sPagado = sTipoPagoGasto(txtTipoPago)
        'Sí es metálico, sPagado tendrá el valor "N", por lo que nos vale para indicar que aún no ha sido pagado, en cualquier otro caso lo marcamos como pagado ya que sería VISA, Facturado o SolRed
        '14-04-21 Como en GES se va a empezar a usar tarjeta como medio de pago, aquí para pasarlo posteriormente a SAP no podemos marcarlo como cobrado
        'If sPagado <> "N" Then sPagado = "S"
        If CodPaisTrabajo = giPAIS_GES Then
            sPagado = "N"
        Else
            If sPagado <> "N" Then sPagado = "S"
        End If
        '14-04-21 Fin
    
        sSql = "INSERT INTO Gastos_Internos(CodPais,Codigo,Departamento,Jerarquia,TipoGasto,Fecha,Importe,Numero,Numero2,Donde,Tipo_Com_Inv,VISA,Validado,Validacion," & _
               "Pagado,Moneda,Factura,FamiliaGasto,Proveedor,Personas,id_Empresa,Cambio,Importe_Ori,PEP_CECO,Env_Mail_Validacion,Archivado_En,Conta_SAP,ModificadoAdm," & _
               "Env_Mail_Modi_Adm,CIF_Proveedor,Cuenta_Contable,NombreFichero,Ruta_Ticket,Fecha_Insert,Proveedor_OK,id_Proveedor,Cambio_Guardado_Por_Usuario," & _
               "CPD_Nombre,CPD_CP,CPD_Ciudad,CPD_Pais,CPD_CIF,Desglose_IVA,Importe_IVA,Base_IVA,Tipo_IVA,Importe_IVA2,Base_IVA2,Tipo_IVA2,Validador," & _
               "Ofi_PreValidacion,Contador_Asociado,PreValidador) " & _
               "VALUES(" & CodPaisTrabajo & "," & miCodEmpleado & "," & miCodDepartamento & ",'" & msJerEmpleado & "'," & _
               sidGasto(msTipo, txtConceptoGasto) & ",'" & Replace(txtFecha, "/", "") & "'," & dImporte & "," & _
               iNumero & "," & iNumero2 & ",'" & sCadena & "','" & sTipoComidasInvitaciones(txtTipoComida) & "','" & sTipoPagoGasto(txtTipoPago) & _
               "','" & sValidado & "','" & sValidacion & "','" & sPagado & "','" & sMoneda & "','" & txtFactura & "','" & sFamiliaGasto(txtFamiliaGasto) & "','" & txtProveedor & _
               "','" & txtPersonas & "'," & cmbEmpresas.ItemData(cmbEmpresas.ListIndex) & "," & sCambio & "," & dImporte_Ori & ",'" & sPEPCECO & "','" & sEnv_Mail_Validacion & _
               "','','N','S','S','" & sCIF_Proveedor & "','" & sCuenta_Contable & "','" & sJustificante & "','" & sRutaMarte & sJustificante & "',CURRENT YEAR TO SECOND,'" & _
               sProveedor_OK & "','" & sid_Proveedor & "','N','" & sCPD_Nombre & "','" & sCPD_CP & "','" & sCPD_Ciudad & "','" & sCPD_Pais & "','" & sCPD_CIF & "','" & _
               sDesglose_IVA & "'," & dImporte_IVA & "," & dBase_IVA & ",'" & UCase(sTipo_IVA) & "'," & dImporte_IVA2 & "," & dBase_IVA2 & ",'" & UCase(sTipo_IVA2) & "','" & _
               sValidador & "'," & iOfi_PreValidacion & ",0,'" & sPreValidador & "')"
        If Not gBdGALILEO.bExecute(sSql) Then
            imgAceptarLinea.Tag = ""
            miPosicionarEnFila = -1
            MsgBox "Error al intentar guardar la información", vbApplicationModal + vbExclamation, "Información"
            Exit Sub
        End If

        '06-08-21 Para Plastiverd, Iqoxe e Iqlit en los gastos de KMs hay que añadir de forma automática otra línea de gasto de tipo varios para que se paguen los KMs a 0.30 en lugar de 0.19
        '         Comprobamos sí es un tipo de gasto KMs y alguno de los paises donde tenemos que hacer la operativa, buscamos el contador del gasto insertado e insertamos el gasto asociado
        If txtConceptoGasto.Text = "Kilómetros" And (CodPaisTrabajo = giPAIS_PLASTIVERD Or CodPaisTrabajo = giPAIS_IQOXE Or CodPaisTrabajo = giPAIS_IQLIT) Then
            
            lContadorAsociado = 0
            sSqlAux = "SELECT NVL(MAX(Contador),0) " & _
                      "FROM Gastos_Internos " & _
                      "WHERE CodPais = " & CodPaisTrabajo & " AND Codigo = " & miCodEmpleado & " AND Jerarquia = '" & msJerEmpleado & "' " & _
                      "AND Departamento = " & miCodDepartamento & " AND TipoGasto = " & sidGasto(msTipo, txtConceptoGasto) & " " & _
                      "AND Fecha = '" & Replace(txtFecha, "/", "") & "'"
            If gBdGALILEO.bRecordset(sSqlAux, afoRegAux) Then lContadorAsociado = NoNull(afoRegAux(0))
            
            '06-08-21 Insertamos un tipo de gasto "Otros gastos" cuyo importe será 0.11 (diferencia entre 0.30 menos 0.19) x nº KMs
            dImporte = 0.11 * iNumero
            sCadena = "Gastos manutencion sin comprobante"
            sSqlAux = "INSERT INTO Gastos_Internos(CodPais,Codigo,Departamento,Jerarquia,TipoGasto,Fecha,Importe,Numero,Numero2,Donde,Tipo_Com_Inv,VISA,Validado,Validacion," & _
                      "Pagado,Moneda,Factura,FamiliaGasto,Proveedor,Personas,id_Empresa,Cambio,Importe_Ori,PEP_CECO,Env_Mail_Validacion,Archivado_En,Conta_SAP,ModificadoAdm," & _
                      "Env_Mail_Modi_Adm,CIF_Proveedor,Cuenta_Contable,NombreFichero,Ruta_Ticket,Fecha_Insert,Proveedor_OK,id_Proveedor,Cambio_Guardado_Por_Usuario," & _
                      "CPD_Nombre,CPD_CP,CPD_Ciudad,CPD_Pais,CPD_CIF,Desglose_IVA,Importe_IVA,Base_IVA,Tipo_IVA,Importe_IVA2,Base_IVA2,Tipo_IVA2,Validador," & _
                      "Ofi_PreValidacion,Contador_Asociado,PreValidador) " & _
                      "VALUES(" & CodPaisTrabajo & "," & miCodEmpleado & "," & miCodDepartamento & ",'" & msJerEmpleado & "',10,'" & Replace(txtFecha, "/", "") & "'," & _
                      dImporte & ",0," & iNumero2 & ",'" & sCadena & "','" & sTipoComidasInvitaciones(txtTipoComida) & "','" & sTipoPagoGasto(txtTipoPago) & _
                      "','" & sValidado & "','" & sValidacion & "','" & sPagado & "','" & sMoneda & "','" & txtFactura & "','" & sFamiliaGasto(txtFamiliaGasto) & "','" & txtProveedor & _
                      "','" & txtPersonas & "'," & cmbEmpresas.ItemData(cmbEmpresas.ListIndex) & ",1," & dImporte & ",'" & sPEPCECO & "','" & sEnv_Mail_Validacion & _
                      "','','N','S','S','" & sCIF_Proveedor & "','" & sCuenta_Contable & "','" & sJustificante & "','" & sRutaMarte & sJustificante & "',CURRENT YEAR TO SECOND,'" & _
                      sProveedor_OK & "','" & sid_Proveedor & "','N','" & sCPD_Nombre & "','" & sCPD_CP & "','" & sCPD_Ciudad & "','" & sCPD_Pais & "','" & sCPD_CIF & "','" & _
                      sDesglose_IVA & "'," & dImporte_IVA & "," & dBase_IVA & ",'" & UCase(sTipo_IVA) & "'," & dImporte_IVA2 & "," & dBase_IVA2 & ",'" & UCase(sTipo_IVA2) & "','" & _
                      sValidador & "'," & iOfi_PreValidacion & "," & lContadorAsociado & ",'" & sPreValidador & "')"
            gBdGALILEO.bExecute (sSqlAux)
        
        End If
    
    Else
    
        sPagado = moParrilla.TextoLinCol(iFila, 24)
        'Sí no está pagado y es metálico, sPagado tendrá el valor "N", por lo que nos vale para indicar que aún no ha sido pagado, en cualquier otro caso lo marcamos como pagado ya que sería VISA, Facturado o SolRed
        '14-04-21 Como en GES se va a empezar a usar tarjeta como medio de pago, aquí para pasarlo posteriormente a SAP no podemos marcarlo como cobrado
        'If sPagado = "N" Then If sTipoPagoGasto(txtTipoPago) <> "N" Then sPagado = "S"
        If CodPaisTrabajo = giPAIS_GES Then
            sPagado = "N"
        Else
            If sPagado = "N" Then If sTipoPagoGasto(txtTipoPago) <> "N" Then sPagado = "S"
        End If
        
        '05-08-19 Solicitan introducir un motivo del cambio realizado para mostrarselo al usuario
        sMotivoModificacion = ""
        If CodPaisTrabajo = giPAIS_GES Then
        
            frmPedirMotivoModificacion.Show MODAL
            sMotivoModificacion = frmPedirMotivoModificacion.MotivoModificacion
        
        End If
        
        'Como vamos a modificar un gasto, procedmos a registrar en el log la línea que se va a modificar
        sSql = "INSERT INTO Log_Gastos_Internos " & _
               "SELECT CURRENT YEAR TO SECOND,'" & Trim(goUsuario.IdUsuario) & "',Contador,CodPais,Codigo,Departamento,Jerarquia," & _
               "TipoGasto,Tipo_Com_Inv,Fecha,Importe,Numero,Numero2,Donde,Visa,Validado,Validacion,Pagado,Moneda,Factura," & _
               "FamiliaGasto,Proveedor,Personas,Ruta_Ticket,Latitud,Longitud,NombreFichero,id_Empresa,Cambio,Importe_Ori," & _
               "PEP_CECO,Env_Mail_Validacion,Archivado_En,Conta_SAP,ModificadoAdm,Env_Mail_Modi_Adm,NVL(CIF_Proveedor,'') CIF_Proveedor,Fecha_Val_Rec,Mot_Rechazo," & _
               "NVL(Cuenta_Contable,'') Cuenta_Contable,'" & sMotivoModificacion & "' MotModi,NVL(Desglose_IVA,'N') Desglose_IVA," & _
               "NVL(Importe_IVA,0) Importe_IVA,NVL(Base_IVA,0) Base_IVA,NVL(Tipo_IVA,'') Tipo_IVA,NVL(Importe_IVA2,0) Importe_IVA2,NVL(Base_IVA2,0) Base_IVA2,NVL(Tipo_IVA2,'') Tipo_IVA2," & _
               "NVL(Validador,'') Validador,NVL(Ofi_PreValidacion,0) Ofi_PreValidacion,NVL(PreValidador,'') PreValidador " & _
               "FROM Gastos_Internos WHERE Contador = " & lblContador
        gBdGALILEO.bExecute sSql
                
        sSql = "UPDATE Gastos_Internos SET TipoGasto = " & sidGasto(msTipo, txtConceptoGasto) & ",Fecha = '" & Replace(txtFecha, "/", "") & "',Importe = " & dImporte & "," & _
               "Numero = " & iNumero & ",Numero2 = " & iNumero2 & ",Donde = '" & sCadena & "',Tipo_Com_Inv = '" & sTipoComidasInvitaciones(txtTipoComida) & "'," & _
               "VISA = '" & sTipoPagoGasto(txtTipoPago) & "',Factura = '" & txtFactura & "',FamiliaGasto = '" & sFamiliaGasto(txtFamiliaGasto) & "'," & _
               "Proveedor = '" & txtProveedor & "',Personas = '" & txtPersonas & "',Pagado = '" & sPagado & "',id_Empresa = " & cmbEmpresas.ItemData(cmbEmpresas.ListIndex) & "," & _
               "Importe_Ori = " & dImporte_Ori & ",PEP_CECO = '" & sPEPCECO & "',Moneda = '" & sMoneda & "',Cambio = " & sCambio & ",ModificadoAdm = 'S'," & _
               "CIF_Proveedor = '" & sCIF_Proveedor & "',Proveedor_OK = '" & sProveedor_OK & "',id_Proveedor = '" & sid_Proveedor & "',Cuenta_Contable = '" & sCuenta_Contable & "'," & _
               "CPD_Nombre = '" & sCPD_Nombre & "',CPD_CP = '" & sCPD_CP & "',CPD_Ciudad = '" & sCPD_Ciudad & "',CPD_Pais = '" & sCPD_Pais & "',CPD_CIF = '" & sCPD_CIF & "'," & _
               "Desglose_IVA = '" & sDesglose_IVA & "',Importe_IVA = " & dImporte_IVA & ",Base_IVA = " & dBase_IVA & ",Tipo_IVA = '" & UCase(sTipo_IVA) & "'," & _
               "Importe_IVA2 = " & dImporte_IVA2 & ",Base_IVA2 = " & dBase_IVA2 & ",Tipo_IVA2 = '" & UCase(sTipo_IVA2) & "'," & _
               "Validador = '" & sValidador & "',Ofi_PreValidacion = " & iOfi_PreValidacion & ",PreValidador = '" & sPreValidador & "'"
        '19-06-18 En GES piden que si el gasto se inserta o modifica desde la aplicación no se marque como validado el gasto para que el responsable tenga que validar las modificaciones.
        If CodPaisTrabajo = giPAIS_GES Then
            '29-11-18 Piden que solo se marque como que no está validado un gasto sí se modifica el Importe o el PEP/CECO. Igual para enviar el mail al trabajador como que se ha modificado desde administración (he movido de arriba el campo Env_Mail_Modi_Adm = 'N')
            '11-02-19 Piden que sí se cambia el tipo de cambio, no se tenga que volver a validar y no se informe al trabajador.
            '         Comparamos el Importe_ori en lugar del Importe, en ese campo el cambio no afecta. Y no mandamos información de que se ha modificado el gasto.
            sSqlAux = "SELECT Importe,Importe_Ori,PEP_CECO FROM Gastos_Internos WHERE Contador = " & lblContador
            If gBdGALILEO.bRecordset(sSqlAux, afoRegAux) Then
                If afoRegAux("Importe_Ori") <> dImporte_Ori Or LCase(Trim(afoRegAux("PEP_CECO"))) <> LCase(Trim(sPEPCECO)) Then
                    sSql = sSql & ",Env_Mail_Modi_Adm = 'N',Validado = '" & sValidado & "',Validacion = '" & sValidacion & "',Env_Mail_Validacion = '" & sEnv_Mail_Validacion & "'"
                End If
                
'''                Este era el código que había antes del cambio del día 11-02-19
'''                If afoRegAux("Importe") <> dImporte Or LCase(Trim(afoRegAux("PEP_CECO"))) <> LCase(Trim(sPEPCECO)) Then
'''                    sSql = sSql & ",Env_Mail_Modi_Adm = 'N',Validado = '" & sValidado & "',Validacion = '" & sValidacion & "',Env_Mail_Validacion = '" & sEnv_Mail_Validacion & "'"
'''                Else
'''                    sSql = sSql & ",Env_Mail_Modi_Adm = 'S'"
'''                End If
            End If
            Set afoRegAux = Nothing
        End If
        
        If sJustificante <> "" Then sSql = sSql & ",NombreFichero = '" & sJustificante & "',Ruta_Ticket = '" & sRutaMarte & sJustificante & "'"
        If sCambioGuardadoPorUsuario <> "" Then sSql = sSql & ",Cambio_Guardado_Por_Usuario = '" & sCambioGuardadoPorUsuario & "'"
        sSql = sSql & " WHERE Contador = " & lblContador
    
        miPosicionarEnFila = moParrilla.Grid.Row
        
        If Not gBdGALILEO.bExecute(sSql) Then
            imgAceptarLinea.Tag = ""
            miPosicionarEnFila = -1
            MsgBox "Error al intentar guardar la información", vbApplicationModal + vbExclamation, "Información"
            Exit Sub
        End If
        
        '06-08-21 Para Plastiverd, Iqoxe e Iqlit en los gastos de KMs hay que añadir de forma automática otra línea de gasto de tipo varios para que se paguen los KMs a 0.30 en lugar de 0.19
        '         Como estamos modificando un gasto, primero borramos un posible gasto generado automáticamente y asociado al gasto que estamos modificando y después comprobamos sí es un tipo
        '         de gasto KMs y alguno de los paises donde tenemos que hacer la operativa e insertamos el gasto asociado
        sSqlAux = "DELETE FROM Gastos_Internos WHERE Contador_Asociado = " & lblContador
        gBdGALILEO.bExecute (sSqlAux)
        
        If txtConceptoGasto.Text = "Kilómetros" And (CodPaisTrabajo = giPAIS_PLASTIVERD Or CodPaisTrabajo = giPAIS_IQOXE Or CodPaisTrabajo = giPAIS_IQLIT) Then
        
            dImporte = 0.11 * iNumero
            sCadena = "Gastos manutencion sin comprobante"
            sSqlAux = "INSERT INTO Gastos_Internos(CodPais,Codigo,Departamento,Jerarquia,TipoGasto,Fecha,Importe,Numero,Numero2,Donde,Tipo_Com_Inv,VISA,Validado,Validacion," & _
                      "Pagado,Moneda,Factura,FamiliaGasto,Proveedor,Personas,id_Empresa,Cambio,Importe_Ori,PEP_CECO,Env_Mail_Validacion,Archivado_En,Conta_SAP,ModificadoAdm," & _
                      "Env_Mail_Modi_Adm,CIF_Proveedor,Cuenta_Contable,NombreFichero,Ruta_Ticket,Fecha_Insert,Proveedor_OK,id_Proveedor,Cambio_Guardado_Por_Usuario," & _
                      "CPD_Nombre,CPD_CP,CPD_Ciudad,CPD_Pais,CPD_CIF,Desglose_IVA,Importe_IVA,Base_IVA,Tipo_IVA,Importe_IVA2,Base_IVA2,Tipo_IVA2,Validador," & _
                      "Ofi_PreValidacion,Contador_Asociado,PreValidador) " & _
                      "VALUES(" & CodPaisTrabajo & "," & miCodEmpleado & "," & miCodDepartamento & ",'" & msJerEmpleado & "',10,'" & Replace(txtFecha, "/", "") & "'," & _
                      dImporte & ",0," & iNumero2 & ",'" & sCadena & "','" & sTipoComidasInvitaciones(txtTipoComida) & "','" & sTipoPagoGasto(txtTipoPago) & _
                      "','" & sValidado & "','" & sValidacion & "','" & sPagado & "','" & sMoneda & "','" & txtFactura & "','" & sFamiliaGasto(txtFamiliaGasto) & "','" & txtProveedor & _
                      "','" & txtPersonas & "'," & cmbEmpresas.ItemData(cmbEmpresas.ListIndex) & ",1," & dImporte & ",'" & sPEPCECO & "','" & sEnv_Mail_Validacion & _
                      "','','N','S','S','" & sCIF_Proveedor & "','" & sCuenta_Contable & "','" & sJustificante & "','" & sRutaMarte & sJustificante & "',CURRENT YEAR TO SECOND,'" & _
                      sProveedor_OK & "','" & sid_Proveedor & "','N','" & sCPD_Nombre & "','" & sCPD_CP & "','" & sCPD_Ciudad & "','" & sCPD_Pais & "','" & sCPD_CIF & "','" & _
                      sDesglose_IVA & "'," & dImporte_IVA & "," & dBase_IVA & ",'" & UCase(sTipo_IVA) & "'," & dImporte_IVA2 & "," & dBase_IVA2 & ",'" & UCase(sTipo_IVA2) & "','" & _
                      sValidador & "'," & iOfi_PreValidacion & "," & lContadorAsociado & ",'" & sPreValidador & "')"
            gBdGALILEO.bExecute (sSqlAux)
        
        End If
        
    End If
    
    RellenarDatosGastos
    
Exit_Rutina:
    imgAceptarLinea.Tag = ""
    Exit Sub
Error_Rutina:
    imgAceptarLinea.Tag = ""
    miPosicionarEnFila = -1
    If Err.Number = 32755 Then Resume Exit_Rutina
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub imgAceptarLinea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgAdjuntar_Click()

    On Error GoTo Error_Rutina
    
    Me.MousePointer = DEFAULT
    
    lblFichero.Tag = ""
    lblFichero = ""
    
    MDIPrincipal.cdCuadroDialogo.CancelError = True
    MDIPrincipal.cdCuadroDialogo.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    MDIPrincipal.cdCuadroDialogo.Filter = "Imagen(*.jpg,*.jpeg,*.png,*.pdf)|*.jpg;*.jpeg;*.png;*.pdf"
    MDIPrincipal.cdCuadroDialogo.InitDir = sDameRutaEquipo(goUsuario.PaisTrabajo)
    MDIPrincipal.cdCuadroDialogo.ShowOpen
    
    lblFichero.Tag = MDIPrincipal.cdCuadroDialogo.FileName
    lblFichero = MDIPrincipal.cdCuadroDialogo.FileTitle
    
    If lblFichero.Tag <> "" Then
        If FileLen(lblFichero.Tag) > 3145728 Then
            MsgBox "El tamaño del fichero debe ser menor de 3MB", vbApplicationModal + vbExclamation, "E R R O R"
            lblFichero.Tag = ""
            lblFichero = ""
            Exit Sub
        End If
    End If

Exit_Rutina:
    Exit Sub
Error_Rutina:
    If Err.Number = 32755 Then Resume Exit_Rutina
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub imgAdjuntar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgArchivarEn_Click()
    
    Dim sArchivarEn As String
    Dim iFila As Integer
    Dim sContadores As String
    Dim sAux As String
    
    Me.MousePointer = DEFAULT
    If CodPaisTrabajo <> giPAIS_GES Then Exit Sub
    miPosicionarEnFila = -1
    
    sContadores = ""
    For iFila = 1 To moParrilla.Grid.Rows - 1
        If moParrilla.TextoLinCol(iFila, 2) = "-->" Then sContadores = sContadores & moParrilla.TextoLinCol(iFila, 0) & ","
    Next
    If sContadores = "" Then
        MsgBox "No hay seleccionado ningún gasto para archivar.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    sContadores = Mid(sContadores, 1, Len(sContadores) - 1)
    
    frmArchivarEn.Show MODAL
    sArchivarEn = Trim(frmArchivarEn.ArchivarEn)
    If sArchivarEn = "" Then Exit Sub
    
    sSql = "UPDATE Gastos_Internos SET Archivado_En = '" & sArchivarEn & "' WHERE Contador IN (" & sContadores & ")"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al intentar guardar los cambios solicitados.", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    
    MsgBox "Cambios realizados correctamente.", vbApplicationModal + vbInformation, "Información"
    RellenarDatosGastos

End Sub

Private Sub imgArchivarEn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgBorrarLinea_Click()
    
    If imgBorrarLinea.Tag <> "" Then Exit Sub
    imgBorrarLinea.Tag = "A"
    
    Me.MousePointer = DEFAULT
    If lblContador = "" Then
        imgBorrarLinea.Tag = ""
        LimpiarCampos
        Exit Sub
    End If
    
    'Para GES piden que no se puedan eliminar líneas
    If CodPaisTrabajo = giPAIS_GES Then
        
        If goUsuario.IdUsuario <> "lanchi" Then
            If Not goUsuario.Permiso(giOP_GVB_BORRARGASTOINTERNO) Then
                imgBorrarLinea.Tag = ""
                LimpiarCampos
                Exit Sub
            End If
        End If
        
        If MsgBox("Va a eliminar el gasto seleccionado, ¿está seguro?", vbApplicationModal + vbQuestion + vbYesNo, "Confirmación") = vbNo Then
            imgBorrarLinea.Tag = ""
            Exit Sub
        End If
        If lblContabilizadoSAP = "S" Then
            If MsgBox("El gasto ya se ha contabilizado en SAP, ¿está seguro de querer eliminarlo?", vbApplicationModal + vbQuestion + vbYesNo, "Confirmación") = vbNo Then
                imgBorrarLinea.Tag = ""
                Exit Sub
            End If
        End If
        If Not bEliminarGastoInterno_GES(lblContador) Then
            imgBorrarLinea.Tag = ""
            MsgBox "Se ha producido un error y no se ha podido borrar el gasto.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        moParrilla.ElimLinea
        LimpiarCampos
        imgBorrarLinea.Tag = ""
        Exit Sub
        
    End If
    
    If MsgBox("Va a eliminar el gasto seleccionado, ¿está seguro?", vbApplicationModal + vbQuestion + vbYesNo, "Confirmación") = vbNo Then
        imgBorrarLinea.Tag = ""
        Exit Sub
    End If
    
    If Not bEliminarGastoInterno(lblContador, msJerEmpleado) Then
        imgBorrarLinea.Tag = ""
        MsgBox "Se ha producido un error y no se ha podido borrar el gasto.", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    'moParrilla.ElimLinea
    RellenarDatosGastos
    LimpiarCampos
    imgBorrarLinea.Tag = ""
        
End Sub

Private Sub imgBorrarLinea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgBuscarUsuariosConGastosValidados_Click()
    
    Dim iCont As Integer
    
    If fraMensaje.Visible = True Then Exit Sub
    If fraPersonal.Enabled = False Then Exit Sub
    
    cmbDepartamentos_Click
    frmUsu_Gas_Pdte_Pago.Show MODAL
    
    If frmUsu_Gas_Pdte_Pago.Codigo <> -1 Then
    
        If frmUsu_Gas_Pdte_Pago.Baja Then chkBajas.value = vbChecked
        
        For iCont = 0 To cmbDepartamentos.ListCount - 1
            If cmbDepartamentos.ItemData(iCont) = frmUsu_Gas_Pdte_Pago.Departamento Then
                cmbDepartamentos.ListIndex = iCont
                Exit For
            End If
        Next
        
        For iCont = 0 To cmbEmpleados.ListCount - 1
            If cmbEmpleados.ItemData(iCont) = frmUsu_Gas_Pdte_Pago.Codigo Then
                cmbEmpleados.ListIndex = iCont
                Exit For
            End If
        Next
        
        dtpFechaDesde = frmUsu_Gas_Pdte_Pago.Fecha
        dtpFechaHasta = Now
        shFiltro_PendienteValidar.Visible = False
        shFiltro_Rechazado.Visible = False
        shFiltro_ValidadoNoPagado.Visible = True
        shFiltro_ValidadoPagado.Visible = False
        shFiltro_ContabilizadoSAP.Visible = False
        
        imgAceptarEmpleado_Click
        
    End If
    
End Sub

Private Sub imgBuscarUsuariosConGastosValidados_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgCancelarContabilizarSAP_Click()

    Dim iFila As Integer
    Dim sContadores As String
    Dim sAux As String

    On Error GoTo Error_Rutina
    
    Me.MousePointer = DEFAULT
    If CodPaisTrabajo <> giPAIS_GES Then Exit Sub
    If goUsuario.IdUsuario <> "lanchi" And goUsuario.IdUsuario <> "agonzalez" And goUsuario.IdUsuario <> "ivazquez" And goUsuario.IdUsuario <> "luciagil" And goUsuario.IdUsuario <> "inmamoya" And goUsuario.IdUsuario <> "ialcal" And goUsuario.IdUsuario <> "lvazquezmo" Then Exit Sub
    
    If imgCancelarContabilizarSAP.Tag <> "" Then Exit Sub
    imgCancelarContabilizarSAP.Tag = "C"
    
    Me.MousePointer = DEFAULT
    If CodPaisTrabajo <> giPAIS_GES Then Exit Sub
    miPosicionarEnFila = -1
    
    sContadores = ""
    For iFila = 1 To moParrilla.Grid.Rows - 1
        If moParrilla.TextoLinCol(iFila, 2) = "-->" And moParrilla.TextoLinCol(iFila, 26) = "S" Then sContadores = sContadores & moParrilla.TextoLinCol(iFila, 0) & ","
    Next
    If sContadores = "" Then
        imgCancelarContabilizarSAP.Tag = ""
        MsgBox "No hay seleccionado ningún gasto para descontabilizar de SAP.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    sContadores = Mid(sContadores, 1, Len(sContadores) - 1)
    
    If MsgBox("Va a marcar como NO CONTABILIZADOS EN SAP los gastos seleccionados, ¿está seguro?", vbApplicationModal + vbQuestion + vbYesNo, "Confirmación") = vbNo Then
        imgCancelarContabilizarSAP.Tag = ""
        Exit Sub
    End If
        
    sSql = "UPDATE Gastos_Internos SET Conta_SAP = 'N',Documento_SAP = '',Fecha_Conta_SAP = NULL,Ticket_En_SAP = 'N' WHERE Contador IN (" & sContadores & ")"
    If Not gBdGALILEO.bExecute(sSql) Then
        imgCancelarContabilizarSAP.Tag = ""
        MsgBox "Error al intentar guardar los cambios solicitados.", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    
    MsgBox "Cambios realizados correctamente.", vbApplicationModal + vbInformation, "Información"
    RellenarDatosGastos
    
Exit_Rutina:
    imgCancelarContabilizarSAP.Tag = ""
    Exit Sub
Error_Rutina:
    Me.MousePointer = DEFAULT
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub imgCancelarContabilizarSAP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgConsultaGastosFecha_Click()
    
    Me.MousePointer = DEFAULT
    miPosicionarEnFila = -1
    RellenarDatosGastos
    
End Sub

Private Sub imgConsultaGastosFecha_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgConsultar_Click()
    
    Dim sRuta As String
    Dim Formulario As Form
    
    Me.MousePointer = DEFAULT
    
    For Each Formulario In Forms
        If UCase(Formulario.Name) = "FRMEXPLORERV2" Then
            'Unload Formulario
            Formulario.Hide
            Exit For
        End If
    Next
    
    If lblContador = "" Then Exit Sub
    If Trim(LCase(txtConceptoGasto.Text)) = "anticipo" Then Exit Sub
    
    'lblRuta = Mid(lblRuta, 15)
    If Trim(lblRuta) = "" Then
        MsgBox "No hay asignado ningún ticket al gasto.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    
    sRuta = LCase(Trim(lblRuta))
    sRuta = Replace(sRuta, "\", "/")
    sRuta = Replace(sRuta, "c:/ficherosgv/", "")
    sRuta = "../Marte/" & sRuta
    
    If LCase(Mid(lblRuta, Len(lblRuta) - 3)) = ".pdf" Then
        'sRuta = "visortickets.cristianlay.com/" & lblRuta
        sRuta = "personal.clgrupoindustrial.com\Gastos\VisorTickets.asp?P=" & sRuta
    Else
        'sRuta = "visortickets.cristianlay.com/VisorTicket.asp?I=" & lblRuta
        sRuta = "personal.clgrupoindustrial.com\Gastos\VisorTickets.asp?I=" & sRuta
    End If
    frmExplorerV2.Create sRuta
    
    If chkEjecucionRemota.value = vbChecked Then
        frmExplorerV2.Left = 0
        frmExplorerV2.Top = 0
    Else
        frmExplorerV2.Left = Me.Width + 300
        frmExplorerV2.Top = Me.Top
    End If
    'frmExplorerV2.Left = Me.Width + 300
    'frmExplorerV2.Top = Me.Top
    frmExplorerV2.Show 'MODAL

End Sub

Private Sub imgConsultar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgContabilizarSAP_Click()
    
    Dim afoGasto As ADODB.Recordset
    
    Dim iFila As Integer
    Dim sContadores As String
    Dim sContador As String
    Dim iNumContadores As Integer
    
    Dim oZSAP As Object
    Dim oRetorno As Object
    Dim iRow As Integer
    Dim sSociedad As String
    Dim sFechaContab As String
    Dim sNumAsiento As String
    Dim sNumDocumento As String
    Dim sidPersonal As String
    Dim sProveedor As String
    Dim sTexto As String
    Dim sCuenta As String
    Dim sCECO As String
    Dim sPEP As String
    Dim sOrden As String
    Dim sImporte As String
    Dim sMoneda As String
    Dim sDesgloseIVA As String
    Dim sImporteIVA As String
    Dim sBaseIVA As String
    Dim sTipoIVA As String
    Dim sImporteIVA2 As String
    Dim sBaseIVA2 As String
    Dim sTipoIVA2 As String
    
    Dim sCPD_Nombre As String
    Dim sCPD_CP As String
    Dim sCPD_Ciudad As String
    Dim sCPD_Pais As String
    Dim sCPD_CIF As String
    
    Dim sSociedadSAP As String
    Dim sNumDocumentoSAP As String
    Dim sEjercicioSAP As String
    Dim sEjercicioAnticipo As String
    
    Dim sAux As String
    Dim bContabilizar As Boolean
    Dim bErrores As Boolean
    
    Dim bHayAnticiposParaContabilizar As Boolean
    
    Dim sTipoPago As String
    
    Dim sRetornoType As String
        
    On Error GoTo Error_Rutina
    
    Me.MousePointer = DEFAULT
    If CodPaisTrabajo <> giPAIS_GES Then Exit Sub
    miPosicionarEnFila = -1
        
    If imgContabilizarSAP.Tag <> "" Then Exit Sub
    imgContabilizarSAP.Tag = "C"
    
    If lblEmpleadoSinNumAcreedor.Visible = True Then
        MsgBox "El empleado no tiene número de acreedor en SAP por lo que no se pueden contabilizar sus gastos/anticipos.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    
    sContadores = ""
    iNumContadores = 0
    bHayAnticiposParaContabilizar = False
    For iFila = 1 To moParrilla.Grid.Rows - 1
        'Para GES el pago se marcará una vez contabilizado en SAP, por lo que no tengo que preguntar que esté pagado
        If CodPaisTrabajo = giPAIS_GES Then
            If moParrilla.TextoLinCol(iFila, 2) = "-->" And moParrilla.TextoLinCol(iFila, 23) = "S" And moParrilla.TextoLinCol(iFila, 26) = "N" Then
                sContadores = sContadores & moParrilla.TextoLinCol(iFila, 0) & ","
                iNumContadores = iNumContadores + 1
                If LCase(moParrilla.TextoLinCol(iFila, 4)) = "anticipo" Then bHayAnticiposParaContabilizar = True
            End If
        Else
            If moParrilla.TextoLinCol(iFila, 2) = "-->" And moParrilla.TextoLinCol(iFila, 23) = "S" And moParrilla.TextoLinCol(iFila, 24) = "S" And moParrilla.TextoLinCol(iFila, 26) = "N" Then sContadores = sContadores & moParrilla.TextoLinCol(iFila, 0) & ","
        End If
    Next
    If sContadores = "" Then
        MsgBox "No hay seleccionado ningún gasto que se pueda contabilizar en SAP.", vbApplicationModal + vbExclamation, "Información"
        imgContabilizarSAP.Tag = ""
        Exit Sub
    End If
    sContadores = Mid(sContadores, 1, Len(sContadores) - 1)
    
    If bHayAnticiposParaContabilizar And msAcreedorEmpleado = "" Then
        MsgBox "El empleado no tiene número de acreedor por lo que los anticipos no se pueden contabilizar.", vbApplicationModal + vbExclamation, "Información"
        imgContabilizarSAP.Tag = ""
        Exit Sub
    End If
    
    msFechaContaSAP = Format(Now, "yyyy-MM-dd")
    If chkPedirFechaContaSAP.value = vbChecked Then
        frmPedirFechaContaSAP.Show MODAL
        msFechaContaSAP = Format(frmPedirFechaContaSAP.FechaContaSAP, "yyyy-MM-dd")
    End If
    
    lblMensaje1 = ""
    lblMensaje2 = ""
    lblMensaje3 = ""
    fraMensaje.Visible = True
    lblMensaje1 = "Conectando con SAP"
    DoEvents
    
    Me.MousePointer = HOURGLASS
    fraDatos.Enabled = False
    If mbConexionSAP = False Then
        
        If goUsuario.IdUsuario = "lanchi" Then
            'mbConexionSAP = Logon("ges_calidad")
            mbConexionSAP = Logon("ges_desarrollo")
        Else
        mbConexionSAP = Logon("ges_produccion")
        End If
        
        If mbConexionSAP = False Then
            Me.MousePointer = DEFAULT
            fraMensaje.Visible = False
            fraDatos.Enabled = True
            MsgBox "No hay conexión a SAP, no se puede hacer la contabilización de los gastos seleccionados.", vbApplicationModal + vbExclamation, "Información"
            imgContabilizarSAP.Tag = ""
            Exit Sub
        End If
    End If
    
    Set oZSAP = oBapiControl.GetSAPObject("ZTRAVELEXP")

    sSociedad = Format(miCodDepartamento, "0000")
    sidPersonal = Format(miCodEmpleado, "00000000")
    bErrores = False
    iFila = 1
    While sContadores <> ""
    
        lblMensaje1 = "Contabilizando en SAP"
        lblMensaje2 = "Fila " & iFila & " de " & iNumContadores
        DoEvents
        iFila = iFila + 1
    
        If InStr(1, sContadores, ",") = 0 Then
            sContador = sContadores
            sContadores = ""
        Else
            sContador = Mid(sContadores, 1, InStr(1, sContadores, ",") - 1)
            sContadores = Mid(sContadores, InStr(1, sContadores, ",") + 1)
        End If
        
        If bDameDatosGastoInterno_Contador(CodPaisTrabajo, CLng(sContador), afoGasto) Then
        
            '20-08-19 Sí es un anticipo, la contabilización ha de hacerse de forma distinta
            
            If afoGasto("TipoGasto") = 99 Then
            'Es un anticipo
            '11-12-19 Según MariMar, no es necesario pasar los datos del Pep_Ceco
            
                bContabilizar = True
                        
                sFechaContab = msFechaContaSAP
                sEjercicioAnticipo = Mid(sFechaContab, 1, 4)
                If msAcreedorEmpleado = "" Then bContabilizar = False
                sMoneda = msMonedaDepartamento
                sImporte = afoGasto("Importe") * -1
                sTexto = Format(afoGasto("Fecha"), "dd-MM-yyyy") & Mid(" - ANTICIPO - " & Trim(afoGasto("Donde")), 1, 50)
                                                
                If bContabilizar Then
                
                    Set oZSAP = oBapiControl.GetSAPObject("ZTRAVELEXP")
                    Set oRetorno = Nothing
                    
                    oZSAP.ZbapiContabilizarAnticipos AWTYP:="BKPF", _
                                            GLVOR:="RFST", _
                                            BKTXT:="ANTICIPO GASTOS", _
                                            POSNR:="0000000001", _
                                            BUKRS:=sSociedad, _
                                            GJAHR:=sEjercicioAnticipo, _
                                            BUZEI:="001", _
                                            BLDAT:=sFechaContab, _
                                            BUDAT:=sFechaContab, _
                                            BLART:="KA", _
                                            XBLNR:="ANTICIPO GASTOS", _
                                            ZUMSK:="L", _
                                            UMSKZ:="F", _
                                            BSTAT:="S", _
                                            BSCHL:="39", _
                                            LIFNR:=msAcreedorEmpleado, _
                                            SGTXT:=sTexto, _
                                            WAERS:=sMoneda, _
                                            WRBTR:=sImporte, _
                                            RETURN:=oRetorno, _
                                            EBUKRS:=sSociedadSAP, _
                                            EBELNR:=sNumDocumentoSAP, _
                                            EGJAHR:=sEjercicioSAP
                                            
                    If oRetorno Is Nothing Then
                        bErrores = True
                    Else
                    
                        sRetornoType = ""
                        On Error Resume Next
                        sRetornoType = Trim(oRetorno("TYPE"))
                        Err.Clear
                        On Error GoTo Error_Rutina
                    
                        If sRetornoType = "E" Then
                            bErrores = True
                            MsgBox "Error al contabilizar el ANTICIPO '" & sTexto & "'." & vbCr & "ERROR: " & Trim(oRetorno("MESSAGE")), vbApplicationModal + vbExclamation, "E R R O R"
    '                        sAux = "ERROR/TYPE:" & Trim(oRetorno("TYPE")) & "/ID:" & Trim(oRetorno("ID")) & _
                                   "/NUMBER:" & Trim(oRetorno("NUMBER")) & "/MESSAGE:" & Trim(oRetorno("MESSAGE")) & "/LOG_NO:" & Trim(oRetorno("LOG_NO")) & _
                                   "/LOG_MSG_NO:" & Trim(oRetorno("LOG_MSG_NO")) & "/MESSAGE_V1:" & Trim(oRetorno("MESSAGE_V1")) & _
                                   "/MESSAGE_V2:" & Trim(oRetorno("MESSAGE_V2")) & "/MESSAGE_V3:" & Trim(oRetorno("MESSAGE_V3")) & _
                                   "/MESSAGE_V4:" & Trim(oRetorno("MESSAGE_V4")) & "/PARAMETER:" & Trim(oRetorno("PARAMETER")) & _
                                   "/ROW:" & Trim(oRetorno("ROW")) & "/FIELD:" & Trim(oRetorno("FIELD")) & "/SYSTEM:" & Trim(oRetorno("SYSTEM"))
                        Else
    '                        sAux = "OK/TYPE:" & Trim(oRetorno("TYPE")) & "/ID:" & Trim(oRetorno("ID")) & _
                                   "/NUMBER:" & Trim(oRetorno("NUMBER")) & "/MESSAGE:" & Trim(oRetorno("MESSAGE")) & "/LOG_NO:" & Trim(oRetorno("LOG_NO")) & _
                                   "/LOG_MSG_NO:" & Trim(oRetorno("LOG_MSG_NO")) & "/MESSAGE_V1:" & Trim(oRetorno("MESSAGE_V1")) & _
                                   "/MESSAGE_V2:" & Trim(oRetorno("MESSAGE_V2")) & "/MESSAGE_V3:" & Trim(oRetorno("MESSAGE_V3")) & _
                                   "/MESSAGE_V4:" & Trim(oRetorno("MESSAGE_V4")) & "/PARAMETER:" & Trim(oRetorno("PARAMETER")) & _
                                   "/ROW:" & Trim(oRetorno("ROW")) & "/FIELD:" & Trim(oRetorno("FIELD")) & "/SYSTEM:" & Trim(oRetorno("SYSTEM"))
                            '18-08-21 Si no nos devuelve documento desde SAP, lo consideramos erróneo.
                            If sNumDocumentoSAP = "" Then
                                bErrores = True
                                MsgBox "Error al contabilizar el ANTICIPO '" & sTexto & "'." & vbCr & "ERROR: SAP no ha devuelto documento.", vbApplicationModal + vbExclamation, "E R R O R"
                            Else
                                bMarcarGasto_ContabilizadoSAP CLng(sContador), sNumDocumentoSAP & sSociedadSAP & sEjercicioSAP, Mid(msFechaContaSAP, 9, 2) & Mid(msFechaContaSAP, 6, 2) & Mid(msFechaContaSAP, 1, 4)
                            End If
                            'If sNumDocumentoSAP = "" Then MsgBox "El anticipo con id " & sContador & " se ha contabilizado en SAP pero no ha devuelto documento.", vbApplicationModal + vbExclamation, "Información"
                            '18-08-21 Fin
                        End If
                    End If
                            
                End If
                
            Else
            'Es un gasto
            
                sTipoPago = "X"
                If NoNulo(afoGasto("Visa"), "T") = "S" Then sTipoPago = "V"
        
                bContabilizar = True
                
                sCPD_Nombre = ""
                sCPD_CP = ""
                sCPD_Ciudad = ""
                sCPD_Pais = ""
                sCPD_CIF = ""
                
                '13-11-19 Vemos sí el gasto lleva desglose de IVA
                sDesgloseIVA = Trim(NoNulo(afoGasto("Desglose_IVA")))
                If sDesgloseIVA = "" Then sDesgloseIVA = "N"
                sImporteIVA = NoNulo(afoGasto("Importe_IVA"), "N")
                sBaseIVA = NoNulo(afoGasto("Base_IVA"), "N")
                sTipoIVA = NoNulo(afoGasto("Tipo_IVA"), "T")
                sImporteIVA2 = NoNulo(afoGasto("Importe_IVA2"), "N")
                sBaseIVA2 = NoNulo(afoGasto("Base_IVA2"), "N")
                sTipoIVA2 = NoNulo(afoGasto("Tipo_IVA2"), "T")
                            
                '09-11-18 Piden tener la opcion de poder cambiar la fecha de contabilización
                'sFechaContab = Format(afoGasto("Fecha"), "yyyy-MM-dd")
                sFechaContab = msFechaContaSAP
                
                MsgBox "Fecha conta " & sFechaContab, vbApplicationModal + vbExclamation, "Información"
                
                sNumAsiento = ""
                sNumDocumento = ""
                sProveedor = ""
                If afoGasto("TipoGasto") = 11 Then sProveedor = Trim(afoGasto("idProveedor"))
                sTexto = Format(afoGasto("Fecha"), "dd-MM-yyyy") & " - " & Mid(Trim(afoGasto("NombreGasto")) & " - " & Trim(afoGasto("Donde")), 1, 50)
                sCuenta = ""
                Select Case afoGasto("TipoGasto")
                Case 1
                'Gasolina
                    sCuenta = "6280100000"
                Case 2
                'Comidas/Invitaciones
                    sCuenta = "6290100000"
                Case 3
                'Hoteles
                    sCuenta = "6290200000"
                Case 4
                'Teléfono
                    sCuenta = "6290500000"
                Case 5, 6
                'Billetes Avión/Tren
                    sCuenta = "6290000000"
                Case 7
                'Taxi
                    sCuenta = "6290300000"
                Case 8, 10
                'Peajes/Parking
                    sCuenta = "6290400000"
                Case 9
                'Otros gastos, sí no hay cuenta en el detalle del gasto le pongo la genérica indicada por Clausin
                    sCuenta = Trim(NoNull(afoGasto("Cuenta_Contable"), "A"))
                    If sCuenta = "" Or sCuenta = "0" Then sCuenta = "6291200000"
                Case 11
                'Vendor/Proveedor, no lleva cuenta contable puesto que se imputará a la cuenta contable del proveedor
                    
                    sCuenta = ""
                    'Sí el dato del proveedor no es correcto no puedo contabilizarlo
                    If Trim(NoNull(afoGasto("Proveedor_OK"), "T")) <> "S" Then
                        bContabilizar = False
                    Else
                        sTexto = Trim(afoGasto("Nombre_Proveedor")) & " - " & Format(afoGasto("Fecha"), "dd-MM-yyyy")
                    End If
                    
                    If LCase(sProveedor) = "cpd" Then
                        
                        sCPD_Nombre = Trim(NoNull(afoGasto("CPD_Nombre"), "T"))
                        sCPD_CP = Trim(NoNull(afoGasto("CPD_CP"), "T"))
                        sCPD_Ciudad = Trim(NoNull(afoGasto("CPD_Ciudad"), "T"))
                        sCPD_Pais = Trim(NoNull(afoGasto("CPD_Pais"), "T"))
                        sCPD_CIF = Trim(NoNull(afoGasto("CPD_CIF"), "T"))
                        
                        'El CIF no lo ponemos como obligatorio
                        If sCPD_Nombre = "" Or sCPD_CP = "" Or sCPD_Ciudad = "" Or sCPD_Pais = "" Then
                            bContabilizar = False
                        End If
                    
                    End If
                    
                End Select
                sCECO = ""
                sPEP = ""
                sOrden = ""
                Select Case Trim(NoNull(afoGasto("Tipo_PEPCECO"), "A"))
                Case "CC"
                    sCECO = Trim(afoGasto("PEP_CECO"))
                Case "PEP"
                    sPEP = Trim(afoGasto("PEP_CECO"))
                Case "ORD"
                    sOrden = Trim(afoGasto("PEP_CECO"))
                End Select
                'Preguntar a la gente de GES, el importe debe ir siempre en euros o en la moneda en la que haya guardado el gasto el empleado. De ser en moneda local, la bapi debería aceptar
                'también el campo cambio?
                'El 06-11-18 Mª Mar nos indica que el apunte en SAP debe hacerse en la moneda de la empresa a la que pertenezca el trabajador, en principio como vamos a trabajar en la zona Euro
                'todo debería ir en Euros. En el campo importe se guarda el importe original introducido por el trabajador aplicándole el cambio indicado bien por el propio trabajador, el de
                'sistema o el que introduzca el administrativo que utiliece la aplicación de Gastos para el chequeo y pago de los gastos. Por ello cojo el importe del campo Importe y la moneda
                'por ahora se la pongo a pelo al estar trabajando sólo con Euros. Quizás deberían indicarnos por sociedad qué tipo de moneda es la que tiene asignada para poder parametrizarlo
                sImporte = afoGasto("Importe")
                '13-11-19 Sí el gasto lleva desglose IVA, del importe del apunte del gasto debe ir restado el Importe IVA
                'If sDesgloseIVA = "S" Then sImporte = afoGasto("Importe") - CDbl(sImporte_IVA)
                'sMoneda = Trim(afoGasto("Moneda"))
                sMoneda = msMonedaDepartamento
                Set oRetorno = Nothing
                sSociedadSAP = ""
                sNumDocumentoSAP = ""
                sEjercicioSAP = ""
                
                If bContabilizar = True Then
                
                    If sProveedor = "" Then
                        
                        '15-03-19 Según indican desde GES para el usuario de José Luís Morlanes aunque el gasto no sea de proveedor hay que pasar a la bapi los datos CPD
                        If CodPaisTrabajo = giPAIS_GES And miCodDepartamento = 610 And miCodEmpleado = 9839 Then
                        
                            sProveedor = "CPD"
                            sCPD_Nombre = "JLM"
                            sCPD_CP = "00000"
                            sCPD_Ciudad = "X"
                            sCPD_Pais = "ES"
                            sCPD_CIF = ""
    
                            oZSAP.ZbapiExpensesPost sociedad:=sSociedad, _
                                             fechacontab:=sFechaContab, _
                                             numasiento:=sNumAsiento, _
                                             numdocumento:=sNumDocumento, _
                                             LIFNR:=sProveedor, _
                                             pernr:=sidPersonal, _
                                             texto:=sTexto, _
                                             Cuenta:=sCuenta, _
                                             ceco:=sCECO, _
                                             pep:=sPEP, _
                                             orden:=sOrden, _
                                             Importe:=sImporte, _
                                             moneda:=sMoneda, _
                                             Nombre:=sCPD_Nombre, _
                                             CP:=sCPD_CP, _
                                             Ciudad:=sCPD_Ciudad, _
                                             Pais:=sCPD_Pais, _
                                             CIF:=sCPD_CIF, _
                                             RETURN:=oRetorno, _
                                             EBUKRS:=sSociedadSAP, _
                                             EBELNR:=sNumDocumentoSAP, _
                                             EGJAHR:=sEjercicioSAP, _
                                             INDIVA1:=sTipoIVA, IMPORTEIVA1:=sImporteIVA, BASEIMPONIBLE1:=sBaseIVA, _
                                             ViaPago:=sTipoPago ', _
                                             INDIVA2:=sTipoIVA2, IMPORTEIVA2:=sImporteIVA2, BASEIMPONIBLE2:=sBaseIVA2
                                                                 
                        Else
                        
                            oZSAP.ZbapiExpensesPost sociedad:=sSociedad, _
                                             fechacontab:=sFechaContab, _
                                             numasiento:=sNumAsiento, _
                                             numdocumento:=sNumDocumento, _
                                             pernr:=sidPersonal, _
                                             texto:=sTexto, _
                                             Cuenta:=sCuenta, _
                                             ceco:=sCECO, _
                                             pep:=sPEP, _
                                             orden:=sOrden, _
                                             Importe:=sImporte, _
                                             moneda:=sMoneda, _
                                             RETURN:=oRetorno, _
                                             EBUKRS:=sSociedadSAP, _
                                             EBELNR:=sNumDocumentoSAP, _
                                             EGJAHR:=sEjercicioSAP, _
                                             INDIVA1:=sTipoIVA, IMPORTEIVA1:=sImporteIVA, BASEIMPONIBLE1:=sBaseIVA, _
                                             ViaPago:=sTipoPago ', _
                                             INDIVA2:=sTipoIVA2, IMPORTEIVA2:=sImporteIVA2, BASEIMPONIBLE2:=sBaseIVA2
                                             
                        End If
                        
                    Else
                                            
                        oZSAP.ZbapiExpensesPost sociedad:=sSociedad, _
                                        fechacontab:=sFechaContab, _
                                        numasiento:=sNumAsiento, _
                                        numdocumento:=sNumDocumento, _
                                        LIFNR:=sProveedor, _
                                        pernr:=sidPersonal, _
                                        texto:=sTexto, _
                                        Cuenta:=sCuenta, _
                                        ceco:=sCECO, _
                                        pep:=sPEP, _
                                        orden:=sOrden, _
                                        Importe:=sImporte, _
                                        moneda:=sMoneda, _
                                        Nombre:=sCPD_Nombre, _
                                        CP:=sCPD_CP, _
                                        Ciudad:=sCPD_Ciudad, _
                                        Pais:=sCPD_Pais, _
                                        CIF:=sCPD_CIF, _
                                        RETURN:=oRetorno, _
                                        EBUKRS:=sSociedadSAP, _
                                        EBELNR:=sNumDocumentoSAP, _
                                        EGJAHR:=sEjercicioSAP, _
                                        INDIVA1:=sTipoIVA, IMPORTEIVA1:=sImporteIVA, BASEIMPONIBLE1:=sBaseIVA, _
                                        ViaPago:=sTipoPago ', _
                                        INDIVA2:=sTipoIVA2, IMPORTEIVA2:=sImporteIVA2, BASEIMPONIBLE2:=sBaseIVA2
                                         
                    End If
                    If oRetorno Is Nothing Then
                        bErrores = True
                    Else
                    
                        sRetornoType = ""
                        On Error Resume Next
                        sRetornoType = Trim(oRetorno("TYPE"))
                        Err.Clear
                        On Error GoTo Error_Rutina
                    
                        If sRetornoType = "E" Then
                            bErrores = True
                            MsgBox "Error al contabilizar el gasto '" & sTexto & "'." & vbCr & "ERROR: " & Trim(oRetorno("MESSAGE")), vbApplicationModal + vbExclamation, "E R R O R"
    '                        sAux = "ERROR/TYPE:" & Trim(oRetorno("TYPE")) & "/ID:" & Trim(oRetorno("ID")) & _
                                   "/NUMBER:" & Trim(oRetorno("NUMBER")) & "/MESSAGE:" & Trim(oRetorno("MESSAGE")) & "/LOG_NO:" & Trim(oRetorno("LOG_NO")) & _
                                   "/LOG_MSG_NO:" & Trim(oRetorno("LOG_MSG_NO")) & "/MESSAGE_V1:" & Trim(oRetorno("MESSAGE_V1")) & _
                                   "/MESSAGE_V2:" & Trim(oRetorno("MESSAGE_V2")) & "/MESSAGE_V3:" & Trim(oRetorno("MESSAGE_V3")) & _
                                   "/MESSAGE_V4:" & Trim(oRetorno("MESSAGE_V4")) & "/PARAMETER:" & Trim(oRetorno("PARAMETER")) & _
                                   "/ROW:" & Trim(oRetorno("ROW")) & "/FIELD:" & Trim(oRetorno("FIELD")) & "/SYSTEM:" & Trim(oRetorno("SYSTEM"))
                        Else
    '                        sAux = "OK/TYPE:" & Trim(oRetorno("TYPE")) & "/ID:" & Trim(oRetorno("ID")) & _
                                   "/NUMBER:" & Trim(oRetorno("NUMBER")) & "/MESSAGE:" & Trim(oRetorno("MESSAGE")) & "/LOG_NO:" & Trim(oRetorno("LOG_NO")) & _
                                   "/LOG_MSG_NO:" & Trim(oRetorno("LOG_MSG_NO")) & "/MESSAGE_V1:" & Trim(oRetorno("MESSAGE_V1")) & _
                                   "/MESSAGE_V2:" & Trim(oRetorno("MESSAGE_V2")) & "/MESSAGE_V3:" & Trim(oRetorno("MESSAGE_V3")) & _
                                   "/MESSAGE_V4:" & Trim(oRetorno("MESSAGE_V4")) & "/PARAMETER:" & Trim(oRetorno("PARAMETER")) & _
                                   "/ROW:" & Trim(oRetorno("ROW")) & "/FIELD:" & Trim(oRetorno("FIELD")) & "/SYSTEM:" & Trim(oRetorno("SYSTEM"))
                            
                            'msFechaContaSAP = Format(Now, "yyyy-MM-dd")
                            'bMarcarGasto_ContabilizadoSAP CLng(sContador), Trim(oRetorno("MESSAGE_V2")), Mid(msFechaContaSAP, 9, 2) & Mid(msFechaContaSAP, 6, 2) & Mid(msFechaContaSAP, 1, 4)
                            
                            '26-11-19 Aunque no haya devuelto error, nos hemos encontrado casos en los que alguna variable viene a nulo y está originando problemas a la hora de la gestión.
                            '         Vamos a poner otra comprobación para avisar al usuario y que no genere errores
                            If IsNull(sNumDocumentoSAP) Or IsNull(sSociedadSAP) Or IsNull(sEjercicioSAP) Then
                                                                
                                'MsgBox "SAP informa de que el gasto se ha contabilizado pero ha devuelto datos erróneos, compruebe la contabilización y pongase en contacto con Sistemas.", vbApplicationModal + vbCritical, "Información"
                                MsgBox "Error al contabilizar el gasto '" & sTexto & "'." & vbCr & "ERROR: SAP no ha devuelto documento.", vbApplicationModal + vbExclamation, "E R R O R"
                                bErrores = True
                                
                            Else
                            
                                '18-08-21 Si no nos devuelve documento desde SAP, lo consideramos erróneo.
                                If sNumDocumentoSAP = "" Then
                                    bErrores = True
                                    MsgBox "Error al contabilizar el gasto '" & sTexto & "'." & vbCr & "ERROR: SAP no ha devuelto documento.", vbApplicationModal + vbExclamation, "E R R O R"
                                Else
                                    bMarcarGasto_ContabilizadoSAP CLng(sContador), sNumDocumentoSAP & sSociedadSAP & sEjercicioSAP, Mid(msFechaContaSAP, 9, 2) & Mid(msFechaContaSAP, 6, 2) & Mid(msFechaContaSAP, 1, 4)
                                'If Trim(oRetorno("MESSAGE_V2")) = "" Then MsgBox "El gasto con id " & sContador & " se ha contabilizado en SAP pero no ha devuelto documento.", vbApplicationModal + vbExclamation, "Información"
                                'If sNumDocumentoSAP = "" Then MsgBox "El gasto con id " & sContador & " se ha contabilizado en SAP pero no ha devuelto documento.", vbApplicationModal + vbExclamation, "Información"
                                End If
                                '18-08-21 Fin
                            End If
                            
                        End If
                    
                    End If
                    
                End If
                
            End If
        
        Else
            bErrores = True
        End If
        Set afoGasto = Nothing
        
    Wend
    If bErrores = True Then
        Me.MousePointer = DEFAULT
        MsgBox "Se han producido errores al intentar contabilizar alguno de los gastos seleccionados.", vbApplicationModal + vbExclamation, "E R R O R"
    End If
        
    fraMensaje.Visible = False
    lblMensaje1 = ""
    lblMensaje2 = ""
    lblMensaje3 = ""
    DoEvents
    fraDatos.Enabled = True
    RellenarDatosGastos
    Me.MousePointer = DEFAULT
    
Exit_Rutina:
    imgContabilizarSAP.Tag = ""
    Exit Sub
Error_Rutina:
    Me.MousePointer = DEFAULT
    fraMensaje.Visible = False
    fraDatos.Enabled = True
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    MsgBox "Antes de repetir la contabilización SAP de los gastos seleccionados, compruebe que realmente NO se ha realizado en SAP.", vbApplicationModal + vbCritical, "Información"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub imgContabilizarSAP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgDatos_ProveedorCPD_Click()

    Me.MousePointer = DEFAULT
    frmDatosProveedorCPD.Create lblCPD_Nombre, lblCPD_CP, lblCPD_Ciudad, lblCPD_Pais, lblCPD_CIF
    frmDatosProveedorCPD.Show MODAL
    lblCPD_Nombre = frmDatosProveedorCPD.Nombre
    lblCPD_CP = frmDatosProveedorCPD.CP
    lblCPD_Ciudad = frmDatosProveedorCPD.Ciudad
    lblCPD_Pais = frmDatosProveedorCPD.Pais
    lblCPD_CIF = frmDatosProveedorCPD.CIF
    
End Sub

Private Sub imgDatos_ProveedorCPD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgDesmarcarTodas_Click()

    Me.MousePointer = DEFAULT
    MarcarTodas False
    
End Sub

Private Sub imgDesmarcarTodas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgExportar_Click()
    Me.MousePointer = DEFAULT
    LimpiarCampos
    moParrilla.SituarEn 0, 1
    If moParrilla.bLineaActualVacia Then Exit Sub
    moParrilla.aClipBoard 1
    MsgBox "Información copiada al Portapapeles.", vbApplicationModal + vbInformation, "Información"
End Sub

Private Sub imgExportar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgImprimirPagos_Click()

    Me.MousePointer = DEFAULT
    If CodPaisTrabajo = giPAIS_GES Then
        ImprimirPagosGES
        Exit Sub
    End If
    ImprimirPagos

End Sub

Private Sub imgImprimirPagos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgInforme_Click()
    
    If fraMensaje.Visible = True Then Exit Sub
    If fraPersonal.Enabled = False Then Exit Sub
    
    Me.MousePointer = DEFAULT
    Me.Visible = False
    frmInfGastosPersonalGES.Show MODAL
    Me.Visible = True
    
End Sub

Private Sub imgInforme_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If fraMensaje.Visible = True Then Exit Sub
    If fraPersonal.Enabled = False Then Exit Sub

    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
    
End Sub

Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
    Call ReleaseCapture
    Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub imgLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgMarcarTodas_Click()

    Me.MousePointer = DEFAULT
    MarcarTodas True
    
End Sub

Private Sub imgMarcarTodas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgModificado_Click()
    
    Dim afoRegHistorico As ADODB.Recordset
    
    Dim sFecha As String
    Dim sConceptoGasto As String
    Dim sTipoGasto As String
    Dim sNomGasto As String
    Dim sEntrada As String
    Dim sDonde As String
    Dim sTipoComida As String
    Dim sNumNoches As String
    Dim sKms As String
    Dim sNumTickets As String
    Dim sNumPersonas As String
    Dim sImporte As String
    Dim sImporte2 As String
    Dim sProveedor As String
    Dim sFactura As String
    Dim sValidado As String
    Dim sPagado As String
    Dim sPersonas As String
    Dim sEmpresa As String
    Dim sImporte_EUR As String
    Dim sPEPCECO As String
    Dim sDivisa As String
    Dim sCambio As String
    Dim sArchivadoEn As String
    Dim sContaSAP As String
    Dim sFechaModi As String
    Dim sUsuModi As String
    Dim sCIF_Proveedor As String
    Dim sFecha_Val_Rec As String
    Dim sMotivo_Rechazo As String
    Dim sCuenta_Contable As String
    Dim sMotivoModi As String
    Dim sDesglose_IVA As String
    Dim sImporte_IVA As String
    Dim sBase_IVA As String
    Dim sTipo_IVA As String
    Dim sImporte_IVA2 As String
    Dim sBase_IVA2 As String
    Dim sTipo_IVA2 As String
    
    Me.MousePointer = DEFAULT
    If lblContador.Caption = "" Then Exit Sub
    
    If bDameDatosHistoricosGastoInterno(CLng(lblContador.Caption), afoRegHistorico) Then
    
        moParrillaHistorico.Vaciar
        moParrillaHistorico.SituarEn 0, 1
        fraDatos.Enabled = False
        fraHistorico.Visible = True
    
        While Not afoRegHistorico.EOF
        
            sFechaModi = Format(afoRegHistorico("Fecha_Modi"), "dd/MM/yy HH:mm:ss")
            sUsuModi = Trim(afoRegHistorico("Usu_Modi"))
            sFecha = Format(afoRegHistorico("Fecha"), "dd/MM/yyyy")
            sConceptoGasto = Trim(NoNull(afoRegHistorico("TipoGasto"), "T"))
            sNomGasto = sNombreGasto(msTipo, sConceptoGasto)
            sTipoGasto = Trim(NoNull(afoRegHistorico("FamiliaGasto"), "T"))
            If sTipoGasto = "" Then sTipoGasto = "O"
            sEntrada = Trim(NoNull(afoRegHistorico("Visa"), "T"))
            If sEntrada = "" Then sEntrada = "N"
            sDonde = Trim(NoNull(afoRegHistorico("Donde"), "T"))
            sTipoComida = Trim(NoNull(afoRegHistorico("Tipo_Com_Inv"), "T"))
            sNumNoches = Trim(NoNull(afoRegHistorico("Numero"), "T"))
            sKms = Trim(NoNull(afoRegHistorico("Numero"), "T"))
            sNumTickets = Trim(NoNull(afoRegHistorico("Numero"), "T"))
            sNumPersonas = Trim(NoNull(afoRegHistorico("Numero"), "T"))
            'Si son Hoteles, el número de personas viene numero2
            If sNomGasto = "Hoteles" Then sNumPersonas = Trim(NoNull(afoRegHistorico("Numero2"), "T"))
            '14-06-18
            'sImporte = Trim(NoNull(mafoRegistro("Importe"), "N"))
            sImporte = Trim(NoNull(afoRegHistorico("Importe_Ori"), "N"))
            sImporte2 = sImporte
            sProveedor = Trim(NoNull(afoRegHistorico("Proveedor"), "T"))
            sFactura = Trim(NoNull(afoRegHistorico("Factura"), "T"))
            sValidado = Trim(NoNull(afoRegHistorico("Validado"), "T"))
            If sValidado = "" Then sValidado = "N"
            sPagado = Trim(NoNull(afoRegHistorico("Pagado"), "T"))
            If sPagado = "" Then sPagado = "N"
            sPersonas = Trim(NoNull(afoRegHistorico("Personas"), "T"))
            sEmpresa = Trim(NoNull(afoRegHistorico("Empresa"), "T"))
            sImporte_EUR = Trim(NoNull(afoRegHistorico("Importe"), "N"))
            sPEPCECO = Trim(NoNull(afoRegHistorico("PEP_CECO"), "T"))
            sDivisa = Trim(NoNull(afoRegHistorico("Moneda"), "T"))
            sCambio = Trim(NoNull(afoRegHistorico("Cambio"), "N"))
            sArchivadoEn = Trim(NoNull(afoRegHistorico("Archivado_En"), "T"))
            sContaSAP = Trim(NoNull(afoRegHistorico("Conta_SAP"), "T"))
            If sContaSAP = "" Then sContaSAP = "N"
            sCIF_Proveedor = Trim(NoNull(afoRegHistorico("CIF_Proveedor"), "T"))
            sFecha_Val_Rec = ""
            If Not IsNull(afoRegHistorico("Fecha_Val_Rec")) Then sFecha_Val_Rec = Format(afoRegHistorico("Fecha_Val_Rec"), "dd/MM/yyyy")
            sMotivo_Rechazo = Trim(NoNull(afoRegHistorico("Motivo_Rechazo"), "T"))
            sCuenta_Contable = Trim(NoNull(afoRegHistorico("Cuenta_Contable"), "T"))
            sMotivoModi = Trim(NoNull(afoRegHistorico("Motivo_Modificacion"), "T"))
            sDesglose_IVA = Trim(NoNull(afoRegHistorico("Desglose_IVA"), "T"))
            sImporte_IVA = Trim(NoNull(afoRegHistorico("Importe_IVA"), "N"))
            sBase_IVA = Trim(NoNull(afoRegHistorico("Base_IVA"), "N"))
            sTipo_IVA = Trim(NoNull(afoRegHistorico("Tipo_IVA"), "N"))
            sImporte_IVA2 = Trim(NoNull(afoRegHistorico("Importe_IVA2"), "N"))
            sBase_IVA2 = Trim(NoNull(afoRegHistorico("Base_IVA2"), "N"))
            sTipo_IVA2 = Trim(NoNull(afoRegHistorico("Tipo_IVA2"), "N"))
            
            Select Case sNomGasto
            Case "Kilómetros"
                sNumNoches = "0"
                sNumTickets = "0"
                sNumPersonas = "0"
                sImporte = sImporte * sKms
            Case "Comidas-Invitaciones", "Comidas/Invitaciones"
                sNumNoches = "0"
                sKms = "0"
                sNumTickets = "0"
            Case "Hoteles"
                sKms = "0"
                sNumTickets = "0"
            Case "Peaje", "Peajes"
                sNumNoches = "0"
                sKms = "0"
                sNumPersonas = "0"
            End Select
                        
            moParrillaHistorico.TextoFila moParrillaHistorico.LineaLibre, "", sFechaModi, sUsuModi, sFecha, sNomGasto, sFamiliaGasto(sTipoGasto), sTipoPagoGasto(sEntrada), sDonde, sTipoComidasInvitaciones(sTipoComida), sNumNoches, sKms, sNumTickets, sNumPersonas, sImporte, sProveedor, sFactura, sPersonas, sEmpresa, sPEPCECO, sDivisa, sCambio, sImporte_EUR, sArchivadoEn, sCIF_Proveedor, sValidado, sPagado, sContaSAP, sFecha_Val_Rec, sMotivo_Rechazo, sCuenta_Contable, sDesglose_IVA, sImporte_IVA, sBase_IVA, sTipo_IVA, sImporte_IVA2, sBase_IVA2, sTipo_IVA2, sMotivoModi
            
            afoRegHistorico.MoveNext
        
        Wend
    
    End If
    
    Set afoRegHistorico = Nothing
    
End Sub

Private Sub imgModificado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgNuevoEmpleado_Click()

    Me.MousePointer = DEFAULT
    If fraHistorico.Visible Then Exit Sub
    If fraMensaje.Visible = True Then Exit Sub
    If fraDatos.Enabled = False And fraPersonal.Enabled = False Then Exit Sub
    
    miColumnaOrdenacion = -1
    miPosicionarEnFila = -1
    
    'cmbDepartamentos.ListIndex = -1
    'miCodDepartamento = -1
    
    EmpleadoDesBloquear CodPaisTrabajo, miCodDepartamento, miCodEmpleado, msJerEmpleado
    cmbEmpleados.ListIndex = -1
    miCodEmpleado = -1
    msJerEmpleado = ""
    lblEmpleadoSinNumAcreedor.Visible = False
    
    moParrilla.Vaciar
    moParrilla.SituarEn 0, 1
    LimpiarCampos
    shFiltro_PendienteValidar.Visible = False
    shFiltro_Rechazado.Visible = False
    shFiltro_ValidadoNoPagado.Visible = False
    shFiltro_ValidadoPagado.Visible = False
    shFiltro_ContabilizadoSAP.Visible = False
    
    fraDatos.Enabled = False
    fraPersonal.Enabled = True
    'cmbDepartamentos.SetFocus
    
    If mid_Ofi_PreValidacion > 0 Then cmbDepartamentos_Click
    
    cmbEmpleados.SetFocus
    
End Sub

Private Sub imgNuevoEmpleado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgPagar_Click()

    Dim iFila As Integer
    Dim sContadores As String
    Dim sAux As String
    
    Me.MousePointer = DEFAULT
    
    miPosicionarEnFila = -1
    
    sContadores = ""
    For iFila = 1 To moParrilla.Grid.Rows - 1
        If moParrilla.TextoLinCol(iFila, 2) = "-->" And moParrilla.TextoLinCol(iFila, 23) = "S" And moParrilla.TextoLinCol(iFila, 24) = "N" Then sContadores = sContadores & moParrilla.TextoLinCol(iFila, 0) & ","
    Next
    If sContadores <> "" Then
        
        sContadores = Mid(sContadores, 1, Len(sContadores) - 1)
        gBdGALILEO.BeginTrans True
        While sContadores <> ""
        
            If InStr(1, sContadores, ",") Then
                sAux = Mid(sContadores, 1, InStr(1, sContadores, ",") - 1)
                sContadores = Mid(sContadores, InStr(1, sContadores, ",") + 1)
            Else
                sAux = sContadores
                sContadores = ""
            End If
            If Not bMarcarGastosInternosPagados(sAux) Then
                gBdGALILEO.RollbackTrans
                MsgBox "Se ha producido un error al intentar marcar los gastos seleccionados como pagados.", vbApplicationModal + vbExclamation, "E R R O R"
                Exit Sub
            End If
            
        Wend
        gBdGALILEO.CommitTrans
        If MsgBox("Cambios realizados correctamente, ¿desea imprimir el informe de pago?.", vbApplicationModal + vbQuestion + vbYesNo, "Confirmación") = vbYes Then ImprimirPagos True
        RellenarDatosGastos
        
    Else
        MsgBox "No hay seleccionado ningún gasto pendiente de ser pagado.", vbApplicationModal + vbExclamation, "Información"
    End If

End Sub

Private Sub ImprimirPagosGES()

    Dim iFila As Integer
    Dim sFilas As String
    Dim sAux As String
    Dim sAux2 As String
    
    Dim dY As Double
    
    Dim dTotal As Double
    Dim dTotalOrdinario As Double
    Dim dTotalFormacion As Double
    Dim dTotalPresentacion As Double
    
    On Error GoTo Error_Rutina
    
    Me.MousePointer = DEFAULT
    
    sFilas = ""
    For iFila = 1 To moParrilla.Grid.Rows - 1
        'veo que el gasto esté contabilizado
        If moParrilla.TextoLinCol(iFila, 2) = "-->" And moParrilla.TextoLinCol(iFila, 26) = "S" Then sFilas = sFilas & iFila & ","
    Next
    
    If sFilas <> "" Then
    
        Me.MousePointer = HOURGLASS
        
        sFilas = Mid(sFilas, 1, Len(sFilas) - 1)
        
        Printer.Orientation = vbPRORLandscape
        Printer.ScaleMode = vbMillimeters
        
        Printer.Font = "Courier New"
        Printer.FontSize = 18
        Printer.FontBold = True
        Printer.CurrentX = 120
        Printer.CurrentY = 15
        Printer.Print "Gastos/Anticipos contabilizados"

        Printer.FontBold = True
        Printer.FontSize = 12
        Printer.CurrentY = 30
        Printer.CurrentX = 20
        Printer.Print "Gastos/Anticipos de " & cmbEmpleados.Text & " (" & cmbDepartamentos.Text & ")"
                
        Printer.FontUnderline = True
        Printer.FontSize = 9
        Printer.CurrentY = 40
        Printer.CurrentX = 5
        Printer.Print "Fecha"
        Printer.CurrentY = 40
        Printer.CurrentX = 25
        Printer.Print "T. Gasto"
        Printer.CurrentY = 40
        Printer.CurrentX = 55
        Printer.Print "Donde / Itinerario / Concepto"
        Printer.CurrentY = 40
        Printer.CurrentX = 135
        Printer.Print "Num."
        Printer.CurrentY = 40
        Printer.CurrentX = 150
        Printer.Print "Importe"
        Printer.CurrentY = 40
        Printer.CurrentX = 170
        Printer.Print "Imp. ML/Div./Cam."
        Printer.CurrentY = 40
        Printer.CurrentX = 210
        Printer.Print "PEP/CECO/IO"
        Printer.CurrentY = 40
        Printer.CurrentX = 240
        Printer.Print "Documento/F.Conta SAP"
        Printer.FontBold = False
        Printer.FontUnderline = False

        dY = 45
        dTotal = 0
        dTotalOrdinario = 0
        dTotalFormacion = 0
        dTotalPresentacion = 0
        While sFilas <> ""
        
            If InStr(1, sFilas, ",") Then
                iFila = Mid(sFilas, 1, InStr(1, sFilas, ",") - 1)
                sFilas = Mid(sFilas, InStr(1, sFilas, ",") + 1)
            Else
                iFila = sFilas
                sFilas = ""
            End If
            
            sAux = ""
            Select Case moParrilla.TextoLinCol(iFila, 4)
            Case "Kilómetros", "Kilometros"
                sAux = moParrilla.TextoLinCol(iFila, 10)
            Case "Comidas-Invitaciones", "Comidas/Invitaciones"
                sAux = moParrilla.TextoLinCol(iFila, 12)
            Case "Hoteles"
                sAux = moParrilla.TextoLinCol(iFila, 9) & "/" & moParrilla.TextoLinCol(iFila, 12)
            Case "Peaje", "Peajes"
                sAux = moParrilla.TextoLinCol(iFila, 11)
            End Select
            sAux2 = Mid(Trim(moParrilla.TextoLinCol(iFila, 7)), 1, 50)
        
            Printer.FontSize = 7
            Printer.CurrentY = dY
            Printer.CurrentX = 5
            Printer.Print moParrilla.TextoLinCol(iFila, 3)
            Printer.CurrentY = dY
            Printer.CurrentX = 25
            Printer.Print Mid(moParrilla.TextoLinCol(iFila, 4), 1, 15)
            Printer.CurrentY = dY
            Printer.CurrentX = 55
            Printer.Print sAux2
            Printer.CurrentY = dY
            Printer.CurrentX = 135
            Printer.Print sAux
            Printer.CurrentY = dY
            Printer.CurrentX = 150
            Printer.Print moParrilla.TextoLinCol(iFila, 27)
            Printer.CurrentY = dY
            Printer.CurrentX = 170
            Printer.Print moParrilla.TextoLinCol(iFila, 13) & "/" & Trim(moParrilla.TextoLinCol(iFila, 19)) & "/" & Trim(moParrilla.TextoLinCol(iFila, 20))
            Printer.CurrentY = dY
            Printer.CurrentX = 210
            Printer.Print Trim(moParrilla.TextoLinCol(iFila, 18))
            Printer.CurrentY = dY
            Printer.CurrentX = 240
            Printer.Print Trim(moParrilla.TextoLinCol(iFila, 33))
            
            dTotal = dTotal + moParrilla.TextoLinCol(iFila, 27)
        
            dY = dY + 4.5
            If sFilas <> "" And dY > 185 Then
                
                Printer.NewPage
                
                dY = 15
                Printer.FontUnderline = True
                Printer.FontSize = 9
                Printer.CurrentY = dY
                Printer.CurrentX = 5
                Printer.Print "Fecha"
                Printer.CurrentY = dY
                Printer.CurrentX = 25
                Printer.Print "T. Gasto"
                Printer.CurrentY = dY
                Printer.CurrentX = 55
                Printer.Print "Donde / Itinerario / Concepto"
                Printer.CurrentY = dY
                Printer.CurrentX = 135
                Printer.Print "Num."
                Printer.CurrentY = dY
                Printer.CurrentX = 150
                Printer.Print "Importe"
                Printer.CurrentY = dY
                Printer.CurrentX = 170
                Printer.Print "Imp. ML/Div./Cam."
                Printer.CurrentY = dY
                Printer.CurrentX = 210
                Printer.Print "PEP/CECO/IO"
                Printer.CurrentY = dY
                Printer.CurrentX = 240
                Printer.Print "Documento/F.Conta SAP"
                Printer.FontBold = False
                Printer.FontUnderline = False
                
                dY = dY + 5
                
            End If
            
        Wend
        
        If dY > 185 Then
            Printer.NewPage
            dY = 15
        End If
        Printer.FontBold = True
        Printer.FontSize = 10
        Printer.CurrentY = dY
        Printer.CurrentX = 55
        Printer.Print "Total a pagar"
        Printer.CurrentY = dY
        Printer.CurrentX = 145
        Printer.Print dTotal
        
        Printer.EndDoc
        Printer.KillDoc
        
        Me.MousePointer = DEFAULT
        MsgBox "Enviado documento a la impresora.", vbApplicationModal + vbInformation, "Información"
    
    Else
        MsgBox "No hay datos para imprimir.", vbApplicationModal + vbExclamation, "Información"
    End If

Exit_Rutina:
    Exit Sub
Error_Rutina:
    Me.MousePointer = DEFAULT
    If Err.Number = 32755 Then Resume Exit_Rutina
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub ImprimirPagos(Optional bVieneDeBotonPagado As Boolean = False)

    Dim iFila As Integer
    Dim sFilas As String
    Dim sAux As String
    Dim sAux2 As String
    
    Dim dY As Double
    
    Dim dTotal As Double
    Dim dTotalOrdinario As Double
    Dim dTotalFormacion As Double
    Dim dTotalPresentacion As Double
    
    Me.MousePointer = DEFAULT
    
    sFilas = ""
    If bVieneDeBotonPagado Then
    'Sí viene tras haber marcado gastos como pagados, lo que compruebo únicamente es que el gasto esté validado, el pagado se habrá puesto en la tabla pero aún no está reflejado en la parrilla
        For iFila = 1 To moParrilla.Grid.Rows - 1
            'veo que el gasto esté validado
            If moParrilla.TextoLinCol(iFila, 2) = "-->" And moParrilla.TextoLinCol(iFila, 23) = "S" Then sFilas = sFilas & iFila & ","
        Next
    
    Else
        For iFila = 1 To moParrilla.Grid.Rows - 1
            'veo que el gasto esté validado y pagado
            If moParrilla.TextoLinCol(iFila, 2) = "-->" And moParrilla.TextoLinCol(iFila, 23) = "S" And moParrilla.TextoLinCol(iFila, 24) = "S" Then sFilas = sFilas & iFila & ","
        Next
    End If
    
    If sFilas <> "" Then
    
        sFilas = Mid(sFilas, 1, Len(sFilas) - 1)
            
        Printer.ScaleMode = vbMillimeters
        
        Printer.Font = "Courier New"
        Printer.FontSize = 22
        Printer.FontBold = True
        Printer.CurrentX = 45
        Printer.CurrentY = 15
        Printer.Print "Gastos/Anticipos pendientes de pago"
        
        Printer.FontBold = True
        Printer.FontSize = 14
        Printer.CurrentY = 30
        Printer.CurrentX = 20
        Printer.Print "FIRMA DEL PAGADOR"
        Printer.CurrentY = 30
        Printer.CurrentX = 120
        Printer.Print "FIRMA DEL EMPLEADO"
        Printer.FontSize = 14
        Printer.CurrentY = 55
        Printer.CurrentX = 70
        Printer.Print "Fecha " & Format(Now, "dd/MM/yyyy")
        Printer.FontBold = False

        Printer.FontBold = True
        Printer.FontSize = 14
        Printer.CurrentY = 75
        Printer.CurrentX = 20
        Printer.Print "Gastos/Anticipos de " & cmbEmpleados.Text
                
        Printer.FontUnderline = True
        Printer.FontSize = 10
        Printer.CurrentY = 85
        Printer.CurrentX = 10
        Printer.Print "Fecha"
        Printer.CurrentY = 85
        Printer.CurrentX = 30
        Printer.Print "T. Gasto"
        Printer.CurrentY = 85
        Printer.CurrentX = 60
        '14-06-18
        sAux = "Donde / Itinerario / Concepto"
        If CodPaisTrabajo = giPAIS_GES Then sAux = "PEP-CECO / Donde / Itinerario / Concepto"
        Printer.Print sAux
        Printer.CurrentY = 85
        Printer.CurrentX = 150
        Printer.Print "Importe"
        Printer.CurrentY = 85
        Printer.CurrentX = 180
        Printer.Print "Num."
        Printer.FontBold = False
        Printer.FontUnderline = False

        dY = 90
        dTotal = 0
        dTotalOrdinario = 0
        dTotalFormacion = 0
        dTotalPresentacion = 0
        While sFilas <> ""
        
            If InStr(1, sFilas, ",") Then
                iFila = Mid(sFilas, 1, InStr(1, sFilas, ",") - 1)
                sFilas = Mid(sFilas, InStr(1, sFilas, ",") + 1)
            Else
                iFila = sFilas
                sFilas = ""
            End If
            
            sAux = ""
            Select Case moParrilla.TextoLinCol(iFila, 4)
            Case "Kilómetros"
                sAux = moParrilla.TextoLinCol(iFila, 10)
            Case "Comidas-Invitaciones", "Comidas/Invitaciones"
                sAux = moParrilla.TextoLinCol(iFila, 12)
            Case "Hoteles"
                sAux = moParrilla.TextoLinCol(iFila, 9) & "/" & moParrilla.TextoLinCol(iFila, 12)
            Case "Peaje", "Peajes"
                sAux = moParrilla.TextoLinCol(iFila, 11)
            End Select
            
            '14-06-18
            sAux2 = moParrilla.TextoLinCol(iFila, 7)
            If CodPaisTrabajo = giPAIS_GES Then sAux2 = Trim(moParrilla.TextoLinCol(iFila, 18)) & "/" & Trim(moParrilla.TextoLinCol(iFila, 7))
            sAux2 = Mid(sAux2, 1, 50)
        
            Printer.FontSize = 8
            Printer.CurrentY = dY
            Printer.CurrentX = 10
            Printer.Print moParrilla.TextoLinCol(iFila, 3)
            Printer.CurrentY = dY
            Printer.CurrentX = 30
            Printer.Print Mid(moParrilla.TextoLinCol(iFila, 4), 1, 15)
            Printer.CurrentY = dY
            Printer.CurrentX = 60
            Printer.Print sAux2
            Printer.CurrentY = dY
            Printer.CurrentX = 150
            '14-06-18 Pintamos importe con el cambio aplicado
            'Printer.Print moParrilla.TextoLinCol(iFila, 13)
            Printer.Print moParrilla.TextoLinCol(iFila, 27)
            Printer.CurrentY = dY
            Printer.CurrentX = 180
            Printer.Print sAux
            
            '14-06-18 Pintamos importe con el cambio aplicado
            'dTotal = dTotal + moParrilla.TextoLinCol(iFila, 13)
            dTotal = dTotal + moParrilla.TextoLinCol(iFila, 27)
            Select Case moParrilla.TextoLinCol(iFila, 5)
            Case "Ordinario"
                '14-06-18 Pintamos importe con el cambio aplicado
                'dTotalOrdinario = dTotalOrdinario + moParrilla.TextoLinCol(iFila, 13)
                dTotalOrdinario = dTotalOrdinario + moParrilla.TextoLinCol(iFila, 27)
            Case "Formación"
                '14-06-18 Pintamos importe con el cambio aplicado
                'dTotalFormacion = dTotalFormacion + moParrilla.TextoLinCol(iFila, 13)
                dTotalFormacion = dTotalFormacion + moParrilla.TextoLinCol(iFila, 27)
            Case "Presentación"
                '14-06-18 Pintamos importe con el cambio aplicado
                'dTotalPresentacion = dTotalPresentacion + moParrilla.TextoLinCol(iFila, 13)
                dTotalPresentacion = dTotalPresentacion + moParrilla.TextoLinCol(iFila, 27)
            End Select
        
            dY = dY + 4.5
            If sFilas <> "" And dY > 260 Then
                
                Printer.NewPage
                
                dY = 15
                Printer.FontBold = True
                Printer.FontUnderline = True
                Printer.FontSize = 10
                Printer.CurrentY = dY
                Printer.CurrentX = 10
                Printer.Print "Fecha"
                Printer.CurrentY = dY
                Printer.CurrentX = 30
                Printer.Print "T. Gasto"
                Printer.CurrentY = dY
                Printer.CurrentX = 60
                Printer.Print "Donde / Itinerario / Concepto"
                Printer.CurrentY = dY
                Printer.CurrentX = 150
                Printer.Print "Importe"
                Printer.CurrentY = dY
                Printer.CurrentX = 180
                Printer.Print "Num."
                Printer.FontBold = False
                Printer.FontUnderline = False
                
                dY = dY + 5
                
            End If
            
        Wend
        
        If dY > 250 Then
            Printer.NewPage
            dY = 15
        End If
        Printer.FontBold = True
        Printer.FontSize = 12
        Printer.CurrentY = dY
        Printer.CurrentX = 60
        Printer.Print "Total a pagar"
        Printer.CurrentY = dY
        Printer.CurrentX = 150
        Printer.Print dTotal
        If dTotalOrdinario > 0 Then
            dY = dY + 5
            Printer.CurrentY = dY
            Printer.CurrentX = 60
            Printer.Print "Ordinario"
            Printer.CurrentY = dY
            Printer.CurrentX = 150
            Printer.Print dTotalOrdinario
        End If
        If dTotalFormacion > 0 Then
            dY = dY + 5
            Printer.CurrentY = dY
            Printer.CurrentX = 60
            Printer.Print "Formación"
            Printer.CurrentY = dY
            Printer.CurrentX = 150
            Printer.Print dTotalFormacion
        End If
        If dTotalPresentacion > 0 Then
            dY = dY + 5
            Printer.CurrentY = dY
            Printer.CurrentX = 60
            Printer.Print "Presentación"
            Printer.CurrentY = dY
            Printer.CurrentX = 150
            Printer.Print dTotalPresentacion
        End If
        
        Printer.EndDoc
        Printer.KillDoc
    
    Else
        MsgBox "No hay datos para imprimir.", vbApplicationModal + vbExclamation, "Información"
    End If

End Sub

Private Sub imgPagar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgPassword_Click()

    Dim sPass As String

    Me.MousePointer = DEFAULT
    
    '08-11-21 Por indicaciones de Filo, quitamos la opción de que puedan ver la clave de los usuarios
    Exit Sub
    '08-11-21 Fin
    
    If cmbDepartamentos.Text = "" Then
        cmbDepartamentos.SetFocus
        Exit Sub
    End If
    If cmbEmpleados.Text = "" Then
        cmbEmpleados.SetFocus
        Exit Sub
    End If
    If cmbEmpleados.ListIndex = -1 Then
        cmbEmpleados.SetFocus
        Exit Sub
    End If
    
    If bDamePassUsuInterno(CodPaisTrabajo, cmbDepartamentos.ItemData(cmbDepartamentos.ListIndex), cmbEmpleados.ItemData(cmbEmpleados.ListIndex), sPass) Then
        frmPass.Create sPass
        frmPass.Show MODAL
    Else
        MsgBox "No se han encontrado datos.", vbApplicationModal + vbExclamation, "Información"
    End If
    
End Sub

Private Sub imgPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgPreValidar_Click()

    Dim iFila As Integer
    Dim sContadores As String
    Dim sAux As String
    Dim sUsuarioPreValidacion As String
    
    Me.MousePointer = DEFAULT
    
    If mid_Ofi_PreValidacion = 0 Then Exit Sub
    miPosicionarEnFila = -1
    
    sUsuarioPreValidacion = Trim(goUsuario.IdUsuario)
    
    sContadores = ""
    For iFila = 1 To moParrilla.Grid.Rows - 1
        If moParrilla.TextoLinCol(iFila, 2) = "-->" Then sContadores = sContadores & moParrilla.TextoLinCol(iFila, 0) & ","
    Next
    If sContadores <> "" Then
    
        sContadores = Mid(sContadores, 1, Len(sContadores) - 1)
        gBdGALILEO.BeginTrans True
        While sContadores <> ""
        
            If InStr(1, sContadores, ",") Then
                sAux = Mid(sContadores, 1, InStr(1, sContadores, ",") - 1)
                sContadores = Mid(sContadores, InStr(1, sContadores, ",") + 1)
            Else
                sAux = sContadores
                sContadores = ""
            End If
            If Not bMarcarGastosInternosPreValidados(sAux, sUsuarioPreValidacion) Then
                gBdGALILEO.RollbackTrans
                MsgBox "Se ha producido un error al intentar marcar los gastos seleccionados como prevalidados.", vbApplicationModal + vbExclamation, "E R R O R"
                Exit Sub
            End If
            
        Wend
        gBdGALILEO.CommitTrans
        MsgBox "Cambios realizados correctamente.", vbApplicationModal + vbInformation, "Información"
        RellenarDatosGastos
        
    Else
        MsgBox "No hay seleccionado ningún gasto pendiente de ser prevalidados.", vbApplicationModal + vbExclamation, "Información"
    End If

End Sub

Private Sub imgPreValidar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgSalir_Click()
    
    Me.MousePointer = DEFAULT
    Set mafoRegistro = Nothing
    Set mafoRegEmpleados = Nothing
    Set mafoRegCuentasContables = Nothing
    If mbConexionSAP Then DesLogon
    If miCodEmpleado <> -1 Then EmpleadoDesBloquear CodPaisTrabajo, miCodDepartamento, miCodEmpleado, msJerEmpleado
    EmpleadoDesBloquear_UsuGasto
    
    Dim Formulario As Form
    For Each Formulario In Forms
        If UCase(Formulario.Name) = "FRMEXPLORERV2" Then Unload Formulario
    Next
    
    Unload Me
    
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgSimuladorCambios_Click()

    Me.MousePointer = DEFAULT
    
    If Trim(txtFecha) = "" Then Exit Sub
    If Trim(txtImporte) = "" Then Exit Sub
    If Trim(cmbDivisa.Text) = "" Then Exit Sub
    If Trim(txtCambio) = "" Then Exit Sub
    If Not IsNumeric(Trim(txtCambio)) Then Exit Sub
    
    frmSimuladorCambio.Create txtFecha, txtImporte, cmbDivisa.Text, CDbl(Trim(txtCambio)), msMonedaDepartamento
    frmSimuladorCambio.Show MODAL
    If frmSimuladorCambio.CambioSeleccionado <> -999999.9999 Then
        txtCambio = frmSimuladorCambio.CambioSeleccionado
        ReCalcularImporteCambio
    End If

End Sub

Private Sub imgSimuladorCambios_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgValidar_Click()
    
    Dim iFila As Integer
    Dim sContadores As String
    Dim sContadoresVISA As String
    Dim sAux As String
    
    Me.MousePointer = DEFAULT
    
    miPosicionarEnFila = -1
    If CodPaisTrabajo = giPAIS_GES Then
        If goUsuario.IdUsuario <> "lanchi" Then
            If Not goUsuario.Permiso(giOP_GVB_VALIDARGASTOINTERNO) Then
                MsgBox "Opción no disponible para el País/Empresa actual.", vbApplicationModal + vbInformation, "Información"
                Exit Sub
            End If
        End If
    End If
    
    sContadores = ""
    sContadoresVISA = ""
    For iFila = 1 To moParrilla.Grid.Rows - 1
        If moParrilla.TextoLinCol(iFila, 2) = "-->" And moParrilla.TextoLinCol(iFila, 23) = "N" Then sContadores = sContadores & moParrilla.TextoLinCol(iFila, 0) & ","
        If moParrilla.TextoLinCol(iFila, 2) = "-->" And moParrilla.TextoLinCol(iFila, 23) = "S" And moParrilla.TextoLinCol(iFila, 6) = "VISA" Then sContadoresVISA = sContadoresVISA & moParrilla.TextoLinCol(iFila, 0) & ","
    Next
    If sContadores = "" And sContadoresVISA = "" Then
        MsgBox "No hay seleccionado ningún gasto pendiente de ser validado.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    
    gBdGALILEO.BeginTrans True
    
    If sContadores <> "" Then
    
        sContadores = Mid(sContadores, 1, Len(sContadores) - 1)
        While sContadores <> ""
        
            If InStr(1, sContadores, ",") Then
                sAux = Mid(sContadores, 1, InStr(1, sContadores, ",") - 1)
                sContadores = Mid(sContadores, InStr(1, sContadores, ",") + 1)
            Else
                sAux = sContadores
                sContadores = ""
            End If
            If Not bMarcarGastosInternosValidados(sAux, "Aplic GastosInt") Then
                gBdGALILEO.RollbackTrans
                MsgBox "Se ha producido un error al intentar marcar los gastos seleccionados como validados.", vbApplicationModal + vbExclamation, "E R R O R"
                Exit Sub
            End If
            
        Wend
    
    End If
    
    If sContadoresVISA <> "" Then
    
        sContadoresVISA = Mid(sContadoresVISA, 1, Len(sContadoresVISA) - 1)
        While sContadoresVISA <> ""
        
            If InStr(1, sContadoresVISA, ",") Then
                sAux = Mid(sContadoresVISA, 1, InStr(1, sContadoresVISA, ",") - 1)
                sContadoresVISA = Mid(sContadoresVISA, InStr(1, sContadoresVISA, ",") + 1)
            Else
                sAux = sContadoresVISA
                sContadoresVISA = ""
            End If
            If Not bMarcarGastosInternosValidados(sAux, "VISA VALIDADA") Then
                gBdGALILEO.RollbackTrans
                MsgBox "Se ha producido un error al intentar marcar los gastos seleccionados como validados.", vbApplicationModal + vbExclamation, "E R R O R"
                Exit Sub
            End If
            
        Wend
    
    End If
    
    gBdGALILEO.CommitTrans
    MsgBox "Cambios realizados correctamente.", vbApplicationModal + vbInformation, "Información"
    RellenarDatosGastos
    
End Sub

Private Sub imgValidar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub lblTexto_Click(Index As Integer)

    Select Case Index
    Case 4
        shFiltro_PendienteValidar.Visible = Not shFiltro_PendienteValidar.Visible
        shFiltro_Rechazado.Visible = False
        shFiltro_ValidadoNoPagado.Visible = False
        shFiltro_ValidadoPagado.Visible = False
        shFiltro_ContabilizadoSAP.Visible = False
        shFiltro_SoloValidado.Visible = False
        shFiltro_PendientePreValidar.Visible = False
    Case 5
        shFiltro_PendienteValidar.Visible = False
        shFiltro_Rechazado.Visible = False
        shFiltro_ValidadoNoPagado.Visible = False
        shFiltro_ValidadoPagado.Visible = Not shFiltro_ValidadoPagado.Visible
        shFiltro_ContabilizadoSAP.Visible = False
        shFiltro_SoloValidado.Visible = False
        shFiltro_PendientePreValidar.Visible = False
    Case 6
        shFiltro_PendienteValidar.Visible = False
        shFiltro_Rechazado.Visible = False
        shFiltro_ValidadoNoPagado.Visible = Not shFiltro_ValidadoNoPagado.Visible
        shFiltro_ValidadoPagado.Visible = False
        shFiltro_ContabilizadoSAP.Visible = False
        shFiltro_SoloValidado.Visible = False
        shFiltro_PendientePreValidar.Visible = False
    Case 26
        shFiltro_PendienteValidar.Visible = False
        shFiltro_Rechazado.Visible = False
        shFiltro_ValidadoNoPagado.Visible = False
        shFiltro_ValidadoPagado.Visible = False
        shFiltro_ContabilizadoSAP.Visible = Not shFiltro_ContabilizadoSAP.Visible
        shFiltro_SoloValidado.Visible = False
        shFiltro_PendientePreValidar.Visible = False
    Case 32
        shFiltro_PendienteValidar.Visible = False
        shFiltro_Rechazado.Visible = Not shFiltro_Rechazado.Visible
        shFiltro_ValidadoNoPagado.Visible = False
        shFiltro_ValidadoPagado.Visible = False
        shFiltro_ContabilizadoSAP.Visible = False
        shFiltro_SoloValidado.Visible = False
    Case 34
        shFiltro_PendienteValidar.Visible = False
        shFiltro_Rechazado.Visible = False
        shFiltro_ValidadoNoPagado.Visible = False
        shFiltro_ValidadoPagado.Visible = False
        shFiltro_ContabilizadoSAP.Visible = False
        shFiltro_SoloValidado.Visible = Not shFiltro_SoloValidado.Visible
        shFiltro_PendientePreValidar.Visible = False
    Case 42
        shFiltro_PendienteValidar.Visible = False
        shFiltro_Rechazado.Visible = False
        shFiltro_ValidadoNoPagado.Visible = False
        shFiltro_ValidadoPagado.Visible = False
        shFiltro_ContabilizadoSAP.Visible = False
        shFiltro_PendientePreValidar.Visible = Not shFiltro_PendientePreValidar.Visible
    Case Else
        Exit Sub
    End Select
    miPosicionarEnFila = -1
    RellenarDatosGastos

End Sub

Private Sub lblTexto_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Index = 4 Or Index = 5 Or Index = 6 Or Index = 26 Or Index = 32 Or Index = 34 Then
        Set Me.MouseIcon = imgCursor.Picture
        Me.MousePointer = 99
    Else
        Me.MousePointer = DEFAULT
    End If

End Sub

Private Sub mfgHistoricoModificaciones_Click()
    If moParrillaHistorico.bLineaActualVacia Then
        moParrillaHistorico.SituarEn 0, 1
        Exit Sub
    End If
    moParrillaHistorico.ResaltarLinea
End Sub

Private Sub mfgParrilla_Click()
    
    Dim iFila As Integer
    Dim iCont As Integer
    Dim sAux As String
    
    LimpiarCampos False
    
    If moParrilla.Grid.Rows > 2 Then
    
        If moParrilla.Grid.MouseRow = 0 Then
            miColumnaOrdenacion = moParrilla.Grid.MouseCol
            moParrilla.OrdenarSiCorresponde
            'moParrilla.SituarEn 0, 1
            Exit Sub
        End If
        
    End If
    
    If moParrilla.bLineaActualVacia Then
        moParrilla.SituarEn 0, 1
        Exit Sub
    End If
    
    moParrilla.ResaltarLinea
    
    iFila = moParrilla.Grid.Row
    
    '06-08-21 Sí es un gasto asociado a otro, no lo llevamos abajo
    If moParrilla.TextoLinCol(iFila, 49) <> "0" Then Exit Sub
    
    lblContador = moParrilla.TextoLinCol(iFila, 0)
    lblRuta = moParrilla.TextoLinCol(iFila, 1)
    txtFecha = moParrilla.TextoLinCol(iFila, 3)
    txtConceptoGasto = moParrilla.TextoLinCol(iFila, 4)
    txtFamiliaGasto = moParrilla.TextoLinCol(iFila, 5)
    txtTipoPago = moParrilla.TextoLinCol(iFila, 6)
    txtDonde = moParrilla.TextoLinCol(iFila, 7)
    txtTipoComida = moParrilla.TextoLinCol(iFila, 8)
    txtNumNoches = moParrilla.TextoLinCol(iFila, 9)
    txtKms = moParrilla.TextoLinCol(iFila, 10)
    txtNumTickets = moParrilla.TextoLinCol(iFila, 11)
    txtNumPersonas = moParrilla.TextoLinCol(iFila, 12)
    txtImporte = moParrilla.TextoLinCol(iFila, 13)
    txtProveedor = moParrilla.TextoLinCol(iFila, 14)
    txtFactura = moParrilla.TextoLinCol(iFila, 15)
    txtPersonas = moParrilla.TextoLinCol(iFila, 16)
    If cmbEmpresas.Enabled Then
        If moParrilla.TextoLinCol(iFila, 17) <> "" Then
            cmbEmpresas.Text = moParrilla.TextoLinCol(iFila, 17)
        Else
        
            For iCont = 0 To cmbEmpresas.ListCount - 1
                If CInt(cmbEmpresas.ItemData(iCont)) = Empresa_CodPaisTrabajo Then
                    cmbEmpresas.ListIndex = iCont
                End If
            Next
            'cmbEmpresas.ListIndex = -1
            
        End If
    End If
    If txtPEPCECO.Enabled Then txtPEPCECO.Text = moParrilla.TextoLinCol(iFila, 18)
    If cmbDivisa.Enabled Then cmbDivisa.Text = moParrilla.TextoLinCol(iFila, 19)
    If txtCambio.Enabled Then
        txtCambio.Text = moParrilla.TextoLinCol(iFila, 20)
        txtCambio.Tag = txtCambio.Text
        If moParrilla.TextoLinCol(iFila, 36) = "S" Then shCambioUsuario.Visible = True
    End If
    If txtCambio.Enabled Then lblImporteEUR = moParrilla.TextoLinCol(iFila, 27)
    lblPagado = moParrilla.TextoLinCol(iFila, 24)
    lblContabilizadoSAP = moParrilla.TextoLinCol(iFila, 26)
    lblValidado = moParrilla.TextoLinCol(iFila, 23)
    imgModificado.Visible = False
    If moParrilla.TextoLinCol(iFila, 28) = "S" Then imgModificado.Visible = True
    txtCIF_Proveedor.Text = moParrilla.TextoLinCol(iFila, 22)
    If LCase(txtCIF_Proveedor) = "cpd" Then
        imgDatos_ProveedorCPD.Visible = True
        lblCPD_Nombre = moParrilla.TextoLinCol(iFila, 37)
        lblCPD_CP = moParrilla.TextoLinCol(iFila, 38)
        lblCPD_Ciudad = moParrilla.TextoLinCol(iFila, 39)
        lblCPD_Pais = moParrilla.TextoLinCol(iFila, 40)
        lblCPD_CIF = moParrilla.TextoLinCol(iFila, 41)
    End If
    sAux = moParrilla.TextoLinCol(iFila, 35)
    If sAux <> "" Then
        For iCont = 0 To cmbCuentasContables.ListCount - 1
            If Trim(Mid(cmbCuentasContables.List(iCont), 1, InStr(1, cmbCuentasContables.List(iCont), "-") - 1)) = sAux Then
                cmbCuentasContables.ListIndex = iCont
                Exit For
            End If
        Next
    End If
    
    If txtConceptoGasto = "Kilómetros" Then
        txtImporte = moParrilla.TextoLinCol(iFila, 25)
    End If

    If moParrilla.TextoLinCol(iFila, 42) = "S" Then
        chkDesgloseIVA.Tag = "C"
        chkDesgloseIVA.value = vbChecked
        chkDesgloseIVA.Tag = ""
        txtImporteIVA.Enabled = True
        txtImporteIVA.Text = moParrilla.TextoLinCol(iFila, 43)
        txtBaseIVA.Enabled = True
        txtBaseIVA.Text = moParrilla.TextoLinCol(iFila, 44)
        txtTipoIVA.Enabled = True
        txtTipoIVA.Text = moParrilla.TextoLinCol(iFila, 45)
'        txtImporteIVA2.Enabled = True
'        If CDbl(moParrilla.TextoLinCol(iFila, 46)) > 0 Then txtImporteIVA2.Text = moParrilla.TextoLinCol(iFila, 46)
'        txtBaseIVA2.Enabled = True
'        If CDbl(moParrilla.TextoLinCol(iFila, 47)) > 0 Then txtBaseIVA2.Text = moParrilla.TextoLinCol(iFila, 47)
'        txtTipoIVA2.Enabled = True
'        txtTipoIVA2.Text = moParrilla.TextoLinCol(iFila, 48)
    End If

    If chkVerTicketAutomaticamente.value = vbChecked Then imgConsultar_Click

    CamposPorConceptoGasto
    'txtConceptoGasto.SetFocus
        
End Sub

Private Sub mfgParrilla_DblClick()

    Dim iFila As Integer
   
    If moParrilla.bLineaActualVacia Then
        moParrilla.SituarEn 0, 1
        Exit Sub
    End If
    
    iFila = moParrilla.Grid.Row
    If moParrilla.TextoLinCol(iFila, 2) = "" Then
        moParrilla.TextoLinCol(iFila, 2) = "-->"
        moParrilla.TextoLinCol(iFila, 50) = "<--"
    Else
        moParrilla.TextoLinCol(iFila, 2) = ""
        moParrilla.TextoLinCol(iFila, 50) = ""
    End If
    moParrilla.ResaltarLinea

End Sub

Private Sub mfgParrilla_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Function sFamiliaGasto(sFamilia As String) As String

    Select Case sFamilia
    Case "O", "1"
        sFamiliaGasto = "Ordinario"
    Case "F", "2"
        sFamiliaGasto = "Formación"
    Case "P", "3"
        sFamiliaGasto = "Presentación"
    Case "Ordinario"
        sFamiliaGasto = "O"
    Case "Formación"
        sFamiliaGasto = "F"
    Case "Presentación"
        sFamiliaGasto = "P"
    Case Else
        sFamiliaGasto = ""
    End Select
    
End Function

Private Function sTipoPagoGasto(sTipo As String) As String

    Select Case sTipo
    Case "N", "1"
        sTipoPagoGasto = "Metálico"
    Case "S", "2"
        sTipoPagoGasto = "VISA"
    Case "F", "3"
        sTipoPagoGasto = "Facturado"
    Case "R", "4"
        sTipoPagoGasto = "SolRed"
    Case "Metálico"
        sTipoPagoGasto = "N"
    Case "VISA"
        sTipoPagoGasto = "S"
    Case "Facturado"
        sTipoPagoGasto = "F"
    Case "SolRed"
        sTipoPagoGasto = "R"
    Case Else
        sTipoPagoGasto = ""
    End Select
    
End Function

Private Function sTipoComidasInvitaciones(sTipo As String) As String

    Select Case sTipo
    Case "N", "1"
        sTipoComidasInvitaciones = "Negocios"
    Case "I", "2"
        sTipoComidasInvitaciones = "Interno"
    Case "Negocios"
        sTipoComidasInvitaciones = "N"
    Case "Interno"
        sTipoComidasInvitaciones = "I"
    Case Else
        sTipoComidasInvitaciones = ""
    End Select
    
End Function

Private Sub RellenarDatosGastos()

    Dim lContador As Long
    Dim lContadorAsociado As Long
    Dim sRutaTicket As String
    Dim sFecha As String
    Dim sConceptoGasto As String
    Dim sTipoGasto As String
    Dim sNomGasto As String
    Dim sEntrada As String
    Dim sDonde As String
    Dim sTipoComida As String
    Dim sNumNoches As String
    Dim sKms As String
    Dim sNumTickets As String
    Dim sNumPersonas As String
    Dim sImporte As String
    Dim sImporte2 As String
    Dim sProveedor As String
    Dim sFactura As String
    Dim sValidado As String
    Dim sPagado As String
    Dim sPndtePreValidar As String
    Dim sPersonas As String
    Dim sEmpresa As String
    Dim sImporte_EUR As String
    Dim sPEPCECO As String
    Dim sDivisa As String
    Dim sCambio As String
    Dim sArchivadoEn As String
    Dim sContaSAP As String
    Dim sModificadoAdm As String
    Dim sCIF_Proveedor As String
    Dim sFecha_Val_Rec As String
    Dim sMotivo_Rechazo As String
    Dim sCuenta_Contable As String
    Dim sCodCuenta_Contable As String
    Dim sFecha_Insert As String
    Dim sEjercicioDocumentoSAP As String
    Dim sCambioGuardadoPorUsuario As String
    Dim sCPD_Nombre As String
    Dim sCPD_CP As String
    Dim sCPD_Ciudad As String
    Dim sCPD_Pais As String
    Dim sCPD_CIF As String
    Dim sValidador As String
    Dim sDesglose_IVA As String
    Dim sImporte_IVA As String
    Dim sBase_IVA As String
    Dim sTipo_IVA As String
    Dim sImporte_IVA2 As String
    Dim sBase_IVA2 As String
    Dim sTipo_IVA2 As String
        
    Dim iFila As Integer
    Dim lColor As Long
    
    Dim bRellenar As Boolean
    
    Dim sAux As String

    Me.MousePointer = HOURGLASS
    DoEvents
    
    moParrilla.Vaciar
    moParrilla.SituarEn 0, 1
    LimpiarCampos
    
    sValidado = "X"
    sPagado = "X"
    sContaSAP = "X"
    sPndtePreValidar = "X"
    If shFiltro_PendienteValidar.Visible = True Then sValidado = "N"
    If shFiltro_Rechazado.Visible = True Then sValidado = "R"
    If shFiltro_ValidadoNoPagado.Visible = True Then
        sValidado = "S"
        '31-10-18 Buscamos solo gastos validados, da igual sí pagados o no
        'sPagado = "N"
    End If
    If shFiltro_SoloValidado.Visible = True Then
        sValidado = "S"
        sPagado = "N"
        sContaSAP = "N"
    End If
    If shFiltro_ValidadoPagado.Visible = True Then
        '31-10-18 Buscamos solo gastos pagados
        'sValidado = "S"
        sPagado = "S"
    End If
    If shFiltro_ContabilizadoSAP.Visible = True Then sContaSAP = "S"
    If shFiltro_PendientePreValidar.Visible = True Then sPndtePreValidar = "S"

    If bDameGastosPersonalInterno_v3(CodPaisTrabajo, miCodDepartamento, miCodEmpleado, msJerEmpleado, Format(dtpFechaDesde.value, "ddMMyyyy"), Format(dtpFechaHasta.value, "ddMMyyyy"), sValidado, sPagado, sContaSAP, mid_Ofi_PreValidacion, mafoRegistro) Then
                
        While Not mafoRegistro.EOF
        
            bRellenar = True
        
            lContador = NoNull(mafoRegistro("Contador"))
            lContadorAsociado = NoNull(mafoRegistro("Contador_Asociado"))
            sRutaTicket = Trim(NoNull(mafoRegistro("Ruta_Ticket"), "S"))
            sFecha = Format(mafoRegistro("Fecha"), "dd/MM/yyyy")
            sConceptoGasto = Trim(NoNull(mafoRegistro("TipoGasto"), "T"))
            sNomGasto = sNombreGasto(msTipo, sConceptoGasto)
            sTipoGasto = Trim(NoNull(mafoRegistro("FamiliaGasto"), "T"))
            If sTipoGasto = "" Then sTipoGasto = "O"
            sEntrada = Trim(NoNull(mafoRegistro("Visa"), "T"))
            If sEntrada = "" Then sEntrada = "N"
            sDonde = Trim(NoNull(mafoRegistro("Donde"), "T"))
            sTipoComida = Trim(NoNull(mafoRegistro("Tipo_Com_Inv"), "T"))
            sNumNoches = Trim(NoNull(mafoRegistro("Numero"), "T"))
            sKms = Trim(NoNull(mafoRegistro("Numero"), "T"))
            sNumTickets = Trim(NoNull(mafoRegistro("Numero"), "T"))
            sNumPersonas = Trim(NoNull(mafoRegistro("Numero"), "T"))
            'Si son Hoteles, el número de personas viene numero2
            If sNomGasto = "Hoteles" Then sNumPersonas = Trim(NoNull(mafoRegistro("Numero2"), "T"))
            '14-06-18
            'sImporte = Trim(NoNull(mafoRegistro("Importe"), "N"))
            sImporte = Trim(NoNull(mafoRegistro("Importe_Ori"), "N"))
            sImporte2 = sImporte
            sProveedor = Trim(NoNull(mafoRegistro("Proveedor"), "T"))
            sFactura = Trim(NoNull(mafoRegistro("Factura"), "T"))
            sValidado = Trim(NoNull(mafoRegistro("Validado"), "T"))
            sPagado = Trim(NoNull(mafoRegistro("Pagado"), "T"))
            sPersonas = Trim(NoNull(mafoRegistro("Personas"), "T"))
            sEmpresa = Trim(NoNull(mafoRegistro("Empresa"), "T"))
            sImporte_EUR = Trim(NoNull(mafoRegistro("Importe"), "N"))
            sPEPCECO = Trim(NoNull(mafoRegistro("PEP_CECO"), "T"))
            sDivisa = Trim(NoNull(mafoRegistro("Moneda"), "T"))
            sCambio = Trim(NoNull(mafoRegistro("Cambio"), "N"))
            sArchivadoEn = Trim(NoNull(mafoRegistro("Archivado_En"), "T"))
            sContaSAP = Trim(NoNull(mafoRegistro("Conta_SAP"), "T"))
            sModificadoAdm = Trim(NoNull(mafoRegistro("ModificadoAdm"), "T"))
            sCIF_Proveedor = Trim(NoNull(mafoRegistro("CIF_Proveedor"), "T"))
            sFecha_Val_Rec = ""
            If Not IsNull(mafoRegistro("Fecha_Val_Rec")) Then sFecha_Val_Rec = Format(mafoRegistro("Fecha_Val_Rec"), "dd/MM/yyyy")
            sMotivo_Rechazo = Trim(NoNull(mafoRegistro("Motivo_Rechazo"), "T"))
            sCuenta_Contable = Trim(NoNull(mafoRegistro("Desc_CuentaContable"), "T"))
            sCodCuenta_Contable = Trim(NoNull(mafoRegistro("Cuenta_Contable"), "T"))
            sFecha_Insert = Trim(NoNull(mafoRegistro("Fecha_Insert"), "T"))
            sEjercicioDocumentoSAP = Trim(NoNull(mafoRegistro("Documento_SAP"), "T"))
            If Not IsNull(mafoRegistro("Fecha_Conta_SAP")) Then
                sAux = Trim(mafoRegistro("Fecha_Conta_SAP"))
                sAux = Mid(sAux, 1, 2) & "-" & Mid(sAux, 3, 2) & "-" & Mid(sAux, 7, 2)
                sEjercicioDocumentoSAP = sEjercicioDocumentoSAP & " / " & sAux
            End If
            sCambioGuardadoPorUsuario = Trim(NoNull(mafoRegistro("Cambio_Guardado_Por_Usuario"), "T"))
            sCPD_Nombre = Trim(NoNull(mafoRegistro("CPD_Nombre"), "T"))
            sCPD_CP = Trim(NoNull(mafoRegistro("CPD_CP"), "T"))
            sCPD_Ciudad = Trim(NoNull(mafoRegistro("CPD_Ciudad"), "T"))
            sCPD_Pais = Trim(NoNull(mafoRegistro("CPD_Pais"), "T"))
            sCPD_CIF = Trim(NoNull(mafoRegistro("CPD_CIF"), "T"))
            sValidador = Trim(NoNull(mafoRegistro("Validador"), "T"))
            If sValidador = "" Then
                sValidador = sDameValidador(lContador, CodPaisTrabajo, miCodDepartamento, miCodEmpleado, sPEPCECO, sValidado)
            Else
                sValidador = sDameNombreUsuario(CodPaisTrabajo, sValidador)
            End If
            sDesglose_IVA = Trim(NoNull(mafoRegistro("Desglose_IVA"), "T"))
            sImporte_IVA = Trim(NoNull(mafoRegistro("Importe_IVA"), "N"))
            sBase_IVA = Trim(NoNull(mafoRegistro("Base_IVA"), "N"))
            sTipo_IVA = Trim(NoNull(mafoRegistro("Tipo_IVA"), "T"))
            sImporte_IVA2 = Trim(NoNull(mafoRegistro("Importe_IVA2"), "N"))
            sBase_IVA2 = Trim(NoNull(mafoRegistro("Base_IVA2"), "N"))
            sTipo_IVA2 = Trim(NoNull(mafoRegistro("Tipo_IVA2"), "T"))
                        
            Select Case sNomGasto
            Case "Kilómetros"
                sNumNoches = "0"
                sNumTickets = "0"
                sNumPersonas = "0"
                sImporte = sImporte * sKms
                sImporte_EUR = sImporte_EUR * sKms
            Case "Comidas-Invitaciones", "Comidas/Invitaciones"
                sNumNoches = "0"
                sKms = "0"
                sNumTickets = "0"
            Case "Hoteles"
                sKms = "0"
                sNumTickets = "0"
            Case "Peaje", "Peajes"
                sNumNoches = "0"
                sKms = "0"
                sNumPersonas = "0"
            End Select
                        
            iFila = moParrilla.LineaLibre
            lColor = vbWhite
            If sValidado = "N" Then lColor = &HC0C0FF
            '31-10-18 Vamos a buscar solo por validado
            'If sValidado = "S" And sPagado = "N" Then lColor = &HC0FFC0
            If sValidado = "S" Then lColor = &HC0FFC0
            If sValidado = "R" Then lColor = &HFFC0FF
            '31-10-18 El contabilizado SAP es anterior al pagado, por lo que lo consultamos antes
            If sContaSAP = "S" Then lColor = &HFFFFC0
            '31-10-18 Vamos a buscar solo por validado
            'If sValidado = "S" And sPagado = "S" Then lColor = vbWhite
            If sPagado = "S" Then lColor = vbWhite
            If Trim(NoNull(mafoRegistro("PreValidado"), "T")) = "N" And sValidado = "N" Then lColor = &HC0FFFF
            
            'Sí el gasto está prevalidado y tenemos marcado el filtro de mostrar solo los pendientes de prevalidar, no pintamos esta línea de gasto
            If sPndtePreValidar = "S" Then
                bRellenar = False
                If Trim(NoNull(mafoRegistro("PreValidado"), "T")) = "N" And sValidado = "N" Then bRellenar = True
            End If
            
            If bRellenar Then
            
                moParrilla.TextoFila CLng(iFila), lContador, sRutaTicket, "", sFecha, sNomGasto, sFamiliaGasto(sTipoGasto), sTipoPagoGasto(sEntrada), sDonde, _
                    sTipoComidasInvitaciones(sTipoComida), sNumNoches, sKms, sNumTickets, sNumPersonas, sImporte, sProveedor, sFactura, sPersonas, sEmpresa, _
                    sPEPCECO, sDivisa, sCambio, sArchivadoEn, sCIF_Proveedor, sValidado, sPagado, sImporte2, sContaSAP, sImporte_EUR, sModificadoAdm, _
                    sFecha_Val_Rec, sMotivo_Rechazo, sCuenta_Contable, sFecha_Insert, sEjercicioDocumentoSAP, sValidador, sCodCuenta_Contable, sCambioGuardadoPorUsuario, _
                    sCPD_Nombre, sCPD_CP, sCPD_Ciudad, sCPD_Pais, sCPD_CIF, sDesglose_IVA, sImporte_IVA, sBase_IVA, sTipo_IVA, sImporte_IVA2, sBase_IVA2, sTipo_IVA2, _
                    lContadorAsociado, ""
                moParrilla.ColorFila CLng(iFila), lColor
                If sModificadoAdm = "S" Then moParrilla.ColorCelda CLng(iFila), 2, vbRed
                If sCambioGuardadoPorUsuario = "S" Then moParrilla.ColorCelda CLng(iFila), 20, &HFFFF&
                
            End If
            
            mafoRegistro.MoveNext
            
        Wend
        moParrilla.SituarEn 0, 1
        
    End If
    Set mafoRegistro = Nothing
    
    If miColumnaOrdenacion <> -1 Then moParrilla.OrdenarPorUnaColumnaDada miColumnaOrdenacion
    If miPosicionarEnFila <> -1 Then
        If moParrilla.Grid.Rows - 1 >= miPosicionarEnFila Then moParrilla.Grid.TopRow = miPosicionarEnFila
        miPosicionarEnFila = -1
    End If
    
    Me.MousePointer = DEFAULT
    txtConceptoGasto.SetFocus
    DoEvents

End Sub

Private Sub LimpiarCampos(Optional bDesmarcarLinea As Boolean = True)
    
    imgModificado.Visible = False
    If bDesmarcarLinea Then moParrilla.SituarEn 0, 1
    lblContador = ""
    lblRuta = ""
    lblFichero = ""
    lblFichero.Tag = ""
    lblPagado = ""
    lblContabilizadoSAP = ""
    lblValidado = ""
    lblCPD_Nombre = ""
    lblCPD_CP = ""
    lblCPD_Ciudad = ""
    lblCPD_Pais = ""
    lblCPD_CIF = ""
    txtConceptoGasto = ""
    txtFecha = ""
    If txtFamiliaGasto.Enabled Then txtFamiliaGasto = ""
    If txtTipoPago.Enabled Then txtTipoPago = ""
    txtDonde = ""
    txtTipoComida = ""
    txtNumNoches = ""
    txtKms = ""
    txtNumTickets = ""
    txtNumPersonas = ""
    txtImporte = ""
    txtProveedor = ""
    txtFactura = ""
    txtPersonas = ""
    txtPEPCECO = ""
    txtCIF_Proveedor = ""
    imgDatos_ProveedorCPD.Visible = False
    cmbCuentasContables.ListIndex = -1
    cmbCuentasContables = ""
    'cmbCuentasContables_Change
    cmbCuentasContables.Enabled = False
    cmbDivisa.ListIndex = -1
    txtCambio = ""
    txtCambio.Tag = ""
    shCambioUsuario.Visible = False
    lblImporteEUR.Caption = ""
    chkDesgloseIVA.Tag = "C"
    chkDesgloseIVA.value = vbUnchecked
    chkDesgloseIVA.Tag = ""
    txtImporteIVA.Text = ""
    txtImporteIVA.Enabled = False
    txtBaseIVA.Text = ""
    txtBaseIVA.Enabled = False
    txtTipoIVA.Text = ""
    txtTipoIVA.Enabled = False
'    txtImporteIVA2.Text = ""
'    txtImporteIVA2.Enabled = False
'    txtBaseIVA2.Text = ""
'    txtBaseIVA2.Enabled = False
'    txtTipoIVA2.Text = ""
'    txtTipoIVA2.Enabled = False
    
    etAyudaConceptoGasto.Visible = False
    mfgTiposGasto.Visible = False
    etFechaAyuda.Visible = False
    etAyudaFamiliaGasto.Visible = False
    etAyudaTipoPago.Visible = False
    etAyudaTipoComida.Visible = False
    On Error Resume Next
    'txtConceptoGasto.SetFocus
    Err.Clear
    
End Sub

Private Sub CamposPorConceptoGasto()
    
    txtTipoComida.Enabled = False
    txtNumNoches.Enabled = False
    txtKms.Enabled = False
    txtNumTickets.Enabled = False
    txtNumPersonas.Enabled = False
    txtPersonas.Enabled = False
    txtImporte.Enabled = True
    txtCIF_Proveedor.Enabled = False
    cmbCuentasContables.Enabled = False
    
    chkDesgloseIVA.Enabled = False
    If CodPaisTrabajo = giPAIS_GES Then chkDesgloseIVA.Enabled = True
    
    Select Case txtConceptoGasto
    Case "Kilómetros"
        txtKms.Enabled = True
        txtImporte.Enabled = False
        txtImporte = msImporteKms
    Case "Comidas-Invitaciones", "Comidas/Invitaciones"
        txtTipoComida.Enabled = True
        txtNumPersonas.Enabled = True
        txtPersonas.Enabled = True
    Case "Hoteles"
        txtNumNoches.Enabled = True
        txtNumPersonas.Enabled = True
    Case "Peaje", "Peajes"
        txtNumTickets.Enabled = True
    Case "Vendor/Proveedor"
        txtCIF_Proveedor.Enabled = True
    Case "Otros Gastos"
        If CodPaisTrabajo = giPAIS_GES Then cmbCuentasContables.Enabled = True
    Case "Anticipo"
        If CodPaisTrabajo = giPAIS_GES Then chkDesgloseIVA.Enabled = False
    Case ""
        
    End Select

End Sub

Private Sub mfgParrilla_SelChange()
    mfgParrilla_Click
End Sub

Private Sub txtBaseIVA_GotFocus()
    TextGotFocus txtBaseIVA
End Sub

Private Sub txtBaseIVA_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    If KeyAscii = 27 Then LimpiarCampos
    EntraNumero KeyAscii, txtBaseIVA, 7, 2, "###0.00"
End Sub

Private Sub txtBaseIVA_LostFocus()
    TextLostFocus txtBaseIVA
End Sub

Private Sub txtBaseIVA2_GotFocus()
    TextGotFocus txtBaseIVA2
End Sub

Private Sub txtBaseIVA2_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    If KeyAscii = 27 Then LimpiarCampos
    EntraNumero KeyAscii, txtBaseIVA2, 7, 2, "###0.00"
End Sub

Private Sub txtBaseIVA2_LostFocus()
    TextLostFocus txtBaseIVA2
End Sub

Private Sub txtCambio_GotFocus()
    TextGotFocus txtCambio
End Sub

Private Sub txtCambio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then LimpiarCampos
    EntraNumero KeyAscii, txtCambio, 7, 4, "###0.000", "S"
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtCambio_LostFocus()
    ReCalcularImporteCambio
    TextLostFocus txtCambio
End Sub

Private Sub ReCalcularImporteCambio()

    lblImporteEUR = ""
    txtImporte = Trim(txtImporte)
    If txtImporte = "" Then Exit Sub
    If Not IsNumeric(txtImporte) Then Exit Sub
    If Trim(txtCambio) <> "" Then
        If IsNumeric(txtCambio) Then
            lblImporteEUR = Round(txtImporte / txtCambio, 2)
        End If
    End If

End Sub

Private Sub txtCIF_Proveedor_GotFocus()
    TextGotFocus txtCIF_Proveedor
End Sub

Private Sub txtCIF_Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then LimpiarCampos
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtCIF_Proveedor_LostFocus()
    If LCase(Trim(txtCIF_Proveedor)) = "cpd" Then
        imgDatos_ProveedorCPD.Visible = True
    Else
        imgDatos_ProveedorCPD.Visible = False
    End If
    TextLostFocus txtCIF_Proveedor
End Sub

Private Sub txtConceptoGasto_GotFocus()
    'etAyudaConceptoGasto.Visible = True
    mfgTiposGasto.Visible = True
    txtConceptoGasto = sidGasto(msTipo, txtConceptoGasto)
    TextGotFocus txtConceptoGasto
End Sub

Private Sub txtConceptoGasto_KeyPress(KeyAscii As Integer)

    EntraTexto KeyAscii, txtConceptoGasto, 2, "0123456789", True
    If KeyAscii = 27 Then LimpiarCampos

End Sub

Private Sub txtConceptoGasto_LostFocus()
    'etAyudaConceptoGasto.Visible = False
    mfgTiposGasto.Visible = False
    TextLostFocus txtConceptoGasto
    txtConceptoGasto = sNombreGasto(msTipo, txtConceptoGasto)
    CamposPorConceptoGasto
End Sub

Private Sub txtDonde_GotFocus()
    TextGotFocus txtDonde
End Sub

Private Sub txtDonde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then LimpiarCampos
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtDonde_LostFocus()
    TextLostFocus txtDonde
End Sub

Private Sub txtFactura_GotFocus()
    TextGotFocus txtFactura
End Sub

Private Sub txtFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then LimpiarCampos
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtFactura_LostFocus()
    TextLostFocus txtFactura
End Sub

Private Sub txtFamiliaGasto_GotFocus()
    etAyudaFamiliaGasto.Visible = True
    txtFamiliaGasto = sFamiliaGasto(txtFamiliaGasto)
    TextGotFocus txtFamiliaGasto
End Sub

Private Sub txtFamiliaGasto_KeyPress(KeyAscii As Integer)

    EntraTexto KeyAscii, txtFamiliaGasto, 1, "1oO2fF3pP", True
    If KeyAscii = 27 Then LimpiarCampos
        
End Sub

Private Sub txtFamiliaGasto_LostFocus()
    etAyudaFamiliaGasto.Visible = False
    TextLostFocus txtFamiliaGasto
    txtFamiliaGasto = sFamiliaGasto(txtFamiliaGasto)
End Sub

Private Sub txtFecha_GotFocus()
    etFechaAyuda.Visible = True
    FechaGotFocus txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)

    EntraTexto KeyAscii, txtFecha, 8, Letras09
    If KeyAscii = 27 Then LimpiarCampos

End Sub

Private Sub txtFecha_LostFocus()

    etFechaAyuda.Visible = False
    FechaLostFocus txtFecha
    
    cmbDivisa_Click

'    If Format(txtFecha, "yyyymmdd") > Format(Date, "yyyymmdd") Then
'        MsgBox "La fecha del gasto no puede ser mayor que la fecha actual.", vbApplicationModal + vbInformation, "Información"
'        txtFecha = ""
'        txtFecha.SetFocus
'    End If

End Sub

Private Sub txtImporte_GotFocus()
    TextGotFocus txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then LimpiarCampos
    EntraNumero KeyAscii, txtImporte, 7, 2, "###0.00", "S"
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtImporte_LostFocus()
    ReCalcularImporteCambio
    TextLostFocus txtImporte
End Sub

Private Sub txtImporteIVA_GotFocus()
    TextGotFocus txtImporteIVA
End Sub

Private Sub txtImporteIVA_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    If KeyAscii = 27 Then LimpiarCampos
    EntraNumero KeyAscii, txtImporteIVA, 7, 2, "###0.00"
End Sub

Private Sub txtImporteIVA_LostFocus()
    TextLostFocus txtImporteIVA
End Sub

Private Sub txtImporteIVA2_GotFocus()
    TextGotFocus txtImporteIVA2
End Sub

Private Sub txtImporteIVA2_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    If KeyAscii = 27 Then LimpiarCampos
    EntraNumero KeyAscii, txtImporteIVA2, 7, 2, "###0.00"
End Sub

Private Sub txtImporteIVA2_LostFocus()
    TextLostFocus txtImporteIVA2
End Sub

Private Sub txtKms_GotFocus()
    TextGotFocus txtKms
End Sub

Private Sub txtKms_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    If KeyAscii = 27 Then LimpiarCampos
    EntraNumero KeyAscii, txtKms, 4
End Sub

Private Sub txtKms_LostFocus()
    TextLostFocus txtKms
End Sub

Private Sub txtNumNoches_GotFocus()
    TextGotFocus txtNumNoches
End Sub

Private Sub txtNumNoches_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    If KeyAscii = 27 Then LimpiarCampos
    EntraNumero KeyAscii, txtNumNoches, 2
End Sub

Private Sub txtNumNoches_LostFocus()
    TextLostFocus txtNumNoches
End Sub

Private Sub txtNumPersonas_GotFocus()
    TextGotFocus txtNumPersonas
End Sub

Private Sub txtNumPersonas_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    If KeyAscii = 27 Then LimpiarCampos
    EntraNumero KeyAscii, txtNumPersonas, 2
End Sub

Private Sub txtNumPersonas_LostFocus()
    TextLostFocus txtNumPersonas
End Sub

Private Sub txtNumTickets_GotFocus()
    TextGotFocus txtNumTickets
End Sub

Private Sub txtNumTickets_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    If KeyAscii = 27 Then LimpiarCampos
    EntraNumero KeyAscii, txtNumTickets, 2
End Sub

Private Sub txtNumTickets_LostFocus()
    TextLostFocus txtNumTickets
End Sub

Private Sub txtPEPCECO_GotFocus()
    TextGotFocus txtPEPCECO
End Sub

Private Sub txtPEPCECO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then LimpiarCampos
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtPEPCECO_LostFocus()
    TextLostFocus txtPEPCECO
End Sub

Private Sub txtPersonas_GotFocus()
    TextGotFocus txtPersonas
End Sub

Private Sub txtPersonas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then LimpiarCampos
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtPersonas_LostFocus()
    TextLostFocus txtPersonas
End Sub

Private Sub txtProveedor_GotFocus()
    TextGotFocus txtProveedor
End Sub

Private Sub txtProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then LimpiarCampos
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtProveedor_LostFocus()
    TextLostFocus txtProveedor
End Sub

Private Sub txtTipoComida_GotFocus()
    etAyudaTipoComida.Visible = True
    txtTipoComida = sTipoComidasInvitaciones(txtTipoComida)
    TextGotFocus txtTipoComida
End Sub

Private Sub txtTipoComida_KeyPress(KeyAscii As Integer)

    EntraTexto KeyAscii, txtTipoComida, 1, "1nN2iI", True
    If KeyAscii = 27 Then LimpiarCampos

End Sub

Private Sub txtTipoComida_LostFocus()
    etAyudaTipoComida.Visible = False
    TextLostFocus txtTipoComida
    txtTipoComida = sTipoComidasInvitaciones(txtTipoComida)
End Sub

Private Sub txtTipoIVA_GotFocus()
    TextGotFocus txtTipoIVA
End Sub

Private Sub txtTipoIVA_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    If KeyAscii = 27 Then LimpiarCampos
End Sub

Private Sub txtTipoIVA_LostFocus()
    TextLostFocus txtTipoIVA
End Sub

Private Sub txtTipoIVA2_GotFocus()
    TextGotFocus txtTipoIVA2
End Sub

Private Sub txtTipoIVA2_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    If KeyAscii = 27 Then LimpiarCampos
End Sub

Private Sub txtTipoIVA2_LostFocus()
    TextLostFocus txtTipoIVA2
End Sub

Private Sub txtTipoPago_GotFocus()
    etAyudaTipoPago.Visible = True
    txtTipoPago = sTipoPagoGasto(txtTipoPago)
    TextGotFocus txtTipoPago
End Sub

Private Sub txtTipoPago_KeyPress(KeyAscii As Integer)

    EntraTexto KeyAscii, txtTipoPago, 1, "1nN2sS3fF4rR", True
    If KeyAscii = 27 Then LimpiarCampos
        
End Sub

Private Sub txtTipoPago_LostFocus()
    etAyudaTipoPago.Visible = False
    TextLostFocus txtTipoPago
    txtTipoPago = sTipoPagoGasto(txtTipoPago)
End Sub

Private Sub MarcarTodas(bMarcar As Boolean)

    Dim iFila As Integer
    Dim sMarca As String
    Dim sMarca2 As String
    
    sMarca = ""
    sMarca2 = ""
    If bMarcar Then
        sMarca = "-->"
        sMarca2 = "<--"
    End If
    
    For iFila = 1 To moParrilla.Grid.Rows - 1
        moParrilla.TextoLinCol(iFila, 2) = sMarca
        moParrilla.TextoLinCol(iFila, 50) = sMarca2
    Next
    
End Sub

Private Function bComprobarDatosProveedorCPD() As Boolean

    bComprobarDatosProveedorCPD = False
    
    If lblCPD_Nombre = "" Then Exit Function
    If lblCPD_CP = "" Then Exit Function
    If lblCPD_Ciudad = "" Then Exit Function
    If lblCPD_Pais = "" Then Exit Function
    'El CIF no lo ponemos como obligatorio
    'If lblCPD_CIF = "" Then Exit Function
    
    bComprobarDatosProveedorCPD = True
    
End Function
