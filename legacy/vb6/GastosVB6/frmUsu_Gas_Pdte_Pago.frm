VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUsu_Gas_Pdte_Pago 
   BorderStyle     =   0  'None
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13890
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   13890
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTipo 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   4200
      TabIndex        =   7
      Top             =   4680
      Width           =   1575
      Begin VB.OptionButton optAnticipos 
         Caption         =   "Anticipos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optGastos 
         Caption         =   "Gastos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.OptionButton optNoValidado 
      Caption         =   "No validado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   5760
      Width           =   3135
   End
   Begin VB.OptionButton optValidadoNoPagado 
      Caption         =   "Validado y no pagado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   5400
      Width           =   3135
   End
   Begin VB.OptionButton optValidadoNoContabilizado 
      Caption         =   "Validado y no contabilizado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   5040
      Value           =   -1  'True
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid mfgParrilla 
      Height          =   3015
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   5318
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgBuscar 
      Height          =   720
      Left            =   6120
      Picture         =   "frmUsu_Gas_Pdte_Pago.frx":0000
      ToolTipText     =   "Buscar"
      Top             =   5400
      Width           =   720
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Realizar búsqueda por"
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
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   2160
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   840
      Left            =   0
      Top             =   15
      Width           =   250
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Doble click para visualizar los gastos/anticipos del empleado"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   5955
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   13200
      Picture         =   "frmUsu_Gas_Pdte_Pago.frx":0ECA
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Buscar empleados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   8
      Left            =   720
      TabIndex        =   0
      Top             =   75
      Width           =   3270
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   4440
      Picture         =   "frmUsu_Gas_Pdte_Pago.frx":13D6
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmUsu_Gas_Pdte_Pago.frx":16E0
      Top             =   60
      Width           =   480
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
      Width           =   13905
   End
End
Attribute VB_Name = "frmUsu_Gas_Pdte_Pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlCodigo As Long
Dim miDepartamento As Integer
Dim msFecha As String
Dim mbEsBaja As Boolean

Dim moParrilla As New CFlexGrid

Private Sub Form_Load()
    
    If DisplayMonitorCount = 1 Then
        Me.Left = (Screen.Width / 2) - (Me.Width / 2)
        Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Else
        If MDIPrincipal.Left + 120 < Screen.Width Then
        'El MDI está en la pantalla 1, lo centro en la pantalla 1
            Me.Left = (Screen.Width / 2) - (Me.Width / 2)
            Me.Top = (Screen.Height / 2) - (Me.Height / 2)
        Else
            Me.Left = gAnchoPantallaPrincipal + (((gAnchoPantallaVirtual - gAnchoPantallaPrincipal) - Me.Width) \ 2)
            Me.Top = (Screen.Height - Me.Height) \ 2
        End If
    End If
    
    shRecuadroForm.Width = Me.Width
    shRecuadroForm.Height = Me.Height
    shRecuadroForm.Left = 0
    shRecuadroForm.Top = 0
    
    mlCodigo = -1
    miDepartamento = -1
    msFecha = ""
    mbEsBaja = False

    moParrilla.Create mfgParrilla, 8, 2
    moParrilla.Formatear 0, 0, "Codigo"
    moParrilla.Formatear 1, 0, "Departamento"
    moParrilla.Formatear 2, 0, "Fecha primer gasto encontrado"
    moParrilla.Formatear 3, 0, "Baja"
    moParrilla.Formatear 4, 2000, "Departamento", 4, , True
    moParrilla.Formatear 5, 4000, "Empleado", 4, , True
    moParrilla.Formatear 6, 3000, "Fecha primer gasto encontrado", 4, 4, True
    moParrilla.Formatear 7, 3000, "Fecha en la que fue guardado", 4, 4, True
    moParrilla.SituarEn 0, 1
    
    RellenarDatos

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub fraTipo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub imgBuscar_Click()
    Me.MousePointer = DEFAULT
    RellenarDatos
End Sub

Private Sub imgBuscar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgSalir_Click()
    mlCodigo = -1
    miDepartamento = -1
    msFecha = ""
    mbEsBaja = False
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub mfgParrilla_Click()
    
    If moParrilla.Grid.MouseRow = 0 Then
        moParrilla.OrdenarSiCorresponde
        moParrilla.SituarEn 0, 1
        Exit Sub
    End If
    
    If moParrilla.bLineaActualVacia Then
        moParrilla.SituarEn 0, 1
        Exit Sub
    End If
    moParrilla.ResaltarLinea

End Sub

Private Sub mfgParrilla_DblClick()
    If moParrilla.bLineaActualVacia Then
        moParrilla.SituarEn 0, 1
        Exit Sub
    End If
    mlCodigo = moParrilla.TextoLinCol(moParrilla.Grid.Row, 0)
    miDepartamento = moParrilla.TextoLinCol(moParrilla.Grid.Row, 1)
    msFecha = moParrilla.TextoLinCol(moParrilla.Grid.Row, 2)
    mbEsBaja = False
    If moParrilla.TextoLinCol(moParrilla.Grid.Row, 3) = "S" Then mbEsBaja = True
    Unload Me
End Sub

Public Property Get Codigo() As Long
    Codigo = mlCodigo
End Property

Public Property Get Departamento() As Integer
    Departamento = miDepartamento
End Property

Public Property Get Fecha() As String
    Fecha = msFecha
End Property

Public Property Get Baja() As Boolean
    Baja = mbEsBaja
End Property

Private Sub RellenarDatos()

    Dim afoRegistro As ADODB.Recordset
    Dim sTipoBusqueda As String
    
    moParrilla.Vaciar
    moParrilla.SituarEn 0, 1

    If optGastos.Value = True Then sTipoBusqueda = "gastos"
    If optAnticipos.Value = True Then sTipoBusqueda = "anticipos"

    If optValidadoNoContabilizado.Value = True Then
        If Not bDameEmpleados_GastosValidados_NoContabilizados(CodPaisTrabajo, afoRegistro, sTipoBusqueda) Then
            MsgBox "No se han encontrado datos.", vbApplicationModal + vbInformation, "Información"
            Exit Sub
        End If
    End If
    If optValidadoNoPagado.Value = True Then
        If Not bDameEmpleados_GastosValidados_NoPagados(CodPaisTrabajo, afoRegistro, sTipoBusqueda) Then
            MsgBox "No se han encontrado datos.", vbApplicationModal + vbInformation, "Información"
            Exit Sub
        End If
    End If
    If optNoValidado.Value = True Then
        If Not bDameEmpleados_GastosNoValidados(CodPaisTrabajo, afoRegistro, sTipoBusqueda) Then
            MsgBox "No se han encontrado datos.", vbApplicationModal + vbInformation, "Información"
            Exit Sub
        End If
    End If
        
    While Not afoRegistro.EOF
        moParrilla.TextoFila moParrilla.LineaLibre, afoRegistro("Codigo"), afoRegistro("Departamento"), Format(afoRegistro("Fecha_Ini"), "dd/MM/yyyy"), _
            Trim(afoRegistro("EsBaja")), Trim(afoRegistro("Nom_Departamento")), Trim(afoRegistro("Nom_Empleado")), _
            Format(afoRegistro("Fecha_Ini"), "dd/MM/yyyy"), Format(afoRegistro("MinFechaInsert"), "dd/MM/yyyy")
        afoRegistro.MoveNext
    Wend
    Set afoRegistro = Nothing
    moParrilla.SituarEn 0, 1

End Sub
