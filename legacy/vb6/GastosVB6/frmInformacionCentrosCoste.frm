VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInformacionCentrosCoste 
   BorderStyle     =   0  'None
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
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
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPersonal 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   17415
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
         TabIndex        =   3
         Top             =   200
         Width           =   5775
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "CONSULTAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   15120
         Picture         =   "frmInformacionCentrosCoste.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   200
         Width           =   2055
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
         TabIndex        =   4
         Top             =   240
         Width           =   1395
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfgParrilla 
      Height          =   6135
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   10821
      _Version        =   393216
   End
   Begin VB.Image imgExportar 
      Height          =   480
      Left            =   17880
      Picture         =   "frmInformacionCentrosCoste.frx":0ECA
      ToolTipText     =   "Copiar al Portapapeles"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmInformacionCentrosCoste.frx":1794
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   14400
      Picture         =   "frmInformacionCentrosCoste.frx":205E
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Informe gastos/anticipos personal interno"
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
      TabIndex        =   0
      Top             =   75
      Width           =   7605
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   18000
      Picture         =   "frmInformacionCentrosCoste.frx":2368
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
      Width           =   18720
   End
End
Attribute VB_Name = "frmInformacionCentrosCoste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mafoRegistro As ADODB.Recordset

Dim moParrilla As New CFlexGrid

Dim cMouseW As New cMouseWheel

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
    
'sociedad    pepceco descripcion fecha_aceptacion    fecha_cierre    responsable tipo    director_proyecto   nomresponsable  nomdirectorproyecto
'0610    CB-0006I010103              Jefe de Obra Civl   04-feb-2018 31-dic-9999 9927    PEP 3630    VICTOR BERNARDO PACHECO                             VICTOR M. NOGUEIRA DOMINGUEZ

    
    moParrilla.Create mfgParrilla, 15, 2
    moParrilla.Formatear 0, 0, "Cod. Sociedad"
    moParrilla.Formatear 1, 1300, "Sociedad", 4, 7, True
    moParrilla.Formatear 2, 3500, "Nombre del empleado o candidato", 4, , True
    moParrilla.Formatear 3, 1400, "Nº sociedad", 4, 7, True
    moParrilla.Formatear 4, 1600, "Centro Coste", 4, , True
    moParrilla.Formatear 5, 2300, "Clase gasto viaje", 4, , True
    moParrilla.Formatear 6, 1800, "Impte. mon. local", 4, 7, True
    moParrilla.Formatear 7, 1000, "Moneda", 4, 4, True
    moParrilla.Formatear 8, 1800, "Importe €", 4, 7, True
    moParrilla.Formatear 9, 1300, "Cambio", 4, 7
    moParrilla.Formatear 10, 1600, "Fe. comprobante", 4, 4, True
    moParrilla.Formatear 11, 4000, "Descripción", 4
    moParrilla.Formatear 12, 1000, "Validado", 4, 4, True
    moParrilla.Formatear 13, 1250, "Conta. SAP", 4, 4, True
    moParrilla.Formatear 14, 1000, "Pagado", 4, 4, True
    moParrilla.SituarEn 0, 1
    
    cmbDepartamentos.Tag = "C"
    If bDameDepartamentosCL(CodPaisTrabajo, mafoRegistro, 0) Then
        While Not mafoRegistro.EOF
            cmbDepartamentos.AddItem Trim(mafoRegistro("Departamento"))
            cmbDepartamentos.ItemData(cmbDepartamentos.NewIndex) = mafoRegistro("CodDepartamento")
            mafoRegistro.MoveNext
        Wend
    End If
    cmbDepartamentos.Tag = ""
    Set mafoRegistro = Nothing
    
    cMouseW.InitMouseWheel mfgParrilla.hwnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub imgExportar_Click()
    Me.MousePointer = DEFAULT
    moParrilla.SituarEn 0, 1
    If moParrilla.bLineaActualVacia Then Exit Sub
    moParrilla.aClipBoard 1
    MsgBox "Información copiada al Portapapeles.", vbApplicationModal + vbInformation, "Información"
End Sub

Private Sub imgExportar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgSalir_Click()
    Me.MousePointer = DEFAULT
    Set mafoRegistro = Nothing
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

