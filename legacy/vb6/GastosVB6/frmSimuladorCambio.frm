VERSION 5.00
Begin VB.Form frmSimuladorCambio 
   BorderStyle     =   0  'None
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
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
   ScaleHeight     =   4575
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "A C E P T A R"
      Height          =   855
      Left            =   360
      Picture         =   "frmSimuladorCambio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   7335
   End
   Begin VB.TextBox txtOtroCambio 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   2640
      TabIndex        =   3
      Top             =   2800
      Width           =   1335
   End
   Begin VB.OptionButton optOtroCambio 
      Caption         =   "Otro cambio"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.OptionButton optCambioSistema 
      Caption         =   "Cambio sistema"
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
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.OptionButton optCambioActual 
      Caption         =   "Cambio actual"
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
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.Label lblImporteEUR 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5280
      TabIndex        =   14
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblCambioSistema 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   13
      Top             =   2200
      Width           =   1335
   End
   Begin VB.Label lblCambioActual 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   12
      Top             =   1600
      Width           =   1335
   End
   Begin VB.Label lblDivisa 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6765
      TabIndex        =   7
      Top             =   915
      Width           =   735
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   6
      Top             =   915
      Width           =   1335
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1080
      TabIndex        =   5
      Top             =   920
      Width           =   1455
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Divisa"
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
      Index           =   2
      Left            =   6120
      TabIndex        =   11
      Top             =   960
      Width           =   570
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Importe"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   960
      Width           =   765
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
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
      TabIndex        =   9
      Top             =   960
      Width           =   600
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   5640
      Picture         =   "frmSimuladorCambio.frx":08CA
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   7560
      Picture         =   "frmSimuladorCambio.frx":0BD4
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   200
      Picture         =   "frmSimuladorCambio.frx":10E0
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Simulador tipo cambio"
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
      TabIndex        =   8
      Top             =   75
      Width           =   4035
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
      Width           =   8145
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Caption         =   "Importe aplicando cambio seleccionado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   27
      Left            =   4920
      TabIndex        =   15
      Top             =   1800
      Width           =   2670
   End
End
Attribute VB_Name = "frmSimuladorCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msFecha As String
Dim mdImporte As Double
Dim msDivisa As String
Dim mdCambioActual As Double
Dim msMonedaDepartamento As String

Dim mdCambioSeleccionado As Double

Public Property Get CambioSeleccionado() As Double
    CambioSeleccionado = mdCambioSeleccionado
End Property

Public Sub Create(sFecha As String, dImporte As Double, sDivisa As String, dCambioActual As Double, sMonedaDepartamento As String)

    msFecha = sFecha
    mdImporte = dImporte
    msDivisa = sDivisa
    mdCambioActual = dCambioActual
    mdCambioSeleccionado = -999999.9999
    msMonedaDepartamento = sMonedaDepartamento

End Sub

Private Sub cmdAceptar_Click()

    If optCambioActual.Value = True Then mdCambioSeleccionado = lblCambioActual
    If optCambioSistema.Value = True Then mdCambioSeleccionado = lblCambioSistema
    If optOtroCambio.Value = True Then
        
        txtOtroCambio = Trim(txtOtroCambio)
        If txtOtroCambio = "" Then
            txtOtroCambio.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtOtroCambio) Then
            txtOtroCambio = ""
            txtOtroCambio.SetFocus
            Exit Sub
        End If
        mdCambioSeleccionado = txtOtroCambio
        
    End If
    
    Unload Me

End Sub

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

    lblFecha = msFecha
    lblImporte = mdImporte
    lblDivisa = msDivisa
    lblCambioActual = mdCambioActual
    lblCambioSistema = dDameCambioDivisa(CodPaisTrabajo, msDivisa, msFecha, msMonedaDepartamento)
    ReCalcularImporte

End Sub

Private Sub imgSalir_Click()
    mdCambioSeleccionado = -999999.9999
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub optCambioActual_Click()
    txtOtroCambio.Enabled = False
    ReCalcularImporte
End Sub

Private Sub optCambioSistema_Click()
    txtOtroCambio.Enabled = False
    ReCalcularImporte
End Sub

Private Sub optOtroCambio_Click()

    txtOtroCambio.Enabled = True
    txtOtroCambio.SetFocus
    ReCalcularImporte

End Sub

Private Sub ReCalcularImporte()
    
    Dim dCambio As Double
    
    lblImporteEUR = ""
    
    If optCambioActual.Value = True Then dCambio = lblCambioActual
    If optCambioSistema.Value = True Then dCambio = lblCambioSistema
    If optOtroCambio.Value = True Then
        
        txtOtroCambio = Trim(txtOtroCambio)
        If txtOtroCambio = "" Then
            txtOtroCambio.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtOtroCambio) Then
            txtOtroCambio = ""
            txtOtroCambio.SetFocus
            Exit Sub
        End If
        dCambio = txtOtroCambio
        
    End If
    
    lblImporteEUR = Round(lblImporte / dCambio, 2)
    
End Sub

Private Sub txtOtroCambio_GotFocus()
    TextGotFocus txtOtroCambio
End Sub

Private Sub txtOtroCambio_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, txtOtroCambio, 7, 4, "###0.000", "S"
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtOtroCambio_LostFocus()
    TextLostFocus txtOtroCambio
    ReCalcularImporte
End Sub
