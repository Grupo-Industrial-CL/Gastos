VERSION 5.00
Begin VB.Form frmPedirMotivoRechazo 
   BorderStyle     =   0  'None
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9345
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
   ScaleHeight     =   3975
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelarRechazo 
      Caption         =   "C A N C E L A R     R E C H A Z O"
      Height          =   855
      Left            =   360
      Picture         =   "frmPedirMotivoRechazo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   8660
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "A C E P T A R"
      Height          =   855
      Left            =   360
      Picture         =   "frmPedirMotivoRechazo.frx":050C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   8660
   End
   Begin VB.TextBox txtMotivoRechazo 
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   8660
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Introduzca el motivo de rechazo del gasto"
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
      TabIndex        =   4
      Top             =   1080
      Width           =   4200
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Motivo de rechazo"
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
      TabIndex        =   3
      Top             =   75
      Width           =   3285
   End
   Begin VB.Image imgIcono 
      Height          =   480
      Left            =   120
      Picture         =   "frmPedirMotivoRechazo.frx":0DD6
      Top             =   40
      Width           =   480
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   150
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
      Width           =   9345
   End
End
Attribute VB_Name = "frmPedirMotivoRechazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msMotivoRechazo As String

Private Sub cmdAceptar_Click()

    txtMotivoRechazo = Trim(txtMotivoRechazo)
    txtMotivoRechazo = Replace(txtMotivoRechazo, "'", "")

    If txtMotivoRechazo = "" Then
        txtMotivoRechazo.SetFocus
        Exit Sub
    End If
    
    msMotivoRechazo = txtMotivoRechazo
    
    Unload Me
    
End Sub

Private Sub cmdCancelarRechazo_Click()
    msMotivoRechazo = "@@CANCELADO##"
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

    msMotivoRechazo = ""

End Sub

Private Sub txtMotivoRechazo_GotFocus()
    TextGotFocus txtMotivoRechazo
End Sub

Private Sub txtMotivoRechazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtMotivoRechazo_LostFocus()
    TextLostFocus txtMotivoRechazo
End Sub

Public Property Get MotivoRechazo() As String
    MotivoRechazo = msMotivoRechazo
End Property
