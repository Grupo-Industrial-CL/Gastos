VERSION 5.00
Begin VB.Form frmPedirMotivoModificacion 
   BorderStyle     =   0  'None
   ClientHeight    =   3240
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
   ScaleHeight     =   3240
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMotivoModificacion 
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   8660
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "A C E P T A R"
      Height          =   855
      Left            =   360
      Picture         =   "frmPedirMotivoModificacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   8660
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   240
      Picture         =   "frmPedirMotivoModificacion.frx":08CA
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Motivo de modificación"
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
      Width           =   4170
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Introduzca el motivo por el que ha modificado el gasto"
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
      Width           =   5415
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
Attribute VB_Name = "frmPedirMotivoModificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msMotivoModificacion As String

Private Sub cmdAceptar_Click()

    txtMotivoModificacion = Trim(txtMotivoModificacion)
    txtMotivoModificacion = Replace(txtMotivoModificacion, "'", "")

    If txtMotivoModificacion = "" Then
        txtMotivoModificacion.SetFocus
        Exit Sub
    End If
    
    msMotivoModificacion = txtMotivoModificacion
    
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

    msMotivoModificacion = ""

End Sub

Private Sub txtMotivoModificacion_GotFocus()
    TextGotFocus txtMotivoModificacion
End Sub

Private Sub txtMotivoModificacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtMotivoModificacion_LostFocus()
    TextLostFocus txtMotivoModificacion
End Sub

Public Property Get MotivoModificacion() As String
    MotivoModificacion = msMotivoModificacion
End Property

