VERSION 5.00
Begin VB.Form frmArchivarEn 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
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
   ScaleHeight     =   2985
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtArchivarEn 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Shape shRecuadro 
      Height          =   2980
      Left            =   0
      Top             =   0
      Width           =   6000
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   3960
      Picture         =   "frmArchivarEn.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Caption         =   "Los justificantes de gastos se archivan en"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5700
   End
   Begin VB.Image imgAceptar 
      Height          =   480
      Left            =   2760
      Picture         =   "frmArchivarEn.frx":030A
      ToolTipText     =   "Aceptar"
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Archivar en"
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
      Width           =   2025
   End
   Begin VB.Image imgIcono 
      Height          =   480
      Left            =   120
      Picture         =   "frmArchivarEn.frx":0BD4
      Top             =   80
      Width           =   480
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   5400
      Picture         =   "frmArchivarEn.frx":149E
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
      Width           =   6000
   End
End
Attribute VB_Name = "frmArchivarEn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msArchivarEn As String

Public Property Get ArchivarEn() As String
    ArchivarEn = msArchivarEn
End Property

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

    msArchivarEn = ""

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub imgAceptar_Click()

    Me.MousePointer = DEFAULT
    
    txtArchivarEn = Trim(Replace(txtArchivarEn.Text, "'", ""))
    If txtArchivarEn = "" Then
        txtArchivarEn.SetFocus
        Exit Sub
    End If
    msArchivarEn = txtArchivarEn
    
    Unload Me

End Sub

Private Sub imgAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgSalir_Click()
    msArchivarEn = ""
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub txtArchivarEn_GotFocus()
    TextGotFocus txtArchivarEn
End Sub

Private Sub txtArchivarEn_LostFocus()
    TextLostFocus txtArchivarEn
End Sub
