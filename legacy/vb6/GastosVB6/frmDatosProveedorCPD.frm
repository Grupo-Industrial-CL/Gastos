VERSION 5.00
Begin VB.Form frmDatosProveedorCPD 
   BorderStyle     =   0  'None
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
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
   ScaleHeight     =   3645
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   735
      Left            =   5760
      Picture         =   "frmDatosProveedorCPD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtCP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1620
      Width           =   1575
   End
   Begin VB.TextBox txtCiudad 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   2
      Top             =   2120
      Width           =   6255
   End
   Begin VB.TextBox txtPais 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   3
      Top             =   2580
      Width           =   495
   End
   Begin VB.TextBox txtCIF 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   16
      TabIndex        =   4
      Top             =   3080
      Width           =   2415
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   0
      Top             =   1160
      Width           =   6255
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "CP"
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
      Index           =   4
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Ciudad"
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
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "País"
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
      Top             =   2640
      Width           =   390
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "CIF"
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
      TabIndex        =   8
      Top             =   3120
      Width           =   330
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
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
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.Image imgIcono 
      Height          =   480
      Left            =   180
      Picture         =   "frmDatosProveedorCPD.frx":08CA
      Top             =   100
      Width           =   480
   End
   Begin VB.Label lblTexto 
      BackColor       =   &H00E2C07A&
      Caption         =   "Datos contabilización CPD"
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
      Height          =   495
      Index           =   8
      Left            =   960
      TabIndex        =   6
      Top             =   100
      Width           =   4935
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   6120
      Picture         =   "frmDatosProveedorCPD.frx":1194
      Top             =   120
      Visible         =   0   'False
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
      Width           =   8265
   End
End
Attribute VB_Name = "frmDatosProveedorCPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msNombre As String
Dim msCP As String
Dim msCiudad As String
Dim msPais As String
Dim msCIF As String

Public Sub Create(sNombre As String, sCP As String, sCiudad As String, sPais As String, sCIF As String)
    msNombre = sNombre
    msCP = sCP
    msCiudad = sCiudad
    msPais = sPais
    msCIF = sCIF
End Sub

Private Sub cmdAceptar_Click()

    txtNombre = Replace(txtNombre, "'", "")
    txtNombre = Trim(txtNombre)
    If txtNombre = "" Then
        txtNombre.SetFocus
        Exit Sub
    End If
    txtCP = Replace(txtCP, "'", "")
    txtCP = Trim(txtCP)
    If txtCP = "" Then
        txtCP.SetFocus
        Exit Sub
    End If
    txtCiudad = Replace(txtCiudad, "'", "")
    txtCiudad = Trim(txtCiudad)
    If txtCiudad = "" Then
        txtCiudad.SetFocus
        Exit Sub
    End If
    txtPais = Replace(txtPais, "'", "")
    txtPais = Trim(txtPais)
    If txtPais = "" Then
        txtPais.SetFocus
        Exit Sub
    End If
    txtCIF = Replace(txtCIF, "'", "")
    txtCIF = Trim(txtCIF)
    'Ponemos el CIF como opcional
'    If txtCIF = "" Then
'        txtCIF.SetFocus
'        Exit Sub
'    End If
    
    msNombre = UCase(txtNombre)
    msCP = UCase(txtCP)
    msCiudad = UCase(txtCiudad)
    msPais = UCase(txtPais)
    msCIF = UCase(txtCIF)
    
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

    shRecuadroForm.Width = Me.Width - 5
    shRecuadroForm.Height = Me.Height - 5
    shRecuadroForm.Left = 5
    shRecuadroForm.Top = 5
    
    txtNombre = msNombre
    txtCP = msCP
    txtCiudad = msCiudad
    txtPais = msPais
    txtCIF = msCIF
        
End Sub

Private Sub txtCIF_GotFocus()
    TextGotFocus txtCIF
End Sub

Private Sub txtCIF_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtCIF_LostFocus()
    TextLostFocus txtCIF
End Sub

Private Sub txtCiudad_GotFocus()
    TextGotFocus txtCiudad
End Sub

Private Sub txtCiudad_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtCiudad_LostFocus()
    TextLostFocus txtCiudad
End Sub

Private Sub txtCP_GotFocus()
    TextGotFocus txtCP
End Sub

Private Sub txtCP_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtCP_LostFocus()
    TextLostFocus txtCP
End Sub

Private Sub txtNombre_GotFocus()
    TextGotFocus txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtNombre_LostFocus()
    TextLostFocus txtNombre
End Sub

Private Sub txtPais_GotFocus()
    TextGotFocus txtPais
End Sub

Private Sub txtPais_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtPais_LostFocus()
    TextLostFocus txtPais
End Sub

Public Property Get Nombre() As String
    Nombre = msNombre
End Property

Public Property Get CP() As String
    CP = msCP
End Property

Public Property Get Ciudad() As String
    Ciudad = msCiudad
End Property

Public Property Get Pais() As String
    Pais = msPais
End Property

Public Property Get CIF() As String
    CIF = msCIF
End Property
