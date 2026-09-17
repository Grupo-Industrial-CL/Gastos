VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmJefesDepartamento 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13125
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
   ScaleHeight     =   5700
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txteMail 
      Height          =   285
      Left            =   2520
      MaxLength       =   150
      TabIndex        =   6
      Top             =   5160
      Width           =   5895
   End
   Begin VB.TextBox txtClave 
      Height          =   285
      Left            =   2040
      MaxLength       =   15
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   360
      MaxLength       =   20
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtDNI 
      Height          =   285
      Left            =   360
      MaxLength       =   25
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   3720
      MaxLength       =   50
      TabIndex        =   4
      Top             =   4440
      Width           =   4695
   End
   Begin VB.CommandButton cmdGuardarJefeDepartamento 
      Caption         =   "Añadir nuevo Jefe Departamento"
      Height          =   735
      Left            =   9240
      Picture         =   "frmJefesDepartamento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4560
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid mfgDepartamentos 
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid mfgJefesDepartamento 
      Height          =   2655
      Left            =   6480
      TabIndex        =   1
      Top             =   1200
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4683
      _Version        =   393216
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "eMail"
      Height          =   195
      Index           =   5
      Left            =   2520
      TabIndex        =   15
      Top             =   4920
      Width           =   435
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Clave"
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   14
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   13
      Top             =   4200
      Width           =   645
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "DNI"
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   12
      Top             =   4920
      Width           =   330
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      Height          =   195
      Index           =   3
      Left            =   3720
      TabIndex        =   11
      Top             =   4200
      Width           =   675
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Jefes Departamento"
      Height          =   195
      Index           =   0
      Left            =   6480
      TabIndex        =   10
      Top             =   960
      Width           =   1725
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Departamentos"
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1320
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   150
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmJefesDepartamento.frx":08CA
      ToolTipText     =   "Pica y mueve el formulario"
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   5280
      Picture         =   "frmJefesDepartamento.frx":1194
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Jefes Departamento"
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
      Width           =   3600
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   12480
      Picture         =   "frmJefesDepartamento.frx":149E
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
      Width           =   13140
   End
End
Attribute VB_Name = "frmJefesDepartamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mafoRegistro As ADODB.Recordset

Dim moDepartamentos As New cFlexGrid2
Dim moJefesDepartamento As New cFlexGrid2

Dim cMouseW As New cMouseWheel

Private Sub CargarDepartamentos()
    
    If bDameDepartamentosCL(CodPaisTrabajo, mafoRegistro, -1) Then
        While Not mafoRegistro.EOF
            moDepartamentos.TextoFila moDepartamentos.LineaLibre, mafoRegistro("CodDepartamento"), Trim(mafoRegistro("Departamento"))
            mafoRegistro.MoveNext
        Wend
    End If
    
End Sub

Private Sub CargarJefesDepartamento()

    moJefesDepartamento.Vaciar
    moJefesDepartamento.SituarEn 0, 1

    If moDepartamentos.bLineaActualVacia Then Exit Sub
    
    If bDameJefesDepartamento(CodPaisTrabajo, moDepartamentos.TextoLinCol(moDepartamentos.Grid.Row, 0), mafoRegistro) Then
        While Not mafoRegistro.EOF
            moJefesDepartamento.TextoFila moJefesDepartamento.LineaLibre, mafoRegistro("Codigo"), Trim(mafoRegistro("Nombre"))
            mafoRegistro.MoveNext
        Wend
    End If
    moDepartamentos.ResaltarLinea
    
End Sub

Private Sub cmdGuardarJefeDepartamento_Click()
    
    Dim sUsuario As String
    Dim sClave As String
    Dim sNombre As String
    Dim sDNI As String
    Dim seMail As String
    Dim sMensaje As String
    
    If moDepartamentos.LineaSeleccionada = -1 Then
        MsgBox "No ha seleccionado un departamento.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    
    sUsuario = LimpiarCadena(Trim(txtUsuario.Text))
    sClave = LimpiarCadena(Trim(txtClave.Text))
    sNombre = LimpiarCadena(Trim(txtNombre.Text))
    sDNI = LimpiarCadena(Trim(txtDNI.Text))
    seMail = LimpiarCadena(Trim(txteMail.Text))
    
    If sUsuario = "" Then
        txtUsuario.SetFocus
        Exit Sub
    End If
    If sClave = "" Then
        txtClave.SetFocus
        Exit Sub
    End If
    If sNombre = "" Then
        txtNombre.SetFocus
        Exit Sub
    End If
    If sDNI = "" Then
        txtDNI.SetFocus
        Exit Sub
    End If
    If seMail = "" Then
        txteMail.SetFocus
        Exit Sub
    End If
            
    If bInsertarJefeDepartamento(CodPaisTrabajo, CInt(moDepartamentos.TextoLinCol(moDepartamentos.LineaSeleccionada, 0)), sUsuario, sClave, sNombre, sDNI, seMail, sMensaje) Then
        
        CargarJefesDepartamento
        txtUsuario.Text = ""
        txtClave.Text = ""
        txtNombre.Text = ""
        txtDNI.Text = ""
        txteMail.Text = ""
        txtUsuario.SetFocus
        
    Else
        MsgBox "No se ha podido dar de alta el jefe de departamento. Error: " & sMensaje, vbApplicationModal + vbExclamation, "Error"
    End If
    
End Sub

Private Sub Form_Load()

    shRecuadroForm.Width = Me.Width
    shRecuadroForm.Height = Me.Height
    shRecuadroForm.Left = 0
    shRecuadroForm.Top = 0
    
    moDepartamentos.Create mfgDepartamentos, 2, 2
    moDepartamentos.Formatear 0, 0, "Codigo"
    moDepartamentos.Formatear 1, 5000, "Departamento", 4
    moDepartamentos.SituarEn 0, 1
    CargarDepartamentos
    
    moJefesDepartamento.Create mfgJefesDepartamento, 2, 2
    moJefesDepartamento.Formatear 0, 0, "Codigo"
    moJefesDepartamento.Formatear 1, 5000, "Jefe departamento", 4
    moJefesDepartamento.SituarEn 0, 1
        
    cMouseW.InitMouseWheel mfgDepartamentos.hwnd
    cMouseW.InitMouseWheel mfgJefesDepartamento.hwnd
    
    If CodPaisTrabajo = giPAIS_GES Then
    
        txtUsuario.Enabled = False
        txtClave.Enabled = False
        txtNombre.Enabled = False
        txtDNI.Enabled = False
        txteMail.Enabled = False
        cmdGuardarJefeDepartamento.Enabled = False
        
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub imgSalir_Click()
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub mfgDepartamentos_Click()
    If moDepartamentos.bLineaActualVacia Then
        moDepartamentos.SituarEn 0, 1
        Exit Sub
    End If
    CargarJefesDepartamento
    moDepartamentos.ResaltarLinea
End Sub

Private Sub mfgJefesDepartamento_Click()
    If moJefesDepartamento.bLineaActualVacia Then
        moJefesDepartamento.SituarEn 0, 1
        Exit Sub
    End If
    moJefesDepartamento.ResaltarLinea
End Sub

Private Sub txtClave_GotFocus()
    TextGotFocus txtClave
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtClave_LostFocus()
    TextLostFocus txtClave
End Sub

Private Sub txtDNI_GotFocus()
    TextGotFocus txtDNI
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtDNI_LostFocus()
    TextLostFocus txtDNI
End Sub

Private Sub txteMail_GotFocus()
    TextGotFocus txteMail
End Sub

Private Sub txteMail_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txteMail_LostFocus()
    TextLostFocus txteMail
End Sub

Private Sub txtNombre_GotFocus()
    TextGotFocus txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtNombre_LostFocus()
    TextLostFocus txtNombre
End Sub

Private Sub txtUsuario_GotFocus()
    TextGotFocus txtUsuario
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtUsuario_LostFocus()
    TextLostFocus txtUsuario
End Sub
