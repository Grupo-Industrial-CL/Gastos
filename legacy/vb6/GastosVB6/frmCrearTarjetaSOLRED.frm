VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCrearTarjetaSOLRED 
   BorderStyle     =   0  'None
   Caption         =   "Tarjeta SOLRED"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13140
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   13140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTarjeta 
      Height          =   285
      Left            =   240
      MaxLength       =   16
      TabIndex        =   3
      Top             =   4365
      Width           =   3855
   End
   Begin VB.CommandButton cmdGuardarTarjeta 
      Caption         =   "Añadir nueva tarjeta"
      Height          =   735
      Left            =   4440
      Picture         =   "frmCrearTarjetaSOLRED.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4080
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid mfgDepartamentos 
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid mfgUsuarios 
      Height          =   2655
      Left            =   6480
      TabIndex        =   1
      Top             =   1080
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4683
      _Version        =   393216
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Tarjeta Solred"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   4125
      Width           =   1230
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Departamentos"
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1320
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Usuarios"
      Height          =   195
      Index           =   0
      Left            =   6480
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Image imgIcono 
      Height          =   720
      Left            =   40
      Picture         =   "frmCrearTarjetaSOLRED.frx":08CA
      Top             =   -40
      Width           =   720
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   150
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   5280
      Picture         =   "frmCrearTarjetaSOLRED.frx":2594
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Tarjetas SOLRED"
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
      Left            =   960
      TabIndex        =   2
      Top             =   75
      Width           =   3030
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   12480
      Picture         =   "frmCrearTarjetaSOLRED.frx":289E
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
Attribute VB_Name = "frmCrearTarjetaSOLRED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mafoRegistro As ADODB.Recordset

Dim moDepartamentos As New cFlexGrid2
Dim moUsuarios As New cFlexGrid2

Dim cMouseW As New cMouseWheel

Private Sub CargarDepartamentos()
    
    If bDameDepartamentosCL(CodPaisTrabajo, mafoRegistro, -1) Then
        While Not mafoRegistro.EOF
            moDepartamentos.TextoFila moDepartamentos.LineaLibre, mafoRegistro("CodDepartamento"), Trim(mafoRegistro("Departamento"))
            mafoRegistro.MoveNext
        Wend
    End If
    
End Sub

Private Sub CargarUsuarios()

    moUsuarios.Vaciar
    moUsuarios.SituarEn 0, 1

    If moDepartamentos.bLineaActualVacia Then Exit Sub
        
    If bDameEmpleadosDepartamento(CodPaisTrabajo, moDepartamentos.TextoLinCol(moDepartamentos.Grid.Row, 0), mafoRegistro, "N", 0, "N") Then
        While Not mafoRegistro.EOF
            moUsuarios.TextoFila moUsuarios.LineaLibre, mafoRegistro("Codigo"), Trim(mafoRegistro("Jerarquia")), Trim(mafoRegistro("Nombre"))
            mafoRegistro.MoveNext
        Wend
    End If
    moDepartamentos.ResaltarLinea
    
End Sub

Private Sub cmdGuardarTarjeta_Click()
    
    Dim sTarjeta As String
    Dim sMensaje As String
    
    If moDepartamentos.LineaSeleccionada = -1 Then
        MsgBox "No ha seleccionado un departamento.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    If moUsuarios.LineaSeleccionada = -1 Then
        MsgBox "No ha seleccionado un usuario.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    
    sTarjeta = LimpiarCadena(Trim(txtTarjeta.Text))
    If sTarjeta = "" Then
        txtTarjeta.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(sTarjeta) Then
        MsgBox "La tarjeta debe ser numérica.", vbApplicationModal + vbExclamation, "Información"
        txtTarjeta.Text = ""
        txtTarjeta.SetFocus
        Exit Sub
    End If
            
    If bInsertarTarjeta(CodPaisTrabajo, CInt(moDepartamentos.TextoLinCol(moDepartamentos.LineaSeleccionada, 0)), CLng(moUsuarios.TextoLinCol(moUsuarios.LineaSeleccionada, 0)), moUsuarios.TextoLinCol(moUsuarios.LineaSeleccionada, 2), moUsuarios.TextoLinCol(moUsuarios.LineaSeleccionada, 1), sTarjeta, sMensaje) Then
        txtTarjeta.Text = ""
        txtTarjeta.SetFocus
    Else
        MsgBox "No se ha podido dar de alta la tarjeta. Error: " & sMensaje, vbApplicationModal + vbExclamation, "Error"
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
    
    moUsuarios.Create mfgUsuarios, 3, 2
    moUsuarios.Formatear 0, 0, "Codigo"
    moUsuarios.Formatear 1, 0, "Jerarquia"
    moUsuarios.Formatear 2, 5000, "Usuario", 4
    moUsuarios.SituarEn 0, 1
        
    cMouseW.InitMouseWheel mfgDepartamentos.hwnd
    cMouseW.InitMouseWheel mfgUsuarios.hwnd
    
    If CodPaisTrabajo = giPAIS_GES Then
    
        txtTarjeta.Enabled = False
        cmdGuardarTarjeta.Enabled = False
        
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
    CargarUsuarios
    moDepartamentos.ResaltarLinea
End Sub

Private Sub mfgUsuarios_Click()
    If moUsuarios.bLineaActualVacia Then
        moUsuarios.SituarEn 0, 1
        Exit Sub
    End If
    moUsuarios.ResaltarLinea
End Sub

Private Sub txtTarjeta_GotFocus()
    TextGotFocus txtTarjeta
End Sub

Private Sub txtTarjeta_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, txtTarjeta, 16, 0
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtTarjeta_LostFocus()
    TextLostFocus txtTarjeta
End Sub
