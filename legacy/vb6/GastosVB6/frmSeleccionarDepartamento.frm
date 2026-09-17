VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSeleccionarDepartamento 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
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
   ScaleHeight     =   5190
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Picture         =   "frmSeleccionarDepartamento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4320
      Width           =   5175
   End
   Begin MSFlexGridLib.MSFlexGrid mfgDepartamentos 
      Height          =   3135
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5530
      _Version        =   393216
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
   Begin VB.Shape shRecuadroForm 
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   150
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   4440
      Picture         =   "frmSeleccionarDepartamento.frx":08CA
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   5280
      Picture         =   "frmSeleccionarDepartamento.frx":0BD4
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Departamentos"
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
      Index           =   13
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Departamentos"
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
      Width           =   2760
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmSeleccionarDepartamento.frx":10E0
      ToolTipText     =   "Pica y mueve el formulario"
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
      Width           =   5940
   End
End
Attribute VB_Name = "frmSeleccionarDepartamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mafoRegistro As ADODB.Recordset

Dim moDepartamentos As New cFlexGrid2

Dim miCodDepartamento As Integer

Dim cMouseW As New cMouseWheel

Public Property Get CodDepartamento() As Integer
    CodDepartamento = miCodDepartamento
End Property

Private Sub CargarDepartamentos()
    
    If bDameDepartamentosCL(CodPaisTrabajo, mafoRegistro, -1) Then
        While Not mafoRegistro.EOF
            moDepartamentos.TextoFila moDepartamentos.LineaLibre, mafoRegistro("CodDepartamento"), Trim(mafoRegistro("Departamento"))
            mafoRegistro.MoveNext
        Wend
    End If
    
End Sub

Private Sub cmdAceptar_Click()
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub Form_Load()

    shRecuadroForm.Width = Me.Width
    shRecuadroForm.Height = Me.Height
    shRecuadroForm.Left = 0
    shRecuadroForm.Top = 0
    
    moDepartamentos.Create mfgDepartamentos, 2, 2
    moDepartamentos.Formatear 0, 0, "Codigo"
    moDepartamentos.Formatear 1, 4500, "Departamento", 4
    moDepartamentos.SituarEn 0, 1
    CargarDepartamentos
    
    cMouseW.InitMouseWheel mfgDepartamentos.hwnd
    
    miCodDepartamento = 0
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub imgSalir_Click()
    miCodDepartamento = 0
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub mfgDepartamentos_Click()
    If moDepartamentos.bLineaActualVacia Then
        miCodDepartamento = 0
        moDepartamentos.SituarEn 0, 1
        Exit Sub
    End If
    moDepartamentos.ResaltarLinea
    miCodDepartamento = CInt(moDepartamentos.TextoLinCol(moDepartamentos.LineaSeleccionada, 0))
End Sub
