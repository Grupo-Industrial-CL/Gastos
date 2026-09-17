VERSION 5.00
Begin VB.Form frmSeleccionarPais 
   BorderStyle     =   0  'None
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
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
   ScaleHeight     =   2520
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbPaises 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   5160
      Picture         =   "frmSeleccionarPais.frx":0000
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAceptar 
      Height          =   480
      Left            =   4200
      Picture         =   "frmSeleccionarPais.frx":030A
      ToolTipText     =   "Aceptar"
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label lblPaisSeleccionado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ESPAÑA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "País/Empresa seleccionado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cambiar país/empresa de trabajo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   7
      Left            =   240
      TabIndex        =   2
      Top             =   1485
      Width           =   3255
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   5760
      Picture         =   "frmSeleccionarPais.frx":0BD4
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgIcono 
      Height          =   480
      Left            =   120
      Picture         =   "frmSeleccionarPais.frx":10E0
      Top             =   80
      Width           =   480
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "País/Empresa de trabajo"
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
      TabIndex        =   1
      Top             =   75
      Width           =   4440
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
      Width           =   6345
   End
End
Attribute VB_Name = "frmSeleccionarPais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mafoRegPaises As New ADODB.Recordset

Private Sub Form_Load()
    
    Dim sSql As String
    
    Dim vElemento As Variant

    On Error GoTo Error_Rutina
    
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
    
    lblPaisSeleccionado = sNombrePais(CodPaisTrabajo)

    sSql = ""
    For Each vElemento In goUsuario.PaisesAmbito
        sSql = sSql & "," & vElemento
    Next
    sSql = "0" & sSql

    sSql = "SELECT a.CodPais,a.Nombre,NVL(b.id,1) idEmpresa " & _
           "FROM Paises a LEFT JOIN Gastos_Empresas b ON  a.CodPais = b.CodPais " & _
           "WHERE a.CodPais IN (" & sSql & ") " & _
           "ORDER BY a.Nombre"
    If gBdGALILEO.bRecordset(sSql, mafoRegPaises) Then
         
        While Not mafoRegPaises.EOF
            cmbPaises.AddItem Trim(mafoRegPaises("Nombre"))
            cmbPaises.ItemData(cmbPaises.NewIndex) = mafoRegPaises("CodPais")
            mafoRegPaises.MoveNext
        Wend
        mafoRegPaises.MoveFirst
        
    End If
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub imgAceptar_Click()

    Dim bEnc As Boolean

    Me.MousePointer = DEFAULT
    
    bEnc = False
    If cmbPaises.Text = "" Then
        cmbPaises.SetFocus
        Exit Sub
    End If
    
    CodPaisTrabajo = cmbPaises.ItemData(cmbPaises.ListIndex)
    While Not mafoRegPaises.EOF And bEnc = False
        If CodPaisTrabajo = mafoRegPaises("CodPais") Then
            Empresa_CodPaisTrabajo = mafoRegPaises("idEmpresa")
            bEnc = True
        End If
        mafoRegPaises.MoveNext
    Wend
    
    Set mafoRegPaises = Nothing
    Unload Me

End Sub

Private Sub imgAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgSalir_Click()
    Set mafoRegPaises = Nothing
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub
