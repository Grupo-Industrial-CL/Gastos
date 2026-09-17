VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Validación de Usuario"
   ClientHeight    =   1860
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3930
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1098.949
   ScaleMode       =   0  'User
   ScaleWidth      =   3690.057
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Picture         =   "Login.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1215
      Picture         =   "Login.frx":088C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtUserName 
      Height          =   270
      Left            =   1695
      MaxLength       =   12
      TabIndex        =   0
      Top             =   120
      Width           =   1365
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1695
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   1365
   End
   Begin VB.Label lblLabels 
      Caption         =   "Usuario:"
      Height          =   270
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   150
      Width           =   705
   End
   Begin VB.Label lblLabels 
      Caption         =   "Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   540
      Width           =   960
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moUsuario As CUsuario
Private miIdAplicacion As Integer
Private rcnPerfil As ADODB.Connection
Private gsUsuario As String
Private gsContraseña As String
Private iPrimeravez As Integer

Public Sub Create(iIdAplicacion As Integer, cnPerfil As ADODB.Connection)
    
    Set moUsuario = New CUsuario
    Set rcnPerfil = cnPerfil
    miIdAplicacion = iIdAplicacion
    
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    If txtUserName = "" Then
        txtUserName.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If txtPassword = "" Then
        txtPassword.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    moUsuario.Create txtUserName, txtPassword, miIdAplicacion, rcnPerfil
    
    If moUsuario.CreadoConExito Then
        Screen.MousePointer = vbDefault
        Unload Me
    Else
        txtUserName.SetFocus
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub cmdCancelar_Click()
'    Set moUsuario = Nothing
'    Unload Me
  End
End Sub

Private Sub Form_Activate()
  If gsUsuario <> "" And gsContraseña <> "" And iPrimeravez <> 1 Then
   txtUserName.Text = gsUsuario
   txtPassword.Text = gsContraseña
   iPrimeravez = 1
   cmdAceptar_Click
   
  End If
End Sub

Private Sub Form_Load()

    iPrimeravez = 0
     
'    Cargar_Textos
 
End Sub

'Private Sub Cargar_Textos()
'
'    Dim sIdentificadores As String
'    Dim sSql As String
'    Dim afoRegistro As ADODB.Recordset
'
'    On Error GoTo Error_Rutina
'
'    sIdentificadores = "(136,137,170,171,172)"
'
'    sSql = "SELECT IdPalabra," & gIdioma & " FROM Lenguaje " & _
'           "WHERE IdPalabra IN " & sIdentificadores & _
'           " ORDER BY IdPalabra"
'    If Not gBd.bRecordset(sSql, afoRegistro) Then
'        Set afoRegistro = Nothing
'        Exit Sub
'    End If
'
'    While Not afoRegistro.EOF
'        Select Case afoRegistro!idpalabra
'            Case 136
'                cmdAceptar.Caption = afoRegistro(1)
'            Case 137
'                cmdCancelar.Caption = afoRegistro(1)
'            Case 170
'                frmUserLogin.Caption = afoRegistro(1)
'            Case 171
'                lblLabels(0).Caption = afoRegistro(1)
'            Case 172
'                lblLabels(1).Caption = afoRegistro(1)
'    End Select
'
'        afoRegistro.MoveNext
'    Wend
'    Set afoRegistro = Nothing
'
'Exit_Rutina:
'    Exit Sub
'Error_Rutina:
'    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
'    Resume Exit_Rutina
'    Resume Next
'
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.BackColor = &H80FFFF
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

    cmdAceptar.DEFAULT = True
    
End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        cmdAceptar.SetFocus
    End If
    
End Sub

Private Sub txtPassword_LostFocus()
    
    txtPassword.BackColor = vbWhite
    cmdAceptar.DEFAULT = False
    
End Sub


Private Sub txtUserName_GotFocus()
    
    txtUserName.BackColor = &H80FFFF
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
    
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtPassword.SetFocus
    End If
    
End Sub


Private Sub txtUserName_LostFocus()
    txtUserName.BackColor = vbWhite
End Sub

Public Property Get Usuario() As CUsuario
    Set Usuario = moUsuario
End Property

Public Property Let Usuarios(sIdUsuario As String)
     gsUsuario = sIdUsuario
End Property

Public Property Let Contraseña(sIdContraseña As String)
     gsContraseña = sIdContraseña
End Property
