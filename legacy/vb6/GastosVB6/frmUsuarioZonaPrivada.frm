VERSION 5.00
Begin VB.Form frmUsuarioZonaPrivada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos - Consulta de Usuarios de la Zona Privada"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsuarioZonaPrivada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPrincipal 
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7815
      Begin VB.TextBox boxCodigo 
         Height          =   285
         Left            =   870
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox boxJerarquia 
         Height          =   285
         Left            =   2955
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton btInicio 
         Caption         =   "Inicio"
         Height          =   330
         Left            =   6120
         TabIndex        =   2
         Top             =   840
         Width           =   645
      End
      Begin VB.CommandButton cmdVolver 
         Caption         =   "Volver"
         Height          =   330
         Left            =   6840
         TabIndex        =   3
         Top             =   840
         Width           =   645
      End
      Begin VB.Label lblClave 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblUsuario 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Clave"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   11
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   10
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label etNombre 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   300
         Width           =   3135
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   3675
         TabIndex        =   8
         Top             =   300
         Width           =   555
      End
      Begin VB.Label etCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   285
         Width           =   495
      End
      Begin VB.Label etJerarquia 
         AutoSize        =   -1  'True
         Caption         =   "Jerarquía"
         Height          =   195
         Left            =   2055
         TabIndex        =   6
         Top             =   300
         Width           =   675
      End
      Begin VB.Label etAyuda 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1:DN  2:DR  3:DP  4:DA  5:SU"
         Height          =   270
         Left            =   2070
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   2310
      End
   End
End
Attribute VB_Name = "frmUsuarioZonaPrivada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVolver_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub boxCodigo_GotFocus()
    etNombre = ""
    TextGotFocus boxCodigo
End Sub

Private Sub boxCodigo_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, boxCodigo, 6
End Sub

Private Sub boxCodigo_LostFocus()
    
    Dim rsNombre As ADODB.Recordset
    Dim Jerarquia As String
    Dim iJerarquia As Integer
 
    On Error GoTo Error_Rutina

    TextLostFocus boxCodigo
 
    If boxCodigo = "" Or boxJerarquia = "" Then Exit Sub

    Select Case boxJerarquia
    Case "1"
        Jerarquia = "DN"
        iJerarquia = 6
    Case "2"
        Jerarquia = "DR"
        iJerarquia = 5
    Case "3"
        Jerarquia = "DP"
        iJerarquia = 7
    Case "4"
        Jerarquia = "DA"
        iJerarquia = 4
    Case "5"
        Jerarquia = "SU"
        iJerarquia = 3
    Case Else
        Exit Sub
    End Select

    If boxJerarquia <> 5 Then
        sSql = "SELECT nombre FROM directores WHERE " & _
               "codpais=" & CodPaisTrabajo & " AND " & _
               "jerarquia='" & Jerarquia & "' AND " & _
               "codigo=" & boxCodigo '& " AND " & _
               "fechabaja IS NULL"
    Else
        sSql = "SELECT nombre FROM supervisoras WHERE " & _
               "codpais=" & CodPaisTrabajo & " AND " & _
               "codsuperv=" & boxCodigo '& " AND " & _
               "fechabaja IS NULL"
    End If

    If gBdGALILEO.bRecordset(sSql, rsNombre) Then
        
        etNombre = rsNombre("nombre")
        Set rsNombre = Nothing
        sSql = "SELECT * FROM Usuarios_Internet " & _
               "WHERE CodPais = " & CodPaisTrabajo & _
               " AND Jerarquia = " & boxJerarquia & _
               " AND Codigo = " & boxCodigo
        If Not gBdGALILEO.bRecordset(sSql, rsNombre) Then
            lblUsuario = ""
            lblClave = ""
        Else
            lblUsuario = Trim(rsNombre("Usuario"))
            lblClave = Trim(rsNombre("Passwd"))
        End If
        Set rsNombre = Nothing
        
    Else
        Set rsNombre = Nothing
        If etNombre = "" Then
            MsgBox "No existe el código introducido o está dado de baja.", vbApplicationModal + vbInformation, "Información"
            btInicio_Click
            Exit Sub
        End If
    End If
 
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub boxJerarquia_GotFocus()
    etNombre = ""
    TextGotFocus boxJerarquia
    etAyuda.Visible = True
End Sub

Private Sub boxJerarquia_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, boxJerarquia, 1
End Sub

Private Sub boxJerarquia_LostFocus()
    
    Dim rsNombre As ADODB.Recordset
    Dim Jerarquia As String
    Dim iJerarquia As Integer

    On Error GoTo Error_Rutina

    etAyuda.Visible = False
    TextLostFocus boxJerarquia

    If boxCodigo = "" Or boxJerarquia = "" Then Exit Sub

    Select Case boxJerarquia
    Case "1"
        Jerarquia = "DN"
        iJerarquia = 6
    Case "2"
        Jerarquia = "DR"
        iJerarquia = 5
    Case "3"
        Jerarquia = "DP"
        iJerarquia = 7
    Case "4"
        Jerarquia = "DA"
        iJerarquia = 4
    Case "5"
        Jerarquia = "SU"
        iJerarquia = 3
    Case Else
        Exit Sub
    End Select

    If boxJerarquia <> 5 Then
        sSql = "SELECT nombre FROM directores WHERE " & _
               "codpais=" & CodPaisTrabajo & " AND " & _
               "jerarquia='" & Jerarquia & "' AND " & _
               "codigo=" & boxCodigo '& " AND " & _
               "fechabaja IS NULL"
    Else
        sSql = "SELECT nombre FROM supervisoras WHERE " & _
               "codpais=" & CodPaisTrabajo & " AND " & _
               "codsuperv=" & boxCodigo '& " AND " & _
               "fechabaja IS NULL"
    End If

    If gBdGALILEO.bRecordset(sSql, rsNombre) Then
    
        etNombre = rsNombre("nombre")
        Set rsNombre = Nothing
        sSql = "SELECT * FROM Usuarios_Internet " & _
               "WHERE CodPais = " & CodPaisTrabajo & _
               " AND Jerarquia = " & iJerarquia & _
               " AND Codigo = " & boxCodigo
        If Not gBdGALILEO.bRecordset(sSql, rsNombre) Then
            lblUsuario = ""
            lblClave = ""
        Else
            lblUsuario = Trim(rsNombre("Usuario"))
            lblClave = Trim(rsNombre("Passwd"))
        End If
        Set rsNombre = Nothing

    Else
        Set rsNombre = Nothing
        If etNombre = "" Then
            MsgBox "No existe el código introducido o está dado de baja.", vbApplicationModal + vbInformation, "Información"
            btInicio_Click
            Exit Sub
        End If
    End If

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub btInicio_Click()
    
    Limpia_Boxes boxCodigo, boxJerarquia, etNombre, lblUsuario, lblClave
    boxCodigo.SetFocus
    
End Sub

