VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEstadoGastos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos - Estado de los gastos"
   ClientHeight    =   5325
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
   Icon            =   "frmEstadoGastos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid mfgParrilla 
      Height          =   3495
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6165
      _Version        =   393216
   End
   Begin VB.Frame fraPrincipal 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton cmdVolver 
         Caption         =   "Volver"
         Height          =   330
         Left            =   6840
         TabIndex        =   6
         Top             =   840
         Width           =   645
      End
      Begin VB.CommandButton btInicio 
         Caption         =   "Inicio"
         Height          =   330
         Left            =   6120
         TabIndex        =   5
         Top             =   840
         Width           =   645
      End
      Begin VB.CommandButton btOk 
         Caption         =   "OK"
         Height          =   330
         Left            =   5400
         TabIndex        =   4
         Top             =   840
         Width           =   645
      End
      Begin VB.TextBox boxCampanya 
         Height          =   285
         Left            =   2310
         TabIndex        =   3
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox boxJerarquia 
         Height          =   285
         Left            =   2955
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox boxAnyo 
         Height          =   285
         Left            =   675
         TabIndex        =   2
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox boxCodigo 
         Height          =   285
         Left            =   870
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label etAyuda 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1:DN  2:DR  3:DP  4:DA  5:SU"
         Height          =   270
         Left            =   2070
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.Label etAnyo1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1020
         Width           =   285
      End
      Begin VB.Label etFecFin 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4230
         TabIndex        =   14
         Top             =   980
         Width           =   945
      End
      Begin VB.Label etFecIni 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2940
         TabIndex        =   13
         Top             =   980
         Width           =   945
      End
      Begin VB.Label etCampanya 
         AutoSize        =   -1  'True
         Caption         =   "Campaña"
         Height          =   195
         Left            =   1470
         TabIndex        =   12
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label etJerarquia 
         AutoSize        =   -1  'True
         Caption         =   "Jerarquía"
         Height          =   195
         Left            =   2055
         TabIndex        =   11
         Top             =   300
         Width           =   675
      End
      Begin VB.Label etCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   285
         Width           =   495
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   3675
         TabIndex        =   9
         Top             =   300
         Width           =   555
      End
      Begin VB.Label etNombre 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   300
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmEstadoGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moParrilla As New CFlexGrid

Private Sub Form_Load()
    
    On Error GoTo Error_Rutina

    moParrilla.Create mfgParrilla, 5, 2
    moParrilla.Formatear 0, 0, ""
    moParrilla.Formatear 1, 1500, "Identificador", 4, 4
    moParrilla.Formatear 2, 1000, "Validado", 4, 4
    moParrilla.Formatear 3, 1500, "Validación", 4
    moParrilla.Formatear 4, 1000, "Chequeado", 4, 4

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub btInicio_Click()
    
    moParrilla.Vaciar
    moParrilla.Grid.Enabled = False
    
    boxCodigo.Enabled = True
    boxJerarquia.Enabled = True
    boxAnyo.Enabled = True
    boxCampanya.Enabled = True
    btOk.Enabled = True
    
    Limpia_Boxes boxCodigo, boxJerarquia, etNombre, boxAnyo, boxCampanya, etFecIni, etFecFin
    boxCodigo.SetFocus
    
End Sub

Private Sub btOk_Click()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    Dim sJerarquia As String
    Dim Box
    Dim dTotal As Double
    
    On Error GoTo Error_Rutina
    
    Set Box = Es_Vacio(boxCodigo, etNombre, boxJerarquia, boxAnyo, etFecIni, etFecFin, boxCampanya)
    If Not (Box Is Nothing) Then
        Box.SetFocus
        Exit Sub
    End If
    Set Box = Nothing
    
    Select Case boxJerarquia
    Case "1"
        sJerarquia = "DN"
    Case "2"
        sJerarquia = "DR"
    Case "3"
        sJerarquia = "DP"
    Case "4"
        sJerarquia = "DA"
    Case "5"
        sJerarquia = "SU"
    Case Else
        Exit Sub
    End Select
    
    sSql = "SELECT Identificador,Validado,Validacion,Chequeado " & _
           "FROM Gastos " & _
           "WHERE Anyo = " & boxAnyo & _
           " AND Campanya = " & boxCampanya & _
           " AND Codigo = " & boxCodigo & _
           " AND Jerarquia = '" & sJerarquia & _
           "' AND CodPais = " & CodPaisTrabajo
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Sub
    End If
    
    While Not afoRegistro.EOF
    
        moParrilla.TextoFila moParrilla.LineaLibre, "", afoRegistro("Identificador"), _
            Trim(afoRegistro("Validado")), Trim(afoRegistro("Validacion")), _
            Trim(afoRegistro("Chequeado"))
        afoRegistro.MoveNext
    
    Wend
    moParrilla.SituarEn 0, 1
    Set afoRegistro = Nothing
    
    moParrilla.Grid.Enabled = True
    boxCodigo.Enabled = False
    boxJerarquia.Enabled = False
    boxAnyo.Enabled = False
    boxCampanya.Enabled = False
    btOk.Enabled = False
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdVolver_Click()
    Unload Me
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
 
    On Error GoTo Error_Rutina

    TextLostFocus boxCodigo
 
    If boxCodigo = "" Or boxJerarquia = "" Then Exit Sub

    Select Case boxJerarquia
    Case "1"
        Jerarquia = "DN"
    Case "2"
        Jerarquia = "DR"
    Case "3"
        Jerarquia = "DP"
    Case "4"
        Jerarquia = "DA"
    Case "5"
        Jerarquia = "SU"
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
    Else
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

    On Error GoTo Error_Rutina

    etAyuda.Visible = False
    TextLostFocus boxJerarquia

    If boxCodigo = "" Or boxJerarquia = "" Then Exit Sub

    Select Case boxJerarquia
    Case "1"
        Jerarquia = "DN"
    Case "2"
        Jerarquia = "DR"
    Case "3"
        Jerarquia = "DP"
    Case "4"
        Jerarquia = "DA"
    Case "5"
        Jerarquia = "SU"
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
    Else
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

Private Sub boxAnyo_GotFocus()
    etFecIni = ""
    etFecFin = ""
    TextGotFocus boxAnyo
End Sub

Private Sub boxAnyo_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, boxAnyo, 4
End Sub

Private Sub boxAnyo_LostFocus()

'    Dim rsCampanyas As ADODB.Recordset
    Dim sFecIni As String
    Dim sFecFin As String

    On Error GoTo Error_Rutina

    TextLostFocus boxAnyo
    
    If boxAnyo <> "" Then
        If CInt(boxAnyo) > Year(Date) Then
            boxAnyo = ""
            boxAnyo.SetFocus
            Exit Sub
        End If
    End If
    
    If (boxAnyo <> "") And (boxCampanya <> "") And (etNombre <> "") And (boxJerarquia <> "") Then
    
        If bDameFechasCampanya(CodPaisTrabajo, boxAnyo, boxCampanya, boxCodigo, boxJerarquia, sFecIni, sFecFin) Then
            etFecIni = sFecIni
            etFecFin = sFecFin
        End If
        
    End If

    
'    If (boxAnyo <> "") And (boxCampanya <> "") Then
'
'        sSql = "SELECT fecinicio,fecfin FROM campanyas " & _
'               " WHERE codpais = " & CodPaisTrabajo & " AND " & _
'               " codcampanya = " & boxCampanya & " AND " & _
'               "anyo = " & boxAnyo '& " AND FecFin <= '" & _
'               Format(Date, "ddmmyyyy") & "'"
'
'        If gBdGALILEO.bRecordset(sSql, rsCampanyas) Then
'            etFecIni = rsCampanyas("fecinicio")
'            etFecFin = rsCampanyas("fecfin")
'        End If
'
'    End If
'
'    Set rsCampanyas = Nothing

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub boxCampanya_GotFocus()
    etFecIni = ""
    etFecFin = ""
    TextGotFocus boxCampanya
End Sub

Private Sub boxCampanya_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, boxCampanya, 2
End Sub

Private Sub boxCampanya_LostFocus()
    
'    Dim rsCampanyas As ADODB.Recordset
    Dim sFecIni As String
    Dim sFecFin As String

    On Error GoTo Error_Rutina


    TextLostFocus boxCampanya
    
    If (boxAnyo <> "") And (boxCampanya <> "") And (etNombre <> "") And (boxJerarquia <> "") Then
    
        If bDameFechasCampanya(CodPaisTrabajo, boxAnyo, boxCampanya, boxCodigo, boxJerarquia, sFecIni, sFecFin) Then
            etFecIni = sFecIni
            etFecFin = sFecFin
        End If
        
    End If

'    If (boxAnyo <> "") And (boxCampanya <> "") Then
'
'        sSql = "SELECT fecinicio,fecfin FROM campanyas " & _
'               " WHERE codpais = " & CodPaisTrabajo & " AND " & _
'               " codcampanya = " & boxCampanya & " AND " & _
'               "anyo = " & boxAnyo '& " AND FecFin <= '" & _
'               Format(Date, "ddmmyyyy") & "'"
'
'        If gBdGALILEO.bRecordset(sSql, rsCampanyas) Then
'            etFecIni = rsCampanyas("fecinicio")
'            etFecFin = rsCampanyas("fecfin")
'        End If
'
'    End If
'
'    Set rsCampanyas = Nothing

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub mfgParrilla_Click()

    If moParrilla.bLineaActualVacia Then
        moParrilla.SituarEn 0, 1
        Exit Sub
    End If

    moParrilla.ResaltarLinea
    
End Sub
