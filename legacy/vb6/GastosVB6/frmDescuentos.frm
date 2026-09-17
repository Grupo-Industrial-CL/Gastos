VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDescuentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos - Anexar descuentos"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   Icon            =   "frmDescuentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInsertar 
      Caption         =   "[^]"
      Enabled         =   0   'False
      Height          =   300
      Left            =   8880
      TabIndex        =   11
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   300
      Left            =   8880
      TabIndex        =   28
      Top             =   4060
      Width           =   375
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Top             =   4440
      Width           =   1000
   End
   Begin VB.TextBox txtConcepto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   255
      TabIndex        =   8
      Top             =   4440
      Width           =   4000
   End
   Begin VB.Frame fraPrincipal 
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   9135
      Begin VB.CommandButton cmdVolver 
         Caption         =   "Volver"
         Height          =   330
         Left            =   8280
         TabIndex        =   6
         Top             =   960
         Width           =   645
      End
      Begin VB.CommandButton btInicio 
         Caption         =   "Inicio"
         Height          =   330
         Left            =   8280
         TabIndex        =   5
         Top             =   600
         Width           =   645
      End
      Begin VB.CommandButton btOk 
         Caption         =   "OK"
         Height          =   330
         Left            =   8280
         TabIndex        =   4
         Top             =   240
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
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.Label etTotalCamp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6480
         TabIndex        =   22
         Top             =   980
         Width           =   1455
      End
      Begin VB.Label lblTotalCampanya 
         Caption         =   "    Total         Campaña"
         Height          =   435
         Left            =   5535
         TabIndex        =   21
         Top             =   915
         Width           =   825
      End
      Begin VB.Label etAnyo1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1020
         Width           =   285
      End
      Begin VB.Label etFecFin 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4230
         TabIndex        =   19
         Top             =   980
         Width           =   945
      End
      Begin VB.Label etFecIni 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2940
         TabIndex        =   18
         Top             =   980
         Width           =   945
      End
      Begin VB.Label etCampanya 
         AutoSize        =   -1  'True
         Caption         =   "Campaña"
         Height          =   195
         Left            =   1470
         TabIndex        =   17
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label etJerarquia 
         AutoSize        =   -1  'True
         Caption         =   "Jerarquía"
         Height          =   195
         Left            =   2055
         TabIndex        =   16
         Top             =   300
         Width           =   675
      End
      Begin VB.Label etCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   285
         Width           =   495
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   3675
         TabIndex        =   14
         Top             =   300
         Width           =   555
      End
      Begin VB.Label etNombre 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   300
         Width           =   3135
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfgParrilla 
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3836
      _Version        =   393216
      Enabled         =   0   'False
   End
   Begin VB.Label lblNomUsuario 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   10
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label lblUsuario 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuario"
      Height          =   255
      Left            =   5280
      TabIndex        =   27
      Top             =   4080
      Width           =   3500
   End
   Begin VB.Label lblImporte 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   4080
      Width           =   1000
   End
   Begin VB.Label lblConcepto 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   4000
   End
   Begin VB.Label lblContador 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moParrillaDescuentos As New CFlexGrid

Private Sub btInicio_Click()
    
    moParrillaDescuentos.Vaciar
    etTotalCamp = ""
    boxCodigo.Enabled = True
    boxJerarquia.Enabled = True
    boxAnyo.Enabled = True
    boxCampanya.Enabled = True
    btOk.Enabled = True

    DesactivarCampos

    Limpia_Boxes boxCodigo, boxJerarquia, etNombre, boxAnyo, boxCampanya, etFecIni, etFecFin
    Limpia_Boxes txtConcepto, txtImporte, lblNomUsuario
    boxCodigo.SetFocus
    
End Sub

Private Sub btOk_Click()

    Dim Box
    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    Dim sJerarquia As String

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
        sJerarquia = "DP','DS"
    Case "4"
        sJerarquia = "DA"
    Case "5"
        sJerarquia = "SU"
    Case Else
        Exit Sub
    End Select

    sSql = "SELECT * FROM Gastos " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " AND Codigo = " & boxCodigo & _
           " AND Jerarquia IN ('" & sJerarquia & _
           "') AND Anyo = " & boxAnyo & _
           " AND Campanya = " & boxCampanya
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        MsgBox "Aún no existen gastos con los datos introducidos.", vbApplicationModal + vbInformation, "Información"
        btInicio_Click
        boxCodigo.SetFocus
        Exit Sub
    End If
    
    ActivarCampos
    ActualizarParrilla
    
    boxCodigo.Enabled = False
    boxJerarquia.Enabled = False
    boxAnyo.Enabled = False
    boxCampanya.Enabled = False
    btOk.Enabled = False
    txtConcepto.SetFocus

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub ActivarCampos()

    mfgParrilla.Enabled = True
    cmdBorrar.Enabled = True
    cmdInsertar.Enabled = True
    txtConcepto.Enabled = True
    txtImporte.Enabled = True

End Sub

Private Sub DesactivarCampos()

    mfgParrilla.Enabled = False
    cmdBorrar.Enabled = False
    cmdInsertar.Enabled = False
    txtConcepto.Enabled = False
    txtImporte.Enabled = False

End Sub

Private Sub cmdBorrar_Click()

    Dim sSql As String
    
    On Error GoTo Error_Rutina

    If moParrillaDescuentos.LineaSeleccionada = -1 Then Exit Sub
    If lblContador = "" Then Exit Sub
    
    If MsgBox("¿Realmente desea elminar el descuento seleccionado?", vbApplicationModal + vbQuestion + vbYesNo, "Información") = vbNo Then Exit Sub
    
    sSql = "DELETE FROM Desc_Gastos " & _
           "WHERE Contador = " & lblContador
    gBdGALILEO.bExecute (sSql)
    
    ActualizarParrilla
    
    lblContador = ""
    txtConcepto = ""
    txtImporte = ""
    lblNomUsuario = ""
    txtConcepto.SetFocus

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdInsertar_Click()

    Dim sSql As String
    Dim sJerarquia As String

    On Error GoTo Error_Rutina

    If Trim(txtConcepto) = "" Then
        txtConcepto.SetFocus
        Exit Sub
    End If
    If Val(txtImporte) = 0 Then
        txtImporte.SetFocus
        Exit Sub
    End If
    
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
    
    If lblContador = "" Then
    'insertamos
    
        sSql = "INSERT INTO Desc_Gastos(CodPais,Jerarquia,Codigo,Anyo,Campanya,Importe,Usuario,Concepto) " & _
               "VALUES (" & CodPaisTrabajo & ",'" & sJerarquia & "'," & boxCodigo & _
               "," & boxAnyo & "," & boxCampanya & "," & txtImporte & ",'" & _
               Trim(goUsuario.Nombre) & "','" & Trim(txtConcepto) & "')"
    
    Else
    'updatear
    
        sSql = "UPDATE Desc_Gastos SET Concepto = '" & Trim(txtConcepto) & "'," & _
               "Importe = " & txtImporte & ",Usuario = '" & Trim(goUsuario.Nombre) & "' " & _
               "WHERE Contador = " & lblContador
               
    End If
    
    gBdGALILEO.bExecute (sSql)
    ActualizarParrilla
    
    lblContador = ""
    txtConcepto = ""
    txtImporte = ""
    lblNomUsuario = ""
    txtConcepto.SetFocus
    
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

Private Sub Form_Load()

    moParrillaDescuentos.Create mfgParrilla, 4, 2
    
    moParrillaDescuentos.Formatear 0, 0, "Contador"
    moParrillaDescuentos.Formatear 1, 4000, "Concepto", 4
    moParrillaDescuentos.Formatear 2, 1000, "Importe", 4, 7
    moParrillaDescuentos.Formatear 3, 3500, "Usuario", 4

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
 
    On Error GoTo Error_Rutina

    TextLostFocus boxCodigo
 
    If boxCodigo = "" Or boxJerarquia = "" Then Exit Sub

    Select Case boxJerarquia
    Case "1"
        Jerarquia = "DN"
    Case "2"
        Jerarquia = "DR"
    Case "3"
        Jerarquia = "DP','DS"
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
               "jerarquia IN ('" & Jerarquia & "') AND " & _
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
        Jerarquia = "DP','DS"
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
               "jerarquia IN ('" & Jerarquia & "') AND " & _
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

    On Error GoTo Error_Rutina

    If moParrillaDescuentos.bLineaActualVacia Then Exit Sub
    moParrillaDescuentos.ResaltarLinea
    
    lblContador = moParrillaDescuentos.TextoLinCol(moParrillaDescuentos.LineaSeleccionada, 0)
    txtConcepto = moParrillaDescuentos.TextoLinCol(moParrillaDescuentos.LineaSeleccionada, 1)
    txtImporte = moParrillaDescuentos.TextoLinCol(moParrillaDescuentos.LineaSeleccionada, 2)
    lblNomUsuario = moParrillaDescuentos.TextoLinCol(moParrillaDescuentos.LineaSeleccionada, 3)

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub txtConcepto_GotFocus()
    TextGotFocus txtConcepto
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtImporte.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtConcepto_LostFocus()
    TextLostFocus txtConcepto
End Sub

Private Sub txtImporte_GotFocus()
    TextGotFocus txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmdInsertar.SetFocus
        Exit Sub
    End If
    
    EntraNumero KeyAscii, txtImporte, 7, 2, "###0.00"

End Sub

Private Sub txtImporte_LostFocus()
    TextLostFocus txtImporte
End Sub

Private Sub ActualizarParrilla()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    Dim sJerarquia As String
    
    On Error GoTo Error_Rutina
    
    moParrillaDescuentos.Vaciar
    etTotalCamp = 0
    
    Select Case boxJerarquia
        Case "1"
            sJerarquia = "DN"
        Case "2"
            sJerarquia = "DR"
        Case "3"
            sJerarquia = "DP','DS"
        Case "4"
            sJerarquia = "DA"
        Case "5"
            sJerarquia = "SU"
        Case Else
            Exit Sub
     End Select
    
    sSql = "SELECT * FROM Desc_Gastos " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " AND Codigo = " & boxCodigo & _
           " AND Jerarquia IN ('" & sJerarquia & _
           "') AND Anyo = " & boxAnyo & _
           " AND Campanya = " & boxCampanya
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        MsgBox "No existen descuentos con los datos actuales.", vbApplicationModal + vbInformation, "Información"
        txtConcepto.SetFocus
        Exit Sub
    End If
    
    While Not afoRegistro.EOF
    
        etTotalCamp = CDbl(etTotalCamp) + afoRegistro("Importe")
        moParrillaDescuentos.TextoFila moParrillaDescuentos.LineaLibre, _
            afoRegistro("Contador"), Trim(afoRegistro("Concepto")), _
            afoRegistro("Importe"), Trim(afoRegistro("Usuario"))
        afoRegistro.MoveNext
        
    Wend
    Set afoRegistro = Nothing
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub
