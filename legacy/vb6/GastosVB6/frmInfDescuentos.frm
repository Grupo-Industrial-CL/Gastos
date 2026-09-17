VERSION 5.00
Begin VB.Form frmInfDescuentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de descuentos"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPrincipal 
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   6000
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdVolver 
         Caption         =   "Volver"
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox boxCodigo 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox boxAnyo 
         Height          =   285
         Left            =   675
         TabIndex        =   2
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox boxJerarquia 
         Height          =   285
         Left            =   2955
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox boxCampanya 
         Height          =   285
         Left            =   2310
         TabIndex        =   3
         Top             =   960
         Width           =   375
      End
      Begin VB.Label etNombre 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   300
         Width           =   3135
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
      Begin VB.Label etCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   285
         Width           =   495
      End
      Begin VB.Label etJerarquia 
         AutoSize        =   -1  'True
         Caption         =   "Jerarquía"
         Height          =   195
         Left            =   2055
         TabIndex        =   12
         Top             =   300
         Width           =   675
      End
      Begin VB.Label etCampanya 
         AutoSize        =   -1  'True
         Caption         =   "Campaña"
         Height          =   195
         Left            =   1470
         TabIndex        =   11
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label etFecIni 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2940
         TabIndex        =   10
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label etFecFin 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4230
         TabIndex        =   9
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label etAnyo1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1020
         Width           =   285
      End
      Begin VB.Label etAyuda 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1:DN  2:DR  3:DP  4:DA  5:SU"
         Height          =   270
         Left            =   2070
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   2310
      End
   End
End
Attribute VB_Name = "frmInfDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()

    Dim Box
    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    Dim sJerarquia As String
    Dim dTotal As Double
    Dim sTablaTemp As String

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
    
    sSql = "SELECT * FROM Desc_Gastos " & _
           "WHERE Codigo = " & boxCodigo & _
           " AND Jerarquia = '" & sJerarquia & _
           "' AND CodPais = " & CodPaisTrabajo & _
           " AND Anyo = " & boxAnyo & _
           " AND Campanya = " & boxCampanya & _
           " ORDER BY Contador"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        MsgBox "No existen descuentos con los datos introducidos.", vbApplicationModal + vbExclamation, "Información"
        VaciarCampos
        Exit Sub
    End If
    
    Randomize
    sTablaTemp = "TD" & Int((1000000 * Rnd) + 1)

    sSql = "CREATE TEMP TABLE " & sTablaTemp & "(" & _
           "Concepto CHAR(255),Importe DECIMAL(9,2)," & _
           "Usuario CHAR(50))"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al crear la tabla temporal.", vbApplicationModal + vbExclamation, "Información"
        Set afoRegistro = Nothing
        Exit Sub
    End If
    
    dTotal = 0
    While Not afoRegistro.EOF
    
        sSql = "INSERT INTO " & sTablaTemp & " (Concepto,Importe,Usuario) " & _
               "VALUES ('" & Trim(afoRegistro("Concepto")) & "'," & _
               afoRegistro("Importe") & ",'" & Trim(afoRegistro("Usuario")) & "')"
        If Not gBdGALILEO.bExecute(sSql) Then
            Set afoRegistro = Nothing
            sSql = "DROP TABLE " & sTablaTemp
            gBdGALILEO.bExecute (sSql)
            MsgBox "Error al guardar datos.", vbApplicationModal + vbExclamation, "Información"
            Exit Sub
        End If
    
        dTotal = dTotal + afoRegistro("Importe")
        afoRegistro.MoveNext
        
    Wend
    Set afoRegistro = Nothing
    
    sSql = "SELECT * FROM " & sTablaTemp
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        MsgBox "Error al consultar resultados.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    
    dtrInfDescuentos.DataMember = ""
    Set dtrInfDescuentos.DataSource = afoRegistro
    dtrInfDescuentos.Sections("EncabezadoPagina").Controls("lblTitulo").Caption = "Descuentos en campaña " & boxCampanya & " - " & boxAnyo
    dtrInfDescuentos.Sections("EncabezadoPagina").Controls("lblNombre").Caption = Trim(etNombre)
    dtrInfDescuentos.Sections("PieInforme").Controls("lblTotal").Caption = dTotal
    dtrInfDescuentos.Show vbModal

    sSql = "DROP TABLE " & sTablaTemp
    gBdGALILEO.bExecute sSql
    Set afoRegistro = Nothing
        
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdVolver_Click()
    Unload Me
    frmInformes.Show vbModal
End Sub

Private Sub Form_Load()
    frmInformes.Hide
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
            VaciarCampos
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
            VaciarCampos
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

    If (boxAnyo <> "") And (boxCampanya <> "") And (etNombre <> "") And (boxJerarquia <> "") Then
    
        If bDameFechasCampanya(CodPaisTrabajo, boxAnyo, boxCampanya, boxCodigo, boxJerarquia, sFecIni, sFecFin) Then
            etFecIni = sFecIni
            etFecFin = sFecFin
        End If
        
    End If


Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub VaciarCampos()

    boxCodigo = ""
    boxJerarquia = ""
    boxAnyo = ""
    boxCampanya = ""
    etFecIni = ""
    etFecFin = ""
    etNombre = ""
    etJerarquia.Visible = False
    boxCodigo.SetFocus
    
End Sub

