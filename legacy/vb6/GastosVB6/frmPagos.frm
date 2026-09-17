VERSION 5.00
Begin VB.Form frmPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de gastos"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmPagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   8145
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtIdentificadorPago 
      Height          =   285
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   25
      Top             =   2840
      Width           =   1695
   End
   Begin VB.CommandButton cmdPagado 
      Caption         =   "Pagado"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5880
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame fraPrincipal 
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7815
      Begin VB.TextBox boxCodigo 
         Height          =   285
         Left            =   870
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
      Begin VB.CommandButton btOk 
         Caption         =   "OK"
         Height          =   330
         Left            =   5400
         TabIndex        =   4
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
      Begin VB.CommandButton cmdVolver 
         Caption         =   "Volver"
         Height          =   330
         Left            =   6840
         TabIndex        =   6
         Top             =   840
         Width           =   645
      End
      Begin VB.Label etNombre 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   300
         Width           =   3135
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   3675
         TabIndex        =   16
         Top             =   300
         Width           =   555
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
      Begin VB.Label etJerarquia 
         AutoSize        =   -1  'True
         Caption         =   "Jerarquía"
         Height          =   195
         Left            =   2055
         TabIndex        =   14
         Top             =   300
         Width           =   675
      End
      Begin VB.Label etCampanya 
         AutoSize        =   -1  'True
         Caption         =   "Campaña"
         Height          =   195
         Left            =   1470
         TabIndex        =   13
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label etFecIni 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2940
         TabIndex        =   12
         Top             =   980
         Width           =   945
      End
      Begin VB.Label etFecFin 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4230
         TabIndex        =   11
         Top             =   980
         Width           =   945
      End
      Begin VB.Label etAnyo1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1020
         Width           =   285
      End
      Begin VB.Label etAyuda 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1:DN  2:DR  3:DP  4:DA  5:SU"
         Height          =   270
         Left            =   2070
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   2310
      End
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Identificador de pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   720
      TabIndex        =   24
      Top             =   2880
      Width           =   1830
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblDescuentos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblGastos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Total a pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   20
      Top             =   2520
      Width           =   1155
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Importe total de descuentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   19
      Top             =   2160
      Width           =   2385
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Importe total de gastos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   18
      Top             =   1800
      Width           =   1965
   End
End
Attribute VB_Name = "frmPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btInicio_Click()
    
    boxCodigo.Enabled = True
    boxJerarquia.Enabled = True
    boxAnyo.Enabled = True
    boxCampanya.Enabled = True
    btOk.Enabled = True
    cmdPagado.Enabled = False
    txtIdentificadorPago.Enabled = False
    txtIdentificadorPago = ""
    
    Limpia_Boxes boxCodigo, boxJerarquia, etNombre, boxAnyo, boxCampanya, etFecIni, etFecFin, lblGastos, lblDescuentos, lblTotal
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
    
    sSql = "SELECT SUM(Importe) FROM Gastos " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " AND Codigo = " & boxCodigo & _
           " AND Jerarquia = '" & sJerarquia & _
           "' AND Anyo = " & boxAnyo & _
           " AND Campanya = " & boxCampanya & _
           " AND TipoGasto <> 'K'" & _
           " AND Pagado <> 'S' AND Validado = 'S'"
    dTotal = 0
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        If Not IsNull(afoRegistro(0)) Then dTotal = afoRegistro(0)
    End If
    Set afoRegistro = Nothing
    
    sSql = "SELECT SUM(Importe * Numero) FROM Gastos " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " AND Codigo = " & boxCodigo & _
           " AND Jerarquia = '" & sJerarquia & _
           "' AND Anyo = " & boxAnyo & _
           " AND Campanya = " & boxCampanya & _
           " AND TipoGasto = 'K'" & _
           " AND Pagado <> 'S' AND Validado = 'S'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        If Not IsNull(afoRegistro(0)) Then dTotal = dTotal + afoRegistro(0)
    End If
    Set afoRegistro = Nothing
    
    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then dTotal = dConvierteMoneda(CStr(dTotal), Local_A_Euros, CodPaisTrabajo)
    
    lblGastos = dTotal
    
    sSql = "SELECT SUM(Importe) FROM Desc_Gastos " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " AND Codigo = " & boxCodigo & _
           " AND Jerarquia = '" & sJerarquia & _
           "' AND Anyo = " & boxAnyo & _
           " AND Campanya = " & boxCampanya
    dTotal = 0
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
       If Not IsNull(afoRegistro(0)) Then dTotal = afoRegistro(0)
    End If
    Set afoRegistro = Nothing
    
    lblDescuentos = dTotal
    lblTotal = CDbl(lblGastos) - CDbl(lblDescuentos)
    
    btOk.Enabled = False
    cmdPagado.Enabled = True
    txtIdentificadorPago.Enabled = True
    If Val(lblGastos) = 0 Then
        cmdPagado.Enabled = False
        txtIdentificadorPago.Enabled = False
        lblTotal = 0
        lblDescuentos = 0
        txtIdentificadorPago = ""
    End If
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdPagado_Click()

    Dim sSql As String
    Dim sJerarquia As String
    
    On Error GoTo Error_Rutina
    
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
    txtIdentificadorPago = Trim(txtIdentificadorPago)
    If txtIdentificadorPago = "" Then
        txtIdentificadorPago.SetFocus
        Exit Sub
    End If
    
    sSql = "UPDATE Gastos SET Pagado = 'S',IdentificadorPago = '" & txtIdentificadorPago & _
           "' WHERE CodPais = " & CodPaisTrabajo & _
           " AND Codigo = " & boxCodigo & _
           " AND Jerarquia = '" & sJerarquia & _
           "' AND Anyo = " & boxAnyo & _
           " AND Campanya = " & boxCampanya & _
           " AND Validado = 'S'"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al marcar los gastos como pagados.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    
    btInicio_Click
        
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

Private Sub txtIdentificadorPago_GotFocus()
    TextGotFocus txtIdentificadorPago
End Sub

Private Sub txtIdentificadorPago_KeyPress(KeyAscii As Integer)

    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    
End Sub

Private Sub txtIdentificadorPago_LostFocus()
    TextLostFocus txtIdentificadorPago
End Sub
