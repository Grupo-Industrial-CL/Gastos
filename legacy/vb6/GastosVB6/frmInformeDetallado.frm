VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInformeDetallado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Detallado"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5670
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbDeducible 
      Height          =   315
      ItemData        =   "frmInformeDetallado.frx":0000
      Left            =   4440
      List            =   "frmInformeDetallado.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1920
      Width           =   1000
   End
   Begin VB.ComboBox cmbFamiliaGasto 
      Height          =   315
      ItemData        =   "frmInformeDetallado.frx":0004
      Left            =   2040
      List            =   "frmInformeDetallado.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.OptionButton optLiras 
      Height          =   255
      Left            =   3240
      Picture         =   "frmInformeDetallado.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2430
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optEscudos 
      Height          =   255
      Left            =   3240
      Picture         =   "frmInformeDetallado.frx":027A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2430
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optEuro 
      Height          =   255
      Left            =   4200
      Picture         =   "frmInformeDetallado.frx":0515
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2430
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optPts 
      Height          =   255
      Left            =   3240
      Picture         =   "frmInformeDetallado.frx":07E7
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2430
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboxTipoPago 
      Height          =   315
      ItemData        =   "frmInformeDetallado.frx":0A8A
      Left            =   120
      List            =   "frmInformeDetallado.frx":0A9A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox boxCampanya 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton btAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox boxJerarquia 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox boxAnyo 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox boxCodigo 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "C&errar"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   2880
      Picture         =   "frmInformeDetallado.frx":0ABE
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2880
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   300
      Left            =   1560
      TabIndex        =   24
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   121962499
      CurrentDate     =   40415
   End
   Begin VB.Label lblDeducible 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deducible"
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia Gasto"
      Height          =   255
      Left            =   2040
      TabIndex        =   25
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fecha del informe"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   2445
      Width           =   1260
   End
   Begin VB.Label etAyuda 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1:DN  2:DR  3:DP  4:DA  5:SU"
      Height          =   465
      Left            =   3435
      TabIndex        =   22
      Top             =   60
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   615
   End
   Begin VB.Label etTipoPago 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Pago"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label etFecFin 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3945
      TabIndex        =   19
      Top             =   1200
      Width           =   990
   End
   Begin VB.Label etFecIni 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   1185
      Width           =   990
   End
   Begin VB.Label etCampanya 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Campaña"
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label etJerarquia 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jerarquía"
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.Label etNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label etAnyo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Año"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label etCodigo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Código"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmInformeDetallado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    On Error GoTo ErrorSub
    
    TextLostFocus boxCodigo
    
    If boxCodigo = "" Or boxJerarquia = "" Then Exit Sub
 
    Select Case boxJerarquia
    Case "1"
        Jerarquia = "DN"
    Case "2"
        Jerarquia = "DR"
    Case "3"
        Jerarquia = "DP"
        If bEsCodigoDS(CodPaisTrabajo, boxCodigo) = True Then Jerarquia = "DS"
        'If (CodPaisTrabajo = 34 And boxCodigo = 979) Or (CodPaisTrabajo = 51 And boxCodigo = 459) Or (CodPaisTrabajo = 7 And boxCodigo = 19) Then Jerarquia = "DS"
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
    End If
 
ExitSub:
    Exit Sub
ErrorSub:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
            Resume Next
        End If
    End With
    GoTo ExitSub
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
    
    On Error GoTo ErrorSub
    
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
        If bEsCodigoDS(CodPaisTrabajo, boxCodigo) = True Then Jerarquia = "DS"
        'If (CodPaisTrabajo = 34 And boxCodigo = 979) Or (CodPaisTrabajo = 51 And boxCodigo = 459) Or (CodPaisTrabajo = 7 And boxCodigo = 19) Then Jerarquia = "DS"
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
    
    End If
 
ExitSub:
    Exit Sub
ErrorSub:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
            Resume Next
        End If
    End With
    GoTo ExitSub
End Sub

Private Sub btAceptar_Click()
    
    Dim rsTotalGastos As ADODB.Recordset
    Dim rsDatos As ADODB.Recordset
    Dim rsRegistro As ADODB.Recordset
    Dim Persona As String
    Dim Jerarquia As String
    Dim Box
    Dim TipoPago As String
    Dim Total As Double
    Dim Totalkm As Double
    Dim dImporte As Double
    Dim sTablaTemp As String

    On Error GoTo ErrorSub

    Set Box = Es_Vacio(boxCodigo, etNombre, boxJerarquia, boxAnyo, etFecIni, etFecFin, boxCampanya)
    If Not (Box Is Nothing) Then
        Box.SetFocus
        Exit Sub
    End If
    
    Set Box = Nothing

    Select Case cboxTipoPago.ListIndex
    Case "0"
        TipoPago = "M"
    Case "1"
        TipoPago = "V"
    Case "2"
        TipoPago = "S"
    Case "3"
        TipoPago = "X"
    End Select

    Select Case boxJerarquia
    Case "1"
        Jerarquia = "DN"
        Persona = "Director Nacional"
    Case "2"
        Jerarquia = "DR"
        Persona = "Director Regional"
    Case "3"
        Jerarquia = "DP"
        Persona = "Director Apoyo"
    Case "4"
        Jerarquia = "DA"
        Persona = "Director Area"
    Case "5"
        Jerarquia = "SU"
        Persona = "Supervisora"
    Case Else
        Exit Sub
    End Select
    
    If cmbFamiliaGasto.Text = "" Then
        cmbFamiliaGasto.SetFocus
        Exit Sub
    End If

    Randomize
    sTablaTemp = "TG" & Int((1000000 * Rnd) + 1)

    sSql = "CREATE TEMP TABLE " & sTablaTemp & "(" & _
           "CodPais SmallInt," & _
           "Codigo Integer," & _
           "Jerarquia Char(2)," & _
           "TipoGasto Char(1)," & _
           "Fecha Date," & _
           "Anyo SmallInt," & _
           "Campanya SmallInt," & _
           "Importe Decimal(9,2)," & _
           "Numero Decimal(9,2)," & _
           "Donde Char(50)," & _
           "TipoPago Char(1)," & _
           "NombreGasto Char(20))"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al crear la tabla temporal.", vbApplicationModal + vbExclamation, "Error"
        Exit Sub
    End If

'    sSql = "SHAPE {SELECT g.*,t.nombregasto FROM gastos g,tiposgasto t WHERE " & _
'           "codigo=" & boxCodigo & " AND " & _
'           "jerarquia = '" & Jerarquia & "' AND " & _
'           "codpais=" & CodPaisTrabajo & " AND " & _
'           "campanya=" & boxCampanya & " AND " & _
'           "anyo=" & boxAnyo & " AND " & _
'           "tipopago='" & TipoPago & "' AND t.Tipo = 'C' AND " & _
'           "g.tipogasto=t.tipogasto " & _
'           "ORDER BY fecha } as InformeDetallado COMPUTE InformeDetallado by 'nombregasto'"
'
' If EntornoDatos.Recordsets("PorTipoGasto").State <> adStateClosed Then _
'    EntornoDatos.Recordsets("PorTipoGasto").Close
' EntornoDatos.Commands("PorTipoGasto").CommandText = sSql
' EntornoDatos.Commands("PorTipoGasto").Execute

    sSql = "SELECT g.*,t.NombreGasto FROM Gastos g,TiposGasto t " & _
           "WHERE Codigo = " & boxCodigo & _
           " AND Jerarquia = '" & Jerarquia & _
           "' AND CodPais = " & CodPaisTrabajo & _
           " AND Campanya = " & boxCampanya & _
           " AND Anyo = " & boxAnyo & _
           " AND TipoPago = '" & TipoPago & _
           "' AND g.TipoGasto = t.TipoGasto AND t.Tipo = 'C' " & _
           "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
           "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
           "ORDER BY Fecha"
    
    If Not gBdGALILEO.bRecordset(sSql, rsRegistro) Then
        MsgBox "No existen datos.", vbApplicationModal + vbInformation, "Información"
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute (sSql)
        Exit Sub
    Else
    
        While Not rsRegistro.EOF
        
            sSql = "INSERT INTO " & sTablaTemp & "(CodPais,Codigo,Jerarquia,TipoGasto," & _
                   "Fecha,Anyo,Campanya,Numero,Donde,TipoPago,NombreGasto,Importe) " & _
                   "VALUES (" & rsRegistro("CodPais") & "," & rsRegistro("Codigo") & ",'" & _
                   rsRegistro("Jerarquia") & "','" & rsRegistro("TipoGasto") & "','" & _
                   rsRegistro("Fecha") & "'," & rsRegistro("Anyo") & "," & _
                   rsRegistro("Campanya") & "," & rsRegistro("Numero") & ",'" & _
                   rsRegistro("Donde") & "','" & rsRegistro("TipoPago") & "','" & _
                   rsRegistro("NombreGasto") & "',"
            
            If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    sSql = sSql & dConvierteMoneda(CStr(rsRegistro("Importe")), 0, CByte(CodPaisTrabajo)) & ")"
                Else
                    sSql = sSql & rsRegistro("Importe") & ")"
                End If
            Else
                sSql = sSql & rsRegistro("Importe") & ")"
            End If
            
            If Not gBdGALILEO.bExecute(sSql) Then
                MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                sSql = "DROP TABLE " & sTablaTemp
                gBdGALILEO.bExecute (sSql)
                Exit Sub
            End If
            
            rsRegistro.MoveNext
                        
        Wend
        
    End If
    
    Set rsRegistro = Nothing
    sSql = "SELECT * FROM " & sTablaTemp
    
    If Not gBdGALILEO.bRecordset(sSql, rsRegistro) Then
        MsgBox "No existen datos.", vbApplicationModal + vbInformation, "Información"
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute (sSql)
        Exit Sub
    Else
        InformeDetallado.DataMember = ""
        Set InformeDetallado.DataSource = rsRegistro
    End If

    InformeDetallado.Sections("CabezaPagina").Controls("etNombre").Caption = etNombre
    InformeDetallado.Sections("CabezaPagina").Controls("etFecha").Caption = Format(dtpFecha.Value, "dd/MM/yy")
    InformeDetallado.Sections("CabezaPagina").Controls("etCodigo").Caption = boxCodigo
    InformeDetallado.Sections("CabezaPagina").Controls("etTipoPago").Caption = cboxTipoPago & " (" & cmbFamiliaGasto & ")"
    InformeDetallado.Sections("CabezaPagina").Controls("etTipoPersona").Caption = Persona
    InformeDetallado.Sections("CabezaPagina").Controls("etFecIni").Caption = etFecIni
    InformeDetallado.Sections("CabezaPagina").Controls("etFecFin").Caption = etFecFin
    InformeDetallado.Sections("CabezaPagina").Controls("etCampanya").Caption = "C" & Format(boxCampanya, "00") & "-" & boxAnyo
 
    sSql = "SELECT  tipogasto,sum(importe) as total FROM gastos WHERE " & _
           "codigo=" & boxCodigo & " AND " & _
           "jerarquia = '" & Jerarquia & "' AND " & _
           "codpais=" & CodPaisTrabajo & " AND " & _
           "campanya=" & boxCampanya & " AND " & _
           "anyo=" & boxAnyo & " AND " & _
           "tipopago='" & TipoPago & "' " & _
           "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
           "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
           "GROUP BY tipogasto"
           
    Total = 0
    If gBdGALILEO.bRecordset(sSql, rsTotalGastos) Then
        
        While Not rsTotalGastos.EOF
            
            If rsTotalGastos("tipogasto") <> "K" Then
                dImporte = dConvierteMoneda(CStr(rsTotalGastos("Total")), 0, CByte(CodPaisTrabajo))
            End If
                
            'Total = Total + rsTotalGastos("total")
            Total = Total + dImporte
            Select Case rsTotalGastos("tipogasto")
            Case "B"
                InformeDetallado.Sections("PieInforme").Controls("etBillete").Caption = FNumero(dImporte, gDecimales)
            Case "G"
                InformeDetallado.Sections("PieInforme").Controls("etGasolina").Caption = FNumero(dImporte, gDecimales)
            Case "H"
                InformeDetallado.Sections("PieInforme").Controls("etHotel").Caption = FNumero(dImporte, gDecimales)
            Case "I"
                InformeDetallado.Sections("PieInforme").Controls("etInvitaciones").Caption = FNumero(dImporte, gDecimales)
            Case "P"
                InformeDetallado.Sections("PieInforme").Controls("etPeaje").Caption = FNumero(dImporte, gDecimales)
            Case "S"
                InformeDetallado.Sections("PieInforme").Controls("etSAlones").Caption = FNumero(dImporte, gDecimales)
            Case "T"
                InformeDetallado.Sections("PieInforme").Controls("etTransporte").Caption = FNumero(dImporte, gDecimales)
            Case "V"
                InformeDetallado.Sections("PieInforme").Controls("etVarios").Caption = FNumero(dImporte, gDecimales)
            Case "F"
                InformeDetallado.Sections("PieInforme").Controls("etTelefono").Caption = FNumero(dImporte, gDecimales)
            End Select
        
            rsTotalGastos.MoveNext
        
        Wend
        
    End If
 
    Set rsTotalGastos = Nothing
    sSql = "SELECT  sum(numero*importe) as total FROM gastos WHERE " & _
           "codigo=" & boxCodigo & " AND " & _
           "jerarquia = '" & Jerarquia & "' AND " & _
           "codpais=" & CodPaisTrabajo & " AND " & _
           "campanya=" & boxCampanya & " AND " & _
           "anyo=" & boxAnyo & " AND " & _
           "tipogasto='K' AND " & _
           "tipopago='" & TipoPago & "' " & _
           "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
           "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "'"

    Totalkm = 0
    
    If gBdGALILEO.bRecordset(sSql, rsTotalGastos) Then
        If Not IsNull(rsTotalGastos("total")) Then
            Totalkm = dConvierteMoneda(CStr(rsTotalGastos("total")), 0, CByte(CodPaisTrabajo))
            InformeDetallado.Sections("PieInforme").Controls("etKm").Caption = FNumero(Totalkm, gDecimales)
        End If
    End If

    InformeDetallado.Sections("PieInforme").Controls("etTotal").Caption = FNumero(Total + Totalkm, gDecimales)

    Set rsTotalGastos = Nothing
 
    If boxJerarquia <> "5" Then
       sSql = "SELECT empresa,canal,cc FROM directores WHERE " & _
              "codigo=" & boxCodigo & " AND " & _
              "jerarquia = '" & Jerarquia & "' AND " & _
              "codpais=" & CodPaisTrabajo
    Else
        sSql = "SELECT empresa,canal,cc FROM contratossup WHERE " & _
               "codsuperv=" & boxCodigo & " AND " & _
               "codpais=" & CodPaisTrabajo
    End If
 
    If gBdGALILEO.bRecordset(sSql, rsDatos) Then
    
        If Not IsNull(rsDatos("empresa")) Then
            InformeDetallado.Sections("PieInforme").Controls("etEmpresa").Caption = rsDatos("Empresa")
        End If
        
        If Not IsNull(rsDatos("canal")) Then
           InformeDetallado.Sections("PieInforme").Controls("etCanal").Caption = rsDatos("Canal")
        End If
        
        If Not IsNull(rsDatos("canal")) Then
           InformeDetallado.Sections("PieInforme").Controls("etCc").Caption = rsDatos("Cc")
        End If
    
    End If
    
    InformeDetallado.Show vbModal
 
    Set rsTotalGastos = Nothing
    Set rsDatos = Nothing
    Set rsRegistro = Nothing
    sSql = "DROP TABLE " & sTablaTemp
    gBdGALILEO.bExecute (sSql)
 
ExitSub:
    Exit Sub
ErrorSub:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
            Resume Next
        End If
    End With
    GoTo ExitSub
End Sub

Private Sub btCerrar_Click()
Unload Me
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
'Dim rsCampanyas As ADODB.Recordset

    Dim sFecIni As String
    Dim sFecFin As String

On Error GoTo ErrorSub

TextLostFocus boxCampanya

If (boxAnyo <> "") And (boxCampanya <> "") And (etNombre <> "") And (boxJerarquia <> "") Then

    If bDameFechasCampanya(CodPaisTrabajo, boxAnyo, boxCampanya, boxCodigo, boxJerarquia, sFecIni, sFecFin) Then
        etFecIni = sFecIni
        etFecFin = sFecFin
    End If
    
End If

'If (boxAnyo <> "") And (boxCampanya <> "") Then
'    sSql = "SELECT fecinicio,fecfin FROM campanyas " & _
'        " WHERE codpais = " & CodPaisTrabajo & " AND " & _
'        " codcampanya = " & boxCampanya & " AND " & _
'        "anyo = " & boxAnyo
'
'    If gBdGALILEO.bRecordset(sSql, rsCampanyas) Then
'      etFecIni = rsCampanyas("fecinicio")
'      etFecFin = rsCampanyas("fecfin")
'    End If
'End If

ExitSub:
    Exit Sub
ErrorSub:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
            Resume Next
        End If
    End With
    GoTo ExitSub
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

'Dim rsCampanyas As ADODB.Recordset

    Dim sFecIni As String
    Dim sFecFin As String

On Error GoTo ErrorSub

TextLostFocus boxAnyo

If (boxAnyo <> "") And (boxCampanya <> "") And (etNombre <> "") And (boxJerarquia <> "") Then

    If bDameFechasCampanya(CodPaisTrabajo, boxAnyo, boxCampanya, boxCodigo, boxJerarquia, sFecIni, sFecFin) Then
        etFecIni = sFecIni
        etFecFin = sFecFin
    End If
    
End If

'If (boxAnyo <> "") And (boxCampanya <> "") Then
'
'    sSql = "SELECT fecinicio,fecfin FROM campanyas " & _
'        " WHERE codpais = " & CodPaisTrabajo & " AND " & _
'        " codcampanya = " & boxCampanya & " AND " & _
'        "anyo = " & boxAnyo
'
'    If gBdGALILEO.bRecordset(sSql, rsCampanyas) Then
'        etFecIni = rsCampanyas("fecinicio")
'        etFecFin = rsCampanyas("fecfin")
'    End If
'End If

ExitSub:
    Exit Sub
ErrorSub:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
            Resume Next
        End If
    End With
    GoTo ExitSub
End Sub

Private Sub cboxTipoPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        DoEvents
    End If
End Sub

Private Sub cmbDeducible_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        DoEvents
    End If
End Sub

Private Sub Form_Load()
    
    cboxTipoPago.Text = "Metálico"
    
    Select Case CodPaisTrabajo
    Case 34
        optPts.Visible = True
        optEuro.Visible = True
    Case 51
        optEscudos.Visible = True
        optEuro.Visible = True
    Case 39
        optLiras.Visible = True
        optEuro.Visible = True
    End Select
    dtpFecha.Value = Format(gBdGALILEO.tFechaSistema, "dd/MM/yy")
    
    cmbFamiliaGasto.AddItem "Ordinario"
    cmbFamiliaGasto.AddItem "Formación"
    cmbFamiliaGasto.AddItem "Presentación"
    cmbFamiliaGasto.AddItem "Incentivo"
    cmbFamiliaGasto.Text = "Ordinario"
    
    cmbDeducible.AddItem "Sí"
    cmbDeducible.AddItem "No"
    cmbDeducible.ListIndex = 0
    lblDeducible.Visible = False
    cmbDeducible.Visible = False
    If CodPaisTrabajo = 52 Or CodPaisTrabajo = 56 Then
        lblDeducible.Visible = True
        cmbDeducible.Visible = True
    End If
    
End Sub

Private Sub optEscudos_Click()
    btAceptar.SetFocus
End Sub

Private Sub optEscudos_GotFocus()
    btAceptar.SetFocus
End Sub

Private Sub optEuro_Click()
    btAceptar.SetFocus
End Sub

Private Sub optEuro_GotFocus()
    btAceptar.SetFocus
End Sub

Private Sub optLiras_Click()
    btAceptar.SetFocus
End Sub

Private Sub optLiras_GotFocus()
    btAceptar.SetFocus
End Sub

Private Sub optPts_Click()
    btAceptar.SetFocus
End Sub

Private Sub optPts_GotFocus()
    btAceptar.SetFocus
End Sub

