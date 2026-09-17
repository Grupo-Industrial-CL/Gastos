VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfDetallado 
   Caption         =   "Informe Detallado"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   6045
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.ComboBox cmbDeducible 
      Height          =   315
      ItemData        =   "frmInfDetallado.frx":0000
      Left            =   4680
      List            =   "frmInfDetallado.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1920
      Width           =   1000
   End
   Begin VB.ComboBox cmbFamiliaGasto 
      Height          =   315
      ItemData        =   "frmInfDetallado.frx":0004
      Left            =   2280
      List            =   "frmInfDetallado.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   300
      Left            =   1680
      TabIndex        =   24
      Top             =   2480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   9502723
      CurrentDate     =   40415
   End
   Begin VB.OptionButton optPts 
      Height          =   255
      Left            =   3360
      Picture         =   "frmInfDetallado.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optEuro 
      Height          =   255
      Left            =   4320
      Picture         =   "frmInfDetallado.frx":02AB
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "C&errar"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   3000
      Picture         =   "frmInfDetallado.frx":057D
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3060
      Width           =   855
   End
   Begin VB.TextBox boxCodigo 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
   Begin VB.TextBox boxAnyo 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   1260
      Width           =   495
   End
   Begin VB.TextBox boxJerarquia 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   180
      Width           =   375
   End
   Begin VB.CommandButton btAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox boxCampanya 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   1260
      Width           =   375
   End
   Begin VB.ComboBox cboxTipoPago 
      Height          =   315
      ItemData        =   "frmInfDetallado.frx":0887
      Left            =   240
      List            =   "frmInfDetallado.frx":0897
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.OptionButton optEscudos 
      Height          =   255
      Left            =   3360
      Picture         =   "frmInfDetallado.frx":08BB
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optLiras 
      Height          =   255
      Left            =   3360
      Picture         =   "frmInfDetallado.frx":0B56
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblMatricula 
      BackColor       =   &H008080FF&
      Height          =   255
      Left            =   4560
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblDeducible 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deducible"
      Height          =   255
      Left            =   4680
      TabIndex        =   26
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia Gasto"
      Height          =   255
      Left            =   2280
      TabIndex        =   25
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fecha del informe"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   2520
      Width           =   1260
   End
   Begin VB.Label etCodigo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Código"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   180
      Width           =   615
   End
   Begin VB.Label etAnyo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Año"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label etNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   660
      Width           =   3135
   End
   Begin VB.Label etJerarquia 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jerarquía"
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   180
      Width           =   735
   End
   Begin VB.Label etCampanya 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Campaña"
      Height          =   255
      Left            =   1440
      TabIndex        =   18
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label etFecIni 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   1245
      Width           =   990
   End
   Begin VB.Label etFecFin 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4065
      TabIndex        =   16
      Top             =   1245
      Width           =   990
   End
   Begin VB.Label etTipoPago 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Pago"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   660
      Width           =   615
   End
   Begin VB.Label etAyuda 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1:DN  2:DR  3:DP  4:DA  5:SU"
      Height          =   465
      Left            =   3555
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1380
   End
End
Attribute VB_Name = "frmInfDetallado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub boxCodigo_GotFocus()
    etNombre = ""
    lblMatricula = ""
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
    sSql = "SELECT nombre,NVL(Matricula,'') Matricula FROM directores WHERE " & _
        "codpais=" & CodPaisTrabajo & " AND " & _
        "jerarquia='" & Jerarquia & "' AND " & _
        "codigo=" & boxCodigo '& " AND " & _
        "fechabaja IS NULL"
 Else
    sSql = "SELECT nombre,'' Matricula FROM supervisoras WHERE " & _
        "codpais=" & CodPaisTrabajo & " AND " & _
        "codsuperv=" & boxCodigo '& " AND " & _
        "fechabaja IS NULL"
 End If
 
 If gBdGALILEO.bRecordset(sSql, rsNombre) Then
 
    etNombre = rsNombre("nombre")
    lblMatricula = Trim(rsNombre("Matricula"))
        
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
    lblMatricula = ""
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
    sSql = "SELECT nombre,NVL(Matricula,'') Matricula FROM directores WHERE " & _
        "codpais=" & CodPaisTrabajo & " AND " & _
        "jerarquia='" & Jerarquia & "' AND " & _
        "codigo=" & boxCodigo '& " AND " & _
        "fechabaja IS NULL"
 Else
    sSql = "SELECT nombre,'' Matricula FROM supervisoras WHERE " & _
        "codpais=" & CodPaisTrabajo & " AND " & _
        "codsuperv=" & boxCodigo '& " AND " & _
        "fechabaja IS NULL"
 End If
 
 If gBdGALILEO.bRecordset(sSql, rsNombre) Then
     
    etNombre = rsNombre("nombre")
    lblMatricula = Trim(rsNombre("Matricula"))
        
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
Dim afoRegistro As ADODB.Recordset
Dim Persona As String
Dim Jerarquia As String
Dim Box
Dim TipoPago As String
Dim TotalGasto As Double
Dim Total As Double
Dim rsDatos  As ADODB.Recordset
Dim rsGastos As ADODB.Recordset
Dim sNombreGasto, sNombreGastoActual As String
Dim dImpGasto As Double
Dim dDescuentos As Double
Dim iNumGasto As Long
Dim Importe As Double
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

sSql = "SELECT t.nombregasto,g.tipogasto,g.fecha,g.donde,g.numero,g.importe as importe " & _
       "FROM gastos g,tiposgasto t WHERE " & _
        "codigo=" & boxCodigo & " AND " & _
        "jerarquia ='" & Jerarquia & "' AND " & _
        "codpais=" & CodPaisTrabajo & " AND " & _
        "campanya=" & boxCampanya & " AND " & _
        "anyo=" & boxAnyo & " AND " & _
        "tipopago='" & TipoPago & "' AND " & _
        "g.tipogasto=t.tipogasto AND t.Tipo = 'C' AND " & _
        "g.Validado = 'S' " & _
        "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
        "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
        "ORDER BY t.nombregasto,fecha"

 If gBdGALILEO.bRecordset(sSql, rsGastos) Then
 Else
    MsgBox "No existen datos para esa campaña", vbApplicationModal + vbInformation
    boxAnyo.SetFocus
    Exit Sub
 End If
 
Randomize
sTablaTemp = "TG" & Int((1000000 * Rnd) + 1)

 sSql = "CREATE temp table " & sTablaTemp & _
    "(contador serial," & _
    "nombregasto char(20)," & _
    "fecha date," & _
    "donde char(40)," & _
    "importe char(15)," & _
    "importetotal char(15)," & _
    "numerototal char(10)," & _
    "numero char(10))"
    
If Not gBdGALILEO.bExecute(sSql) Then
    MsgBox "Error al crear la tabla temporal", vbApplicationModal + vbExclamation
    Exit Sub
End If
 
  Total = 0
 
 While Not rsGastos.EOF
    dImpGasto = 0
    iNumGasto = 0
    
    sNombreGasto = rsGastos("nombregasto")
    sNombreGastoActual = rsGastos("nombregasto")
    
    sSql = "INSERT INTO " & sTablaTemp & " (nombregasto) VALUES " & _
        "('" & sNombreGasto & "')"
            
    If Not gBdGALILEO.bExecute(sSql) Then
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute sSql
        Stop
    End If
    
    While sNombreGasto = sNombreGastoActual And Not rsGastos.EOF
    
        If rsGastos("tipogasto") <> "K" Then
            Importe = rsGastos("importe")
        Else
            Importe = rsGastos("importe") * rsGastos("numero")
        End If
        
        If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
            If optEuro Then
                Importe = dConvierteMoneda(CStr(Importe), 0, CByte(CodPaisTrabajo))
            End If
        End If
                
        sSql = "INSERT INTO " & sTablaTemp & " (fecha,donde,importe,numero) VALUES " & _
            "('" & rsGastos("fecha") & "','" & rsGastos("donde") & "','" & FNumero(Importe, gDecimales) & "','" & rsGastos("numero") & "')"
        
        If Not gBdGALILEO.bExecute(sSql) Then Stop
        
        dImpGasto = Importe + dImpGasto
        iNumGasto = iNumGasto + rsGastos("numero")
        
        rsGastos.MoveNext
        If Not rsGastos.EOF Then sNombreGastoActual = rsGastos("nombregasto")
    Wend
    sSql = "INSERT INTO " & sTablaTemp & " (importetotal,numerototal) VALUES " & _
            "('" & FNumero(dImpGasto, gDecimales) & "','" & FNumero(iNumGasto) & "')"
    
    Total = dImpGasto + Total
     
     Select Case Trim(sNombreGasto)
            Case "Billetes"
               InfDetallado.Sections("PieInforme").Controls("etBillete").Caption = FNumero(dImpGasto, gDecimales)
            Case "Gasolina"
               InfDetallado.Sections("PieInforme").Controls("etGasolina").Caption = FNumero(dImpGasto, gDecimales)
            Case "Hotel"
               InfDetallado.Sections("PieInforme").Controls("etHotel").Caption = FNumero(dImpGasto, gDecimales)
            Case "Invitaciones"
               InfDetallado.Sections("PieInforme").Controls("etInvitaciones").Caption = FNumero(dImpGasto, gDecimales)
            Case "Kilometros"
               InfDetallado.Sections("PieInforme").Controls("etKm").Caption = FNumero(dImpGasto, gDecimales)
            Case "Peaje"
               InfDetallado.Sections("PieInforme").Controls("etPeaje").Caption = FNumero(dImpGasto, gDecimales)
            Case "Salones"
               InfDetallado.Sections("PieInforme").Controls("etSAlones").Caption = FNumero(dImpGasto, gDecimales)
            Case "Transporte"
               InfDetallado.Sections("PieInforme").Controls("etTransporte").Caption = FNumero(dImpGasto, gDecimales)
            Case "Varios"
               InfDetallado.Sections("PieInforme").Controls("etVarios").Caption = FNumero(dImpGasto, gDecimales)
            Case "Telefono"
               InfDetallado.Sections("PieInforme").Controls("etTelefono").Caption = FNumero(dImpGasto, gDecimales)
            Case "Cena"
               InfDetallado.Sections("PieInforme").Controls("etCena").Caption = FNumero(dImpGasto, gDecimales)
            Case "Comida"
               InfDetallado.Sections("PieInforme").Controls("etComida").Caption = FNumero(dImpGasto, gDecimales)
            Case "Cafeteria"
               InfDetallado.Sections("PieInforme").Controls("etCafeteria").Caption = FNumero(dImpGasto, gDecimales)
            Case "Parking"
               InfDetallado.Sections("PieInforme").Controls("etParking").Caption = FNumero(dImpGasto, gDecimales)
        End Select
    
    If Not gBdGALILEO.bExecute(sSql) Then
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute sSql
        Stop
    End If
 Wend
 
 InfDetallado.Sections("Cabezapagina").Controls("etNombre").Caption = etNombre
 'If CodPaisTrabajo = 34 And boxCodigo = "1" And boxJerarquia = "1" Then InfDetallado.Sections("Cabezapagina").Controls("etNombre").Caption = "JULIAN CARBALLO COBO"
 InfDetallado.Sections("Cabezapagina").Controls("etCodigo").Caption = boxCodigo
 InfDetallado.Sections("Cabezapagina").Controls("etTipoPago").Caption = cboxTipoPago & " (" & cmbFamiliaGasto & ")"
 InfDetallado.Sections("Cabezapagina").Controls("etTipoPersona").Caption = Persona
 InfDetallado.Sections("Cabezapagina").Controls("etFecIni").Caption = etFecIni
 InfDetallado.Sections("Cabezapagina").Controls("etFecFin").Caption = etFecFin
 InfDetallado.Sections("Cabezapagina").Controls("etCampanya").Caption = "C" & Format(boxCampanya, "00") & "-" & boxAnyo
 InfDetallado.Sections("Cabezapagina").Controls("etFecha").Caption = Format(dtpFecha, "dd/MM/yy")
 InfDetallado.Sections("Cabezapagina").Controls("etMatricula").Caption = ""
 If CodPaisTrabajo = giPAIS_PORTUGAL Then InfDetallado.Sections("Cabezapagina").Controls("etMatricula").Caption = "Matrícula: " & lblMatricula
 
 'cambios
 dDescuentos = 0
 If TipoPago = "M" Then
 
    sSql = "SELECT SUM(Importe) FROM Desc_Gastos " & _
          "WHERE Codigo = " & boxCodigo & _
          " AND Jerarquia = '" & Jerarquia & _
          "' AND Campanya = " & boxCampanya & _
          " AND Anyo = " & boxAnyo & _
          " AND CodPais = " & CodPaisTrabajo
    dDescuentos = 0
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
      If Not IsNull(afoRegistro(0)) Then dDescuentos = afoRegistro(0)
    End If
    
 End If
 InfDetallado.Sections("PieInforme").Controls("etDescuentos").Caption = "- " & FNumero(dDescuentos, gDecimales)
 'cambios
  
 Total = Total - dDescuentos
 InfDetallado.Sections("PieInforme").Controls("etTotal").Caption = FNumero(Total, gDecimales)

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
       InfDetallado.Sections("PieInforme").Controls("etEmpresa").Caption = rsDatos("Empresa")
    End If
    
    If Not IsNull(rsDatos("canal")) Then
       InfDetallado.Sections("PieInforme").Controls("etCanal").Caption = rsDatos("Canal")
    End If
    
    If Not IsNull(rsDatos("canal")) Then
       InfDetallado.Sections("PieInforme").Controls("etCc").Caption = rsDatos("Cc")
    End If
 End If
 
Set rsDatos = Nothing

sSql = "SELECT * from " & sTablaTemp & " ORDER BY contador"

If gBdGALILEO.bRecordset(sSql, rsDatos) Then
    InfDetallado.DataMember = ""
    Set InfDetallado.DataSource = rsDatos
    InfDetallado.Show vbModal
Else
    MsgBox "Error al visualizar el informe", vbApplicationModal + vbExclamation
End If

 
sSql = "DROP TABLE " & sTablaTemp
gBdGALILEO.bExecute sSql

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

Private Sub cmbFamiliaGasto_KeyPress(KeyAscii As Integer)
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
    dtpFecha = Format(gBdGALILEO.tFechaSistema, "dd/MM/yy")
    
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

