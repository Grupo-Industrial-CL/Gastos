VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInformeGastos 
   Caption         =   "Informe de Gastos"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDeducible 
      Height          =   315
      ItemData        =   "frmInformeGastos.frx":0000
      Left            =   4080
      List            =   "frmInformeGastos.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
      Width           =   1000
   End
   Begin VB.ComboBox cmbFamiliaGasto 
      Height          =   315
      ItemData        =   "frmInformeGastos.frx":0004
      Left            =   1680
      List            =   "frmInformeGastos.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.OptionButton optLiras 
      Height          =   255
      Left            =   240
      Picture         =   "frmInformeGastos.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optEscudos 
      Height          =   255
      Left            =   240
      Picture         =   "frmInformeGastos.frx":027A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optEuro 
      Height          =   255
      Left            =   240
      Picture         =   "frmInformeGastos.frx":0515
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optPts 
      Height          =   255
      Left            =   240
      Picture         =   "frmInformeGastos.frx":07E7
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "C&errar"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   3615
      Picture         =   "frmInformeGastos.frx":0A8A
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox boxCodigo 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox boxAnyo 
      Height          =   285
      Left            =   570
      TabIndex        =   2
      Top             =   1050
      Width           =   495
   End
   Begin VB.TextBox boxJerarquia 
      Height          =   285
      Left            =   2865
      TabIndex        =   1
      Top             =   135
      Width           =   375
   End
   Begin VB.CommandButton btAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2070
      TabIndex        =   10
      Top             =   2415
      Width           =   975
   End
   Begin VB.TextBox boxCampanya 
      Height          =   285
      Left            =   2085
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   300
      Left            =   120
      TabIndex        =   22
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   120258563
      CurrentDate     =   40415
   End
   Begin VB.Label lblMatricula 
      BackColor       =   &H008080FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblDeducible 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deducible"
      Height          =   255
      Left            =   4080
      TabIndex        =   24
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia Gasto"
      Height          =   255
      Left            =   1680
      TabIndex        =   23
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fecha del informe"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label etAyuda 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1:DN  2:DR  3:DP  4:DA  5:SU"
      Height          =   465
      Left            =   3450
      TabIndex        =   20
      Top             =   45
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label etNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   945
      TabIndex        =   19
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   255
      Left            =   105
      TabIndex        =   18
      Top             =   600
      Width           =   630
   End
   Begin VB.Label etCodigo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Código"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   615
   End
   Begin VB.Label etAnyo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Año"
      Height          =   255
      Left            =   90
      TabIndex        =   16
      Top             =   1065
      Width           =   375
   End
   Begin VB.Label etJerarquia 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jerarquía"
      Height          =   255
      Left            =   2025
      TabIndex        =   15
      Top             =   135
      Width           =   750
   End
   Begin VB.Label etCampanya 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Campaña"
      Height          =   255
      Left            =   1245
      TabIndex        =   14
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label etFecIni 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2715
      TabIndex        =   13
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label etFecFin 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3990
      TabIndex        =   12
      Top             =   1080
      Width           =   945
   End
End
Attribute VB_Name = "frmInformeGastos"
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

    Dim Persona As String
    Dim Jerarquia As String
    Dim Box
    Dim rsRegistro As ADODB.Recordset
    Dim sTablaTemp As String
    
    Dim lParteDiario As Long
    Dim afoRegAux As ADODB.Recordset
    Dim sFechaAnt As String
    Dim sAux As String
    
    On Error GoTo ErrorSub
    
    Set Box = Es_Vacio(boxCodigo, etNombre, boxJerarquia, boxAnyo, etFecIni, etFecFin, boxCampanya)

    If Not (Box Is Nothing) Then
        Box.SetFocus
        Exit Sub
    End If

    Set Box = Nothing

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

'sSql = "SHAPE {SELECT g.*,t.nombregasto FROM gastos g,tiposgasto t WHERE " & _
'        "codigo=" & boxCodigo & " AND " & _
'        "jerarquia = '" & Jerarquia & "' AND " & _
'        "codpais=" & CodPaisTrabajo & " AND " & _
'        "campanya=" & boxCampanya & " AND " & _
'        "anyo=" & boxAnyo & " AND " & _
'        "g.tipogasto=t.tipogasto AND t.Tipo = 'C' " & _
'        "ORDER BY fecha } as InformeGastos COMPUTE InformeGastos by 'fecha'"
'
' If EntornoDatos.Recordsets("PorFecha").State <> adStateClosed Then _
'    EntornoDatos.Recordsets("PorFecha").Close
' EntornoDatos.Commands("PorFecha").CommandText = sSql
' EntornoDatos.Commands("PorFecha").Execute

    Randomize
    sTablaTemp = "TG" & Int((1000000 * Rnd) + 1)

    sSql = "CREATE TEMP TABLE " & sTablaTemp & "(" & _
           "Fecha Date," & _
           "Importe Decimal(9,2)," & _
           "Numero Decimal(9,2)," & _
           "Donde Char(150)," & _
           "NombreGasto Char(20)," & _
           "Hora Char(2)," & _
           "DiaCam Char(3))"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al crear la tabla temporal.", vbApplicationModal + vbExclamation, "Error"
        Exit Sub
    End If

    sSql = "SELECT g.Importe,g.Numero,g.Donde,g.Fecha,t.NombreGasto,g.Hora " & _
           "FROM Gastos g,TiposGasto t " & _
           "WHERE Codigo = " & boxCodigo & _
           " AND Jerarquia = '" & Jerarquia & _
           "' AND CodPais = " & CodPaisTrabajo & _
           " AND Campanya = " & boxCampanya & _
           " AND Anyo = " & boxAnyo & _
           " AND g.TipoGasto = t.TipoGasto AND t.Tipo = 'C' AND " & _
           "Validado = 'S' " & _
           "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
           "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
           "GROUP BY Fecha,NombreGasto,Importe,Numero,Donde,Hora " & _
           "ORDER BY Fecha,Hora,NombreGasto,Importe"
    If Not gBdGALILEO.bRecordset(sSql, rsRegistro) Then
        MsgBox "No existen datos.", vbApplicationModal + vbInformation, "Información"
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute (sSql)
        Exit Sub
    End If
    
    sSql = "SELECT * FROM PartDiario_General " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " AND Campanya = " & boxCampanya & _
           " AND Anyo = " & boxAnyo & _
           " AND Codigo = " & boxCodigo & _
           " AND Jerarquia = " & boxJerarquia
    lParteDiario = -1
    If gBdGALILEO.bRecordset(sSql, afoRegAux) Then lParteDiario = afoRegAux("Id_ParteDiario")
    Set afoRegAux = Nothing
    
    sFechaAnt = ""
    sAux = ""
    While Not rsRegistro.EOF
    
        If sFechaAnt <> CStr(Format(rsRegistro("Fecha"), "ddMMyyyy")) Then
            
            sFechaAnt = CStr(Format(rsRegistro("Fecha"), "ddMMyyyy"))
            sAux = sDondeParteDiario(lParteDiario, rsRegistro("Fecha"), CodPaisTrabajo, boxAnyo, boxCampanya)
            
        End If
    
        sSql = "INSERT INTO " & sTablaTemp & "(Fecha,Numero,Donde,NombreGasto,Importe,Hora,DiaCam) " & _
               "VALUES ('" & rsRegistro("Fecha") & "'," & rsRegistro("Numero") & ",'" & _
               Trim(rsRegistro("Donde")) & sAux & "','" & Trim(rsRegistro("NombreGasto")) & "',"
        
        If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
            If optEuro Then
                sSql = sSql & dConvierteMoneda(CStr(rsRegistro("Importe")), 0, CByte(CodPaisTrabajo))
            Else
                sSql = sSql & rsRegistro("Importe")
            End If
        Else
            sSql = sSql & rsRegistro("Importe")
        End If
        sSql = sSql & ",'" & NoNull(rsRegistro("Hora"), "A") & "','" & sDiaCampanya(rsRegistro("Fecha")) & "')"
        
        If Not gBdGALILEO.bExecute(sSql) Then
            MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
            sSql = "DROP TABLE " & sTablaTemp
            gBdGALILEO.bExecute (sSql)
            Exit Sub
        End If
        
        rsRegistro.MoveNext
                    
    Wend
    
    Set rsRegistro = Nothing
    sSql = "SELECT * FROM " & sTablaTemp & " ORDER BY Fecha,Hora"
    
    If Not gBdGALILEO.bRecordset(sSql, rsRegistro) Then
        MsgBox "No existen datos.", vbApplicationModal + vbInformation, "Información"
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute (sSql)
        Exit Sub
    Else
        InformeGastos.DataMember = ""
        Set InformeGastos.DataSource = rsRegistro
    End If

    InformeGastos.Sections("CabezaPagina").Controls("etNombre").Caption = etNombre
    If CodPaisTrabajo = 34 And boxCodigo = "1" And boxJerarquia = "1" Then InformeGastos.Sections("CabezaPagina").Controls("etNombre").Caption = ""
    InformeGastos.Sections("CabezaPagina").Controls("etFecha").Caption = Format(dtpFecha.Value, "dd/MM/yy")
    InformeGastos.Sections("CabezaPagina").Controls("etCodigo").Caption = boxCodigo
    InformeGastos.Sections("CabezaPagina").Controls("etTipoPersona").Caption = Persona
    InformeGastos.Sections("CabezaPagina").Controls("etFecIni").Caption = etFecIni
    InformeGastos.Sections("CabezaPagina").Controls("etFecFin").Caption = etFecFin
    InformeGastos.Sections("CabezaPagina").Controls("etCampanya").Caption = "C" & Format(boxCampanya, "00") & "-" & boxAnyo
    InformeGastos.Sections("CabezaPagina").Controls("etFamiliaGasto").Caption = cmbFamiliaGasto.Text
    InformeGastos.Sections("CabezaPagina").Controls("etMatricula").Caption = ""
    If CodPaisTrabajo = giPAIS_PORTUGAL Then InformeGastos.Sections("CabezaPagina").Controls("etMatricula").Caption = "Matrícula: " & lblMatricula
    
    InformeGastos.Show vbModal
    
    Set rsRegistro = Nothing
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

Private Function sDondeParteDiario(lId_ParteDiario As Long, Fecha As Date, _
    iCodPais As Integer, iAnyo As Integer, yCampanya As Byte) As String
    
    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    Dim FechaInicio As Date
    Dim lSemanaCampanya As Long
    Dim lDiaCampanya As Long
    Dim lCuantosDiasCampanya
    Dim sAux As String
    
    Dim sFecIni As String
    Dim sFecFin As String
        
    On Error GoTo Error_Rutina

    If goUsuario.IdUsuario <> "jfsantos" Then
        sDondeParteDiario = ""
        Exit Function
    End If
    sDondeParteDiario = "(---)"
    
'    sSql = "SELECT UNIQUE FecInicio FROM Campanyas " & _
'           "WHERE CodPais = " & iCodPais & _
'           " AND Anyo = " & iAnyo & _
'           " AND CodCampanya = " & yCampanya
'    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
'        Set afoRegistro = Nothing
'        Exit Function
'    End If

    If (boxAnyo <> "") And (boxCampanya <> "") And (etNombre <> "") And (boxJerarquia <> "") Then
    
        If Not bDameFechasCampanya(CodPaisTrabajo, boxAnyo, boxCampanya, boxCodigo, boxJerarquia, sFecIni, sFecFin) Then
            Exit Function
        End If
        
    Else
        Exit Function
    End If

    FechaInicio = CDate(sFecIni)
    'FechaInicio = afoRegistro("FecInicio")
    'Set afoRegistro = Nothing

    While Weekday(FechaInicio) <> vbSunday
        FechaInicio = FechaInicio + 1
    Wend
 
    lSemanaCampanya = 1
    lDiaCampanya = 0
    lCuantosDiasCampanya = Fecha - FechaInicio
    If lCuantosDiasCampanya < 0 Then lCuantosDiasCampanya = 0
    lSemanaCampanya = (lCuantosDiasCampanya \ 7) + 1
    lDiaCampanya = (lCuantosDiasCampanya Mod 7)
 
    sAux = CStr(lSemanaCampanya) & CStr(lDiaCampanya)

    sSql = "SELECT DISTINCT Lugar FROM PartDiario_DetDia " & _
           "WHERE Id_ParteDiario = " & lId_ParteDiario & _
           " AND NOT Lugar IS NULL AND Dia_Campanya = " & sAux & _
           " UNION " & _
           "SELECT DISTINCT Lugar FROM PartDiario_DetDiaF " & _
           "WHERE Id_ParteDiario = " & lId_ParteDiario & _
           " AND NOT Lugar IS NULL AND Dia_Campanya = " & sAux
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Function
    End If
    sAux = "("
    While Not afoRegistro.EOF
        sAux = sAux & Trim(afoRegistro("Lugar")) & ","
        afoRegistro.MoveNext
    Wend
    sAux = Mid(sAux, 1, Len(sAux) - 1) & ")"
    Set afoRegistro = Nothing
    
    sDondeParteDiario = sAux
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Function

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

'    Dim rsCampanyas As ADODB.Recordset
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

'    If (boxAnyo <> "") And (boxCampanya <> "") Then
'
'        sSql = "SELECT fecinicio,fecfin FROM campanyas " & _
'               " WHERE codpais = " & CodPaisTrabajo & " AND " & _
'               " codcampanya = " & boxCampanya & " AND " & _
'               "anyo = " & boxAnyo
'
'        If gBdGALILEO.bRecordset(sSql, rsCampanyas) Then
'            etFecIni = rsCampanyas("fecinicio")
'            etFecFin = rsCampanyas("fecfin")
'        End If
'
'    End If
'
'    Set rsCampanyas = Nothing

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

'    Dim rsCampanyas As ADODB.Recordset
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

'    If (boxAnyo <> "") And (boxCampanya <> "") Then
'
'        sSql = "SELECT fecinicio,fecfin FROM campanyas " & _
'               " WHERE codpais = " & CodPaisTrabajo & " AND " & _
'               " codcampanya = " & boxCampanya & " AND " & _
'               "anyo = " & boxAnyo
'
'        If gBdGALILEO.bRecordset(sSql, rsCampanyas) Then
'            etFecIni = rsCampanyas("fecinicio")
'            etFecFin = rsCampanyas("fecfin")
'        End If
'
'    End If
'
'    Set rsCampanyas = Nothing

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

Private Function sDiaCampanya(sFecha As String) As String

    Dim fFecIni As Date
    Dim fFecFin As Date
    Dim fFecha As Date
    Dim bEnc As Boolean
    Dim iCont As Integer
    
    On Error GoTo Error_Rutina
    
    fFecha = sFecha
    fFecIni = etFecIni
    fFecFin = etFecFin
    
    fFecIni = fFecIni + 1
    bEnc = False
    For iCont = 1 To 15
    
        If fFecIni = fFecha Then
            bEnc = True
            Exit For
        End If
        
        fFecIni = fFecIni + 1
        If iCont = 5 Or iCont = 10 Then fFecIni = fFecIni + 2
        
    Next
    If Not bEnc Then
        sDiaCampanya = "-/-"
        Exit Function
    End If
    
    Select Case iCont
    Case 1
        sDiaCampanya = "1/1"
    Case 2
        sDiaCampanya = "1/2"
    Case 3
        sDiaCampanya = "1/3"
    Case 4
        sDiaCampanya = "1/4"
    Case 5
        sDiaCampanya = "1/5"
    Case 6
        sDiaCampanya = "2/1"
    Case 7
        sDiaCampanya = "2/2"
    Case 8
        sDiaCampanya = "2/3"
    Case 9
        sDiaCampanya = "2/4"
    Case 10
        sDiaCampanya = "2/5"
    Case 11
        sDiaCampanya = "3/1"
    Case 12
        sDiaCampanya = "3/2"
    Case 13
        sDiaCampanya = "3/3"
    Case 14
        sDiaCampanya = "3/4"
    Case 15
        sDiaCampanya = "3/5"
    Case Else
        sDiaCampanya = "-/-"
    End Select
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Function
