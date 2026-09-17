VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInformeMensual 
   Caption         =   "Informe Mensual"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDeducible 
      Height          =   315
      ItemData        =   "frmInformeMensual.frx":0000
      Left            =   2760
      List            =   "frmInformeMensual.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2280
      Width           =   1000
   End
   Begin VB.ComboBox cmbFamiliaGasto 
      Height          =   315
      ItemData        =   "frmInformeMensual.frx":0004
      Left            =   240
      List            =   "frmInformeMensual.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.OptionButton optPts 
      Height          =   255
      Left            =   240
      Picture         =   "frmInformeMensual.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optEuro 
      Height          =   255
      Left            =   240
      Picture         =   "frmInformeMensual.frx":02AB
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optEscudos 
      Height          =   255
      Left            =   240
      Picture         =   "frmInformeMensual.frx":057D
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optLiras 
      Height          =   255
      Left            =   240
      Picture         =   "frmInformeMensual.frx":0818
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox boxHasta 
      Height          =   300
      Left            =   3090
      TabIndex        =   3
      Top             =   1200
      Width           =   1050
   End
   Begin VB.TextBox boxDesde 
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   1050
   End
   Begin VB.CommandButton btAceptar 
      Caption         =   "Aceptar"
      Height          =   450
      Left            =   1335
      TabIndex        =   11
      Top             =   3405
      Width           =   975
   End
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "C&errar"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   3840
      Picture         =   "frmInformeMensual.frx":0A8A
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3195
      Width           =   855
   End
   Begin VB.TextBox boxJerarquia 
      Height          =   285
      Left            =   3015
      TabIndex        =   1
      Top             =   105
      Width           =   375
   End
   Begin VB.TextBox boxCodigo 
      Height          =   285
      Left            =   975
      TabIndex        =   0
      Top             =   105
      Width           =   735
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   300
      Left            =   1680
      TabIndex        =   6
      Top             =   2880
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
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblDeducible 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deducible"
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia Gasto"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fecha del informe"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   2925
      Width           =   1260
   End
   Begin VB.Label etAyuda 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1:DN  2:DR  3:DP  4:DA  5:SU"
      Height          =   465
      Left            =   3525
      TabIndex        =   20
      Top             =   45
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label etFechaAyuda 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha en Formato ""ddmmaaaa"""
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1350
      TabIndex        =   19
      Top             =   1605
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.Label etHasta 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
      Height          =   300
      Left            =   2310
      TabIndex        =   18
      Top             =   1170
      Width           =   540
   End
   Begin VB.Label etDesde 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
      Height          =   300
      Left            =   240
      TabIndex        =   17
      Top             =   1185
      Width           =   585
   End
   Begin VB.Label etJerarquia 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jerarquía"
      Height          =   300
      Left            =   2100
      TabIndex        =   16
      Top             =   120
      Width           =   795
   End
   Begin VB.Label etCodigo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Código"
      Height          =   285
      Left            =   255
      TabIndex        =   15
      Top             =   105
      Width           =   615
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   300
      Left            =   240
      TabIndex        =   14
      Top             =   645
      Width           =   660
   End
   Begin VB.Label etNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   645
      Width           =   3135
   End
End
Attribute VB_Name = "frmInformeMensual"
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

Private Sub boxDesde_GotFocus()
    etFechaAyuda.Visible = True
    FechaGotFocus boxDesde
End Sub

Private Sub boxDesde_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, boxDesde, 8
End Sub

Private Sub boxDesde_LostFocus()
    FechaLostFocus boxDesde
    etFechaAyuda.Visible = False
End Sub

Private Sub boxHasta_GotFocus()
    etFechaAyuda.Visible = True
    FechaGotFocus boxHasta
End Sub

Private Sub boxHasta_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, boxHasta, 8
End Sub

Private Sub boxHasta_LostFocus()
    FechaLostFocus boxHasta
    etFechaAyuda.Visible = False
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

    Dim rsGastos As ADODB.Recordset
    Dim rsGastosTemp As ADODB.Recordset
    Dim rsCampanya As ADODB.Recordset
    Dim Donde, Importe, Numero, DiaRuta As String
    Dim Box
    Dim FechaDesde, FechaHasta, Jerarquia, Persona As String
    Dim Dias As Integer
    Dim FecInicio As Date
    Dim ImporteTotal As Double
    Dim sTablaTemp As String
    
    On Error GoTo ErrorSub
    
    Set Box = Es_Vacio(boxCodigo, etNombre, boxJerarquia, boxDesde, boxHasta)
    If Not (Box Is Nothing) Then
        Box.SetFocus
        Exit Sub
    End If

    Set Box = Nothing
    
    If cmbFamiliaGasto.Text = "" Then
        cmbFamiliaGasto.SetFocus
        Exit Sub
    End If

    Randomize
    sTablaTemp = "TG" & Int((1000000 * Rnd) + 1)

    sSql = "CREATE TEMP TABLE " & sTablaTemp & _
           "(Contador Serial, DiaRuta Char(5),Fecha Date," & _
           "G Decimal(9,2),DG Char(50),NG Decimal(9,2)," & _
           "I Decimal(9,2),DI Char(50),NI Decimal(9,2)," & _
           "H Decimal(9,2),DH Char(50),NH Decimal(9,2)," & _
           "S Decimal(9,2),DS Char(50),NS Decimal(9,2)," & _
           "B Decimal(9,2),DB Char(50),NB Decimal(9,2)," & _
           "F Decimal(9,2),DF Char(50),NF Decimal(9,2)," & _
           "T Decimal(9,2),DT Char(50),NT Decimal(9,2)," & _
           "P Decimal(9,2),DP Char(50),NP Decimal(9,2)," & _
           "V Decimal(9,2),DV Char(50),NV Decimal(9,2)," & _
           "K Decimal(9,2),DK Char(50),NK Decimal(9,2)," & _
           "O Decimal(9,2),DO Char(50),NO Decimal(9,2)," & _
           "E Decimal(9,2),DE Char(50),NE Decimal(9,2)," & _
           "X Decimal(9,2),DX Char(50),NX Decimal(9,2)," & _
           "A Decimal(9,2),DA Char(50),NA Decimal(9,2))"
    
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al crear la tabla temporal", vbApplicationModal + vbExclamation
        Exit Sub
    End If
 
    FechaDesde = Mid(boxDesde, 1, 2) + Mid(boxDesde, 4, 2) + Mid(boxDesde, 7, 4)

    FechaHasta = Mid(boxHasta, 1, 2) + Mid(boxHasta, 4, 2) + Mid(boxHasta, 7, 4)

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
        sSql = "DROP TABLE TempGastos"
        gBdGALILEO.bExecute sSql
        Exit Sub
    End Select
 
    sSql = "SELECT * FROM Gastos " & _
           "WHERE Codigo = " & boxCodigo & _
           " AND Jerarquia = '" & Jerarquia & _
           "' AND Codpais = " & CodPaisTrabajo & _
           " AND Fecha >= '" & FechaDesde & _
           "' AND Fecha <= '" & FechaHasta & "' AND " & _
           "Validado = 'S' " & _
           "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
           "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
           "ORDER BY Fecha,TipoGasto,Contador"
    
    If gBdGALILEO.bRecordset(sSql, rsGastos) Then
        
        Screen.MousePointer = 11
        ImporteTotal = 0
    
        While Not rsGastos.EOF
            
            Donde = "": Importe = "": Numero = ""
    
            If Not IsNull(rsGastos("Donde")) Then
                Donde = Trim(rsGastos("Donde"))
            End If
                
            If Not IsNull(rsGastos("Importe")) Then
            
                If (rsGastos("Importe")) <> 0 Then
                    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                        If optEuro Then
                            If rsGastos("TipoGasto") = "K" Then
                               Importe = dConvierteMoneda(CStr(rsGastos("Importe")), 0, CByte(CodPaisTrabajo)) * rsGastos("Numero")
                            Else
                               Importe = dConvierteMoneda(CStr(rsGastos("Importe")), 0, CByte(CodPaisTrabajo))
                            End If
                        Else
                            If rsGastos("TipoGasto") = "K" Then
                               Importe = rsGastos("Importe") * rsGastos("Numero")
                            Else
                               Importe = rsGastos("Importe")
                            End If
                        End If
                    Else
                        If rsGastos("TipoGasto") = "K" Then
                           Importe = rsGastos("Importe") * rsGastos("Numero")
                        Else
                           Importe = rsGastos("Importe")
                        End If
                    End If
                   
                   ImporteTotal = ImporteTotal + Importe
                   
                End If
            End If
                    
            If Not IsNull(rsGastos("numero")) Then
                If (rsGastos("Numero")) <> 0 Then Numero = rsGastos("Numero")
            End If
            
            sSql = "SELECT * FROM " & sTablaTemp & _
                   " WHERE Fecha = '" & rsGastos("Fecha") & "'"
            If Not gBdGALILEO.bRecordset(sSql, rsGastosTemp) Then
                         
                sSql = "SELECT Min(CodCampanya) As MinCampanya,FecInicio " & _
                       "FROM Campanyas " & _
                       "WHERE CodPais = " & CodPaisTrabajo & _
                       " AND Anyo = " & Year(rsGastos("Fecha")) & _
                       " AND '" & rsGastos("Fecha") & "' " & _
                       "BETWEEN FecInicio AND FecFin " & _
                       "GROUP BY FecInicio"
                If gBdGALILEO.bRecordset(sSql, rsCampanya) Then
                
                    FecInicio = rsCampanya("FecInicio")
                    
                    While Weekday(FecInicio) <> vbMonday
                        FecInicio = FecInicio + 1
                    Wend
                    
                    Dias = DateDiff("d", FecInicio, rsGastos("fecha"))
                    DiaRuta = (Int(Dias / 7)) + 1 & "/" & (Dias Mod 7) + 1
                    
                End If
                
                sSql = "INSERT INTO " & sTablaTemp & " (DiaRuta,Fecha," & rsGastos("TipoGasto") & _
                       ",D" & rsGastos("TipoGasto") & ",N" & rsGastos("TipoGasto") & ") " & _
                       "VALUES ('" & DiaRuta & " ','" & _
                        rsGastos("Fecha") & "','" & _
                        Importe & "',""" & _
                        Donde & """,'" & _
                        Numero & "')"
                
                If Not gBdGALILEO.bExecute(sSql) Then
                    sSql = "DROP TABLE " & sTablaTemp
                    gBdGALILEO.bExecute sSql
                    Stop
                End If
                
            Else
                
                sSql = "SELECT * FROM " & sTablaTemp & _
                       " WHERE Fecha = '" & rsGastos("Fecha") & _
                       "' AND " & rsGastos("TipoGasto") & " IS NULL " & _
                       "AND D" & rsGastos("TipoGasto") & " IS NULL " & _
                       "AND N" & rsGastos("TipoGasto") & " IS NULL " & _
                       "ORDER BY Contador"
                If gBdGALILEO.bRecordset(sSql, rsGastosTemp) Then
                    
                    sSql = "UPDATE " & sTablaTemp & " SET " & _
                           rsGastos("TipoGasto") & " = '" & Importe & "', " & _
                           "D" & rsGastos("TipoGasto") & " = """ & Donde & """, " & _
                           "N" & rsGastos("TipoGasto") & " = '" & Numero & "' " & _
                           "WHERE Contador = " & rsGastosTemp("Contador")
                    
                    If Not gBdGALILEO.bExecute(sSql) Then
                        sSql = "DROP TABLE " & sTablaTemp
                        gBdGALILEO.bExecute sSql
                        Stop
                    End If
                Else
                
                    sSql = "INSERT INTO " & sTablaTemp & " (" & rsGastos("TipoGasto") & ",D" & rsGastos("TipoGasto") & ",N" & rsGastos("TipoGasto") & ") " & _
                           "VALUES ('" & Importe & "',""" & Donde & """,'" & Numero & "')"
                    
                    If Not gBdGALILEO.bExecute(sSql) Then
                        sSql = "DROP TABLE " & sTablaTemp
                        gBdGALILEO.bExecute sSql
                        Stop
                    End If
                
                End If
                
            End If
            
            Set rsGastosTemp = Nothing
            rsGastos.MoveNext
        
        Wend
        
        Screen.MousePointer = 0
    
    Else
        MsgBox "No hay gastos entre esas fechas", vbApplicationModal + vbCritical
    End If

    Set rsGastos = Nothing


    sSql = "SELECT * FROM " & sTablaTemp & _
           " ORDER BY Contador"

    If gBdGALILEO.bRecordset(sSql, rsGastosTemp) Then
    
        InformeMensual.DataMember = ""
        Set InformeMensual.DataSource = rsGastosTemp
        InformeMensual.Orientation = rptOrientLandscape
        'InformeMensual.ReportWidth = 125
        InformeMensual.Sections("CabezaInforme").Controls("etNombre").Caption = etNombre
        If CodPaisTrabajo = 34 And boxCodigo = "1" And boxJerarquia = "1" Then InformeMensual.Sections("CabezaInforme").Controls("etNombre").Caption = "JULIAN CARBALLO COBO"
        InformeMensual.Sections("CabezaInforme").Controls("etFecha").Caption = Format(dtpFecha.Value, "dd/MM/yy")
        InformeMensual.Sections("CabezaInforme").Controls("etCodigo").Caption = boxCodigo
        InformeMensual.Sections("CabezaInforme").Controls("etTipoPersona").Caption = Persona
        InformeMensual.Sections("CabezaInforme").Controls("etFecIni").Caption = boxDesde
        InformeMensual.Sections("CabezaInforme").Controls("etFecFin").Caption = boxHasta
        InformeMensual.Sections("PieInforme").Controls("etimportetotal").Caption = FNumero(ImporteTotal, gDecimales)
        InformeMensual.Sections("CabezaInforme").Controls("etFamiliaGasto").Caption = cmbFamiliaGasto.Text
        InformeMensual.Sections("CabezaInforme").Controls("etMatricula").Caption = ""
        If CodPaisTrabajo = giPAIS_PORTUGAL Then InformeMensual.Sections("CabezaInforme").Controls("etMatricula").Caption = "Matrícula: " & lblMatricula
    
        InformeMensual.Show vbModal
        
    End If

    Set rsGastosTemp = Nothing
    
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

