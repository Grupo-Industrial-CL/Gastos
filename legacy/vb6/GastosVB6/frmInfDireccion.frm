VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmInfDireccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfDireccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar3 
      Caption         =   "Aceptar Región"
      Height          =   735
      Left            =   5400
      Picture         =   "frmInfDireccion.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin Crystal.CrystalReport crpInforme 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdAceptar2 
      Caption         =   "Aceptar Área"
      Height          =   735
      Left            =   3960
      Picture         =   "frmInfDireccion.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtCampanya 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1160
      Width           =   375
   End
   Begin VB.TextBox txtAnyo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1160
      Width           =   735
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Aceptar"
      Height          =   735
      Left            =   0
      Picture         =   "frmInfDireccion.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbDirArea 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   680
      Width           =   5175
   End
   Begin VB.ComboBox cmbRegion 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   320
      Width           =   5175
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   735
      Left            =   6960
      Picture         =   "frmInfDireccion.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Campaña"
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   10
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Año"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione director de área"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   2355
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione región"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   1515
   End
End
Attribute VB_Name = "frmInfDireccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbDirArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAnyo.SetFocus
End Sub

Private Sub cmbRegion_Click()
    CargarDirArea
End Sub

Private Sub cmbRegion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbDirArea.SetFocus
End Sub

Private Sub cmdAceptar_Click()
    
    Dim sPersona As String
    Dim sJerarquia As String
    Dim afoRegistro As ADODB.Recordset
    Dim sTablaTemp As String
    Dim sSql As String
    
    Dim iCodigo As Integer
    Dim sFecIni As String
    Dim sFecFin As String
     
    On Error GoTo Error_Rutina
    
    If cmbRegion.Text = "" Then
        cmbRegion.SetFocus
        Exit Sub
    End If
    If cmbDirArea.Text = "" Then
        cmbDirArea.SetFocus
        Exit Sub
    End If
    txtAnyo = Trim(txtAnyo)
    If txtAnyo = "" Then
        txtAnyo.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtAnyo) Then
        txtAnyo = ""
        txtAnyo.SetFocus
        Exit Sub
    End If
    txtCampanya = Trim(txtCampanya)
    If txtCampanya = "" Then
        txtCampanya.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtCampanya) Then
        txtCampanya = ""
        txtCampanya.SetFocus
        Exit Sub
    End If
'    If txtCampanya < 1 Or txtCampanya > 17 Then
'        txtCampanya = ""
'        txtCampanya.SetFocus
'        Exit Sub
'    End If
    
'    sSql = "SELECT FecInicio,FecFin FROM Campanyas " & _
'           "WHERE CodPais = " & CodPaisTrabajo & _
'           " AND Anyo = " & txtAnyo & _
'           " AND CodCampanya = " & txtCampanya
'    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
'        Set afoRegistro = Nothing
'        txtAnyo.SetFocus
'        Exit Sub
'    End If
'    If CLng(Format(afoRegistro("FecInicio"), "yyyyMMdd")) > CLng(Format(Now, "yyyyMMdd")) Then
'        Set afoRegistro = Nothing
'        txtAnyo = ""
'        txtCampanya = ""
'        txtAnyo.SetFocus
'        Exit Sub
'    End If
'    sFecIni = Format(afoRegistro("FecInicio"), "ddMMyyyy")
'    sFecFin = Format(afoRegistro("FecFin"), "ddMMyyyy")
'    Set afoRegistro = Nothing

    iCodigo = CInt(Mid(cmbDirArea.Text, 1, InStr(1, cmbDirArea.Text, "-") - 2))
 
    sJerarquia = "DA"
    sPersona = "Director Area"
    
    If Not bDameFechasCampanya(CodPaisTrabajo, txtAnyo, txtCampanya, CByte(iCodigo), 4, sFecIni, sFecFin) Then
        txtAnyo.SetFocus
        Exit Sub
    End If
    
    Randomize
    sTablaTemp = "TG" & Int((1000000 * Rnd) + 1)

    sSql = "CREATE TEMP TABLE " & sTablaTemp & "(" & _
           "Fecha Date," & _
           "Importe Decimal(9,2)," & _
           "Numero Decimal(9,2)," & _
           "Donde Char(50)," & _
           "NombreGasto Char(20)," & _
           "Hora Char(2)," & _
           "DiaCam Char(3))"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al crear la tabla temporal.", vbApplicationModal + vbExclamation, "Error"
        Exit Sub
    End If

    sSql = "SELECT g.Importe,g.Numero,g.Donde,g.Fecha,t.NombreGasto,g.Hora " & _
           "FROM Gastos g,TiposGasto t " & _
           "WHERE Codigo = " & iCodigo & _
           " AND Jerarquia = '" & sJerarquia & _
           "' AND CodPais = " & CodPaisTrabajo & _
           " AND Campanya = " & txtCampanya & _
           " AND Anyo = " & txtAnyo & _
           " AND g.TipoGasto = t.TipoGasto AND t.Tipo = 'C' AND " & _
           "Validado = 'S' " & _
           "GROUP BY Fecha,NombreGasto,Importe,Numero,Donde,Hora " & _
           "ORDER BY Fecha,Hora,NombreGasto,Importe"
    
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        MsgBox "No existen datos.", vbApplicationModal + vbInformation, "Información"
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute (sSql)
        Exit Sub
    Else
    
        While Not afoRegistro.EOF
        
            sSql = "INSERT INTO " & sTablaTemp & "(Fecha,Numero,Donde,NombreGasto,Importe,Hora,DiaCam) " & _
                   "VALUES ('" & afoRegistro("Fecha") & "'," & afoRegistro("Numero") & ",'" & _
                   Trim(afoRegistro("Donde")) & "','" & Trim(afoRegistro("NombreGasto")) & "',"
            
            If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                sSql = sSql & dConvierteMoneda(CStr(afoRegistro("Importe")), 0, CByte(CodPaisTrabajo))
            Else
                sSql = sSql & afoRegistro("Importe")
            End If
            sSql = sSql & ",'" & NoNull(afoRegistro("Hora"), "A") & "','" & sDiaCampanya(afoRegistro("Fecha"), sFecIni, sFecFin) & "')"
            
            If Not gBdGALILEO.bExecute(sSql) Then
                MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                sSql = "DROP TABLE " & sTablaTemp
                gBdGALILEO.bExecute (sSql)
                Exit Sub
            End If
            
            afoRegistro.MoveNext
                        
        Wend
        
    End If
    
    Set afoRegistro = Nothing
    sSql = "SELECT * FROM " & sTablaTemp & " ORDER BY Fecha,Hora"
    
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        MsgBox "No existen datos.", vbApplicationModal + vbInformation, "Información"
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute (sSql)
        Exit Sub
    Else
        InformeGastos.DataMember = ""
        Set InformeGastos.DataSource = afoRegistro
    End If

    InformeGastos.Sections("CabezaPagina").Controls("etNombre").Caption = Trim(Mid(cmbDirArea.Text, InStr(1, cmbDirArea.Text, "-") + 1))
    InformeGastos.Sections("CabezaPagina").Controls("etCodigo").Caption = iCodigo
    InformeGastos.Sections("CabezaPagina").Controls("etTipoPersona").Caption = sPersona
    InformeGastos.Sections("CabezaPagina").Controls("etFecIni").Caption = sFecIni
    InformeGastos.Sections("CabezaPagina").Controls("etFecFin").Caption = sFecFin
    InformeGastos.Sections("CabezaPagina").Controls("etCampanya").Caption = "C" & Format(txtCampanya, "00") & "-" & txtAnyo
     
    InformeGastos.Show vbModal
    
    Set afoRegistro = Nothing
    sSql = "DROP TABLE " & sTablaTemp
    gBdGALILEO.bExecute sSql

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdAceptar2_Click()

    Dim sPersona As String
    Dim sJerarquia As String
    Dim afoRegistro As ADODB.Recordset
    Dim sSql As String
    Dim dAux As Double
    
    Dim iCodigo As Integer
    Dim sFecIni As String
    Dim sFecFin As String
    Dim lAleatorio As Long
    
    Dim iCont As Integer
    Dim sDiaRuta As String
     
    On Error GoTo Error_Rutina
    
    If cmbRegion.Text = "" Then
        cmbRegion.SetFocus
        Exit Sub
    End If
    If cmbDirArea.Text = "" Then
        cmbDirArea.SetFocus
        Exit Sub
    End If
    txtAnyo = Trim(txtAnyo)
    If txtAnyo = "" Then
        txtAnyo.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtAnyo) Then
        txtAnyo = ""
        txtAnyo.SetFocus
        Exit Sub
    End If
    txtCampanya = Trim(txtCampanya)
    If txtCampanya = "" Then
        txtCampanya.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtCampanya) Then
        txtCampanya = ""
        txtCampanya.SetFocus
        Exit Sub
    End If
'    If txtCampanya < 1 Or txtCampanya > 17 Then
'        txtCampanya = ""
'        txtCampanya.SetFocus
'        Exit Sub
'    End If
    
'    sSql = "SELECT FecInicio,FecFin FROM Campanyas " & _
'           "WHERE CodPais = " & CodPaisTrabajo & _
'           " AND Anyo = " & txtAnyo & _
'           " AND CodCampanya = " & txtCampanya
'    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
'        Set afoRegistro = Nothing
'        txtAnyo.SetFocus
'        Exit Sub
'    End If
'    If CLng(Format(afoRegistro("FecInicio"), "yyyyMMdd")) > CLng(Format(Now, "yyyyMMdd")) Then
'        Set afoRegistro = Nothing
'        txtAnyo = ""
'        txtCampanya = ""
'        txtAnyo.SetFocus
'        Exit Sub
'    End If
'    sFecIni = Format(afoRegistro("FecInicio"), "ddMMyyyy")
'    sFecFin = Format(afoRegistro("FecFin"), "ddMMyyyy")
'    Set afoRegistro = Nothing
 
    iCodigo = CInt(Mid(cmbDirArea.Text, 1, InStr(1, cmbDirArea.Text, "-") - 2))
 
    sJerarquia = "DA"
    sPersona = "Director Area"
    
    If Not bDameFechasCampanya(CodPaisTrabajo, txtAnyo, txtCampanya, CLng(iCodigo), 4, sFecIni, sFecFin) Then
        txtAnyo.SetFocus
        Exit Sub
    End If
    
    Randomize
    lAleatorio = Int((1000000 * Rnd) + 1)
    sSql = "DELETE FROM Inf_Gasto WHERE Aleatorio = " & lAleatorio
    gBdGALILEO.bExecute sSql

    sSql = "SELECT g.Importe,g.Numero,g.Donde,g.Fecha,t.NombreGasto,g.Hora " & _
           "FROM Gastos g,TiposGasto t " & _
           "WHERE Codigo = " & iCodigo & _
           " AND Jerarquia = '" & sJerarquia & _
           "' AND CodPais = " & CodPaisTrabajo & _
           " AND Campanya = " & txtCampanya & _
           " AND Anyo = " & txtAnyo & _
           " AND g.TipoGasto = t.TipoGasto AND t.Tipo = 'C' AND " & _
           "Validado = 'S' " & _
           "GROUP BY Fecha,NombreGasto,Importe,Numero,Donde,Hora " & _
           "ORDER BY Fecha,Hora,NombreGasto,Importe"
    
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        MsgBox "No existen datos.", vbApplicationModal + vbInformation, "Información"
        gBdGALILEO.bExecute (sSql)
        Exit Sub
    Else
    
        While Not afoRegistro.EOF
        
            sSql = "INSERT INTO Inf_Gasto(Aleatorio,Fecha,Numero,Donde,NombreGasto,Importe,Hora,DiaCam) " & _
                   "VALUES (" & lAleatorio & ",'" & afoRegistro("Fecha") & "'," & afoRegistro("Numero") & ",'" & _
                   Trim(afoRegistro("Donde")) & "','" & Trim(afoRegistro("NombreGasto")) & "',"
            
            If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                dAux = dConvierteMoneda(CStr(afoRegistro("Importe")), 0, CByte(CodPaisTrabajo))
            Else
                dAux = afoRegistro("Importe")
            End If
            If Trim(afoRegistro("NombreGasto")) = "Kilometros" Then dAux = dAux * afoRegistro("Numero")
            sSql = sSql & dAux & ",'" & NoNull(afoRegistro("Hora"), "A") & "','" & sDiaCampanya(afoRegistro("Fecha"), sFecIni, sFecFin) & "')"
            
            If Not gBdGALILEO.bExecute(sSql) Then
                MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                gBdGALILEO.bExecute (sSql)
                Exit Sub
            End If
            
            afoRegistro.MoveNext
                        
        Wend
        
    End If
    
    For iCont = 1 To 15
    
        Select Case iCont
        Case 1
            sDiaRuta = "1/1"
        Case 2
            sDiaRuta = "1/2"
        Case 3
            sDiaRuta = "1/3"
        Case 4
            sDiaRuta = "1/4"
        Case 5
            sDiaRuta = "1/5"
        Case 6
            sDiaRuta = "2/1"
        Case 7
            sDiaRuta = "2/2"
        Case 8
            sDiaRuta = "2/3"
        Case 9
            sDiaRuta = "2/4"
        Case 10
            sDiaRuta = "2/5"
        Case 11
            sDiaRuta = "3/1"
        Case 12
            sDiaRuta = "3/2"
        Case 13
            sDiaRuta = "3/3"
        Case 14
            sDiaRuta = "3/4"
        Case 15
            sDiaRuta = "3/5"
        End Select
    
        Set afoRegistro = Nothing
        sSql = "SELECT * FROM Inf_Gasto " & _
               "WHERE Aleatorio = " & lAleatorio & _
               " AND DiaCam = '" & sDiaRuta & "'"
        If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
            sSql = "INSERT INTO Inf_Gasto (Aleatorio,DiaCam) " & _
                   "VALUES (" & lAleatorio & ",'" & sDiaRuta & "')"
            gBdGALILEO.bExecute sSql
        End If
        Set afoRegistro = Nothing
        
    Next
    
    crpInforme.Reset
    crpInforme.Connect = gBdGALILEO.cn.ConnectionString
    crpInforme.ReportFileName = "\\ibmserver\Aplic\InformesCC\Gastos.rpt"
    crpInforme.WindowTitle = "INFORME GASTOS"
    
    sFecIni = Mid(sFecIni, 1, 2) & "/" & Mid(sFecIni, 3, 2) & "/" & Mid(sFecIni, 5)
    sFecFin = Mid(sFecFin, 1, 2) & "/" & Mid(sFecFin, 3, 2) & "/" & Mid(sFecFin, 5)
    crpInforme.ParameterFields(1) = "Aleatorio;" & lAleatorio & ";True"
    crpInforme.ParameterFields(2) = "Nombre;" & cmbDirArea.Text & ";True"
    crpInforme.ParameterFields(3) = "Jerarquia;" & sPersona & ";True"
    crpInforme.ParameterFields(4) = "FecIni;" & sFecIni & ";True"
    crpInforme.ParameterFields(5) = "FecFin;" & sFecFin & ";True"
    crpInforme.Destination = crptToWindow
    crpInforme.WindowState = crptMaximized
    If crpInforme.PrintReport <> 0 Then MsgBox "Error número:" & crpInforme.LastErrorNumber & " - " & crpInforme.LastErrorString, vbApplicationModal + vbCritical, "Información"
    crpInforme.DiscardSavedData = True
    
    sSql = "DELETE FROM Inf_Gasto WHERE Aleatorio = " & lAleatorio
    gBdGALILEO.bExecute sSql
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdAceptar3_Click()

    Dim sPersona As String
    Dim sJerarquia As String
    Dim afoRegistro As ADODB.Recordset
    Dim sSql As String
    Dim dAux As Double
    
    Dim iCodigo As Integer
    Dim sFecIni As String
    Dim sFecFin As String
    Dim sNombreDirRegional As String
    Dim lAleatorio As Long
    
    Dim iCont As Integer
    Dim sDiaRuta As String
     
    On Error GoTo Error_Rutina
    
    If cmbRegion.Text = "" Then
        cmbRegion.SetFocus
        Exit Sub
    End If
    txtAnyo = Trim(txtAnyo)
    If txtAnyo = "" Then
        txtAnyo.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtAnyo) Then
        txtAnyo = ""
        txtAnyo.SetFocus
        Exit Sub
    End If
    txtCampanya = Trim(txtCampanya)
    If txtCampanya = "" Then
        txtCampanya.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtCampanya) Then
        txtCampanya = ""
        txtCampanya.SetFocus
        Exit Sub
    End If
'    If txtCampanya < 1 Or txtCampanya > 17 Then
'        txtCampanya = ""
'        txtCampanya.SetFocus
'        Exit Sub
'    End If
    
'    sSql = "SELECT FecInicio,FecFin FROM Campanyas " & _
'           "WHERE CodPais = " & CodPaisTrabajo & _
'           " AND Anyo = " & txtAnyo & _
'           " AND CodCampanya = " & txtCampanya
'    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
'        Set afoRegistro = Nothing
'        txtAnyo.SetFocus
'        Exit Sub
'    End If
'    If CLng(Format(afoRegistro("FecInicio"), "yyyyMMdd")) > CLng(Format(Now, "yyyyMMdd")) Then
'        Set afoRegistro = Nothing
'        txtAnyo = ""
'        txtCampanya = ""
'        txtAnyo.SetFocus
'        Exit Sub
'    End If
'    sFecIni = Format(afoRegistro("FecInicio"), "ddMMyyyy")
'    sFecFin = Format(afoRegistro("FecFin"), "ddMMyyyy")
'    Set afoRegistro = Nothing
 
    iCodigo = CInt(Mid(cmbRegion.Text, 1, InStr(1, cmbRegion.Text, "-") - 2))
    sSql = "SELECT Codigo,Nombre FROM Directores " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " AND CodGeografia = " & iCodigo & _
           " AND Jerarquia = 'DR'"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        cmbRegion.SetFocus
        Exit Sub
    End If
    iCodigo = afoRegistro("Codigo")
    sNombreDirRegional = Trim(afoRegistro("Nombre"))
    Set afoRegistro = Nothing
 
    sJerarquia = "DR"
    sPersona = "Director Regional"
    
    If Not bDameFechasCampanya(CodPaisTrabajo, txtAnyo, txtCampanya, CLng(iCodigo), 5, sFecIni, sFecFin) Then
        txtAnyo.SetFocus
        Exit Sub
    End If
    
    Randomize
    lAleatorio = Int((1000000 * Rnd) + 1)
    sSql = "DELETE FROM Inf_Gasto WHERE Aleatorio = " & lAleatorio
    gBdGALILEO.bExecute sSql

    sSql = "SELECT g.Importe,g.Numero,g.Donde,g.Fecha,t.NombreGasto,g.Hora " & _
           "FROM Gastos g,TiposGasto t " & _
           "WHERE Codigo = " & iCodigo & _
           " AND Jerarquia = '" & sJerarquia & _
           "' AND CodPais = " & CodPaisTrabajo & _
           " AND Campanya = " & txtCampanya & _
           " AND Anyo = " & txtAnyo & _
           " AND g.TipoGasto = t.TipoGasto AND t.Tipo = 'C' AND " & _
           "Validado = 'S' " & _
           "GROUP BY Fecha,NombreGasto,Importe,Numero,Donde,Hora " & _
           "ORDER BY Fecha,Hora,NombreGasto,Importe"
    
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        MsgBox "No existen datos.", vbApplicationModal + vbInformation, "Información"
        gBdGALILEO.bExecute (sSql)
        Exit Sub
    Else
    
        While Not afoRegistro.EOF
        
            sSql = "INSERT INTO Inf_Gasto(Aleatorio,Fecha,Numero,Donde,NombreGasto,Importe,Hora,DiaCam) " & _
                   "VALUES (" & lAleatorio & ",'" & afoRegistro("Fecha") & "'," & afoRegistro("Numero") & ",'" & _
                   Trim(afoRegistro("Donde")) & "','" & Trim(afoRegistro("NombreGasto")) & "',"
            
            If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                dAux = dConvierteMoneda(CStr(afoRegistro("Importe")), 0, CByte(CodPaisTrabajo))
            Else
                dAux = afoRegistro("Importe")
            End If
            If Trim(afoRegistro("NombreGasto")) = "Kilometros" Then dAux = dAux * afoRegistro("Numero")
            sSql = sSql & dAux & ",'" & NoNull(afoRegistro("Hora"), "A") & "','" & sDiaCampanya(afoRegistro("Fecha"), sFecIni, sFecFin) & "')"
            
            If Not gBdGALILEO.bExecute(sSql) Then
                MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                gBdGALILEO.bExecute (sSql)
                Exit Sub
            End If
            
            afoRegistro.MoveNext
                        
        Wend
        
    End If
    
    For iCont = 1 To 15
    
        Select Case iCont
        Case 1
            sDiaRuta = "1/1"
        Case 2
            sDiaRuta = "1/2"
        Case 3
            sDiaRuta = "1/3"
        Case 4
            sDiaRuta = "1/4"
        Case 5
            sDiaRuta = "1/5"
        Case 6
            sDiaRuta = "2/1"
        Case 7
            sDiaRuta = "2/2"
        Case 8
            sDiaRuta = "2/3"
        Case 9
            sDiaRuta = "2/4"
        Case 10
            sDiaRuta = "2/5"
        Case 11
            sDiaRuta = "3/1"
        Case 12
            sDiaRuta = "3/2"
        Case 13
            sDiaRuta = "3/3"
        Case 14
            sDiaRuta = "3/4"
        Case 15
            sDiaRuta = "3/5"
        End Select
    
        Set afoRegistro = Nothing
        sSql = "SELECT * FROM Inf_Gasto " & _
               "WHERE Aleatorio = " & lAleatorio & _
               " AND DiaCam = '" & sDiaRuta & "'"
        If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
            sSql = "INSERT INTO Inf_Gasto (Aleatorio,DiaCam) " & _
                   "VALUES (" & lAleatorio & ",'" & sDiaRuta & "')"
            gBdGALILEO.bExecute sSql
        End If
        Set afoRegistro = Nothing
        
    Next
    
    crpInforme.Reset
    crpInforme.Connect = gBdGALILEO.cn.ConnectionString
    crpInforme.ReportFileName = "\\ibmserver\Aplic\InformesCC\Gastos.rpt"
    crpInforme.WindowTitle = "INFORME GASTOS"
    
    sFecIni = Mid(sFecIni, 1, 2) & "/" & Mid(sFecIni, 3, 2) & "/" & Mid(sFecIni, 5)
    sFecFin = Mid(sFecFin, 1, 2) & "/" & Mid(sFecFin, 3, 2) & "/" & Mid(sFecFin, 5)
    crpInforme.ParameterFields(1) = "Aleatorio;" & lAleatorio & ";True"
    crpInforme.ParameterFields(2) = "Nombre;" & sNombreDirRegional & ";True"
    crpInforme.ParameterFields(3) = "Jerarquia;" & sPersona & ";True"
    crpInforme.ParameterFields(4) = "FecIni;" & sFecIni & ";True"
    crpInforme.ParameterFields(5) = "FecFin;" & sFecFin & ";True"
    crpInforme.Destination = crptToWindow
    crpInforme.WindowState = crptMaximized
    If crpInforme.PrintReport <> 0 Then MsgBox "Error número:" & crpInforme.LastErrorNumber & " - " & crpInforme.LastErrorString, vbApplicationModal + vbCritical, "Información"
    crpInforme.DiscardSavedData = True
    
    sSql = "DELETE FROM Inf_Gasto WHERE Aleatorio = " & lAleatorio
    gBdGALILEO.bExecute sSql
    
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
    
    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    On Error GoTo Error_Rutina

    sSql = "SELECT * FROM Regiones " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " ORDER BY CodRegion"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Sub
    End If
    
    While Not afoRegistro.EOF
        cmbRegion.AddItem afoRegistro("CodRegion") & " - " & Trim(afoRegistro("Nombre"))
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

Private Sub CargarDirArea()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    Dim iRegion As Integer
    
    On Error GoTo Error_Rutina
    
    VaciarBox cmbDirArea
    If cmbRegion.Text = "" Then Exit Sub

    iRegion = CInt(Mid(cmbRegion.Text, 1, InStr(1, cmbRegion.Text, "-") - 2))
    sSql = "SELECT * FROM Directores " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " AND Jerarquia = 'DA' " & _
           "AND FechaBaja IS NULL " & _
           "AND CodGeografia IN (SELECT CodArea FROM Areas " & _
                                "WHERE CodPais = " & CodPaisTrabajo & _
                                " AND CodRegion = " & iRegion & _
           ") ORDER BY Codigo"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Sub
    End If
    
    While Not afoRegistro.EOF
        cmbDirArea.AddItem afoRegistro("Codigo") & " - " & Trim(afoRegistro("Nombre"))
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

Private Sub VaciarBox(Control As Control)
    
    Dim iCont As Integer
    
    For iCont = 0 To Control.ListCount - 1
        Control.RemoveItem 0
    Next iCont
    
    Control.Clear

End Sub

Private Sub txtAnyo_GotFocus()
    TextGotFocus txtAnyo
End Sub

Private Sub txtAnyo_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, txtAnyo, 4
End Sub

Private Sub txtAnyo_LostFocus()
    TextLostFocus txtAnyo
End Sub

Private Sub txtCampanya_GotFocus()
    TextGotFocus txtCampanya
End Sub

Private Sub txtCampanya_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, txtCampanya, 2
End Sub

Private Sub txtCampanya_LostFocus()
    TextLostFocus txtCampanya
End Sub

Private Function sDiaCampanya(sFecha As String, sFecIni As String, _
    sFecFin As String) As String

    Dim fFecIni As Date
    Dim fFecFin As Date
    Dim fFecha As Date
    Dim bEnc As Boolean
    Dim iCont As Integer
    
    On Error GoTo Error_Rutina
    
    fFecha = sFecha
    fFecIni = Mid(sFecIni, 1, 2) & "/" & Mid(sFecIni, 3, 2) & "/" & Mid(sFecIni, 5)
    fFecFin = Mid(sFecFin, 1, 2) & "/" & Mid(sFecFin, 3, 2) & "/" & Mid(sFecFin, 5)
    
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

