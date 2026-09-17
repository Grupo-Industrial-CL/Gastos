VERSION 5.00
Begin VB.Form frmInformeRegion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos - Informe de Gastos"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ControlBox      =   0   'False
   Icon            =   "frmInformeRegion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDeducible 
      Height          =   315
      ItemData        =   "frmInformeRegion.frx":08CA
      Left            =   4800
      List            =   "frmInformeRegion.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox cmbFamiliaGasto 
      Height          =   315
      ItemData        =   "frmInformeRegion.frx":08CE
      Left            =   1560
      List            =   "frmInformeRegion.frx":08D0
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtAnyo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Top             =   1280
      Width           =   615
   End
   Begin VB.TextBox txtCampanyaFin 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Top             =   1280
      Width           =   375
   End
   Begin VB.TextBox txtCampanyaInicio 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1280
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox cmbTipoGasto 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   795
      Width           =   1815
   End
   Begin VB.ComboBox cmbRegion 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   315
      Width           =   5055
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblDeducible 
      AutoSize        =   -1  'True
      Caption         =   "Deducible"
      Height          =   195
      Left            =   3960
      TabIndex        =   15
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Familia Gasto"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Año"
      Height          =   195
      Index           =   4
      Left            =   3960
      TabIndex        =   13
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Campaña fin"
      Height          =   195
      Index           =   3
      Left            =   2160
      TabIndex        =   12
      Top             =   1320
      Width           =   885
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Campaña inicio"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de gasto"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione región"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1275
   End
End
Attribute VB_Name = "frmInformeRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub cmdAceptar_Click()
    
    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    Dim afoRegAux As ADODB.Recordset
    Dim iRegion As Integer
    Dim sTablaTemp1 As String
    Dim sTablaTemp2 As String
    
    On Error GoTo Error_Rutina
    
    If cmbRegion = "" Then
        cmbRegion.SetFocus
        Exit Sub
    End If
    If cmbTipoGasto = "" Then
        cmbTipoGasto.SetFocus
        Exit Sub
    End If
    If txtCampanyaInicio = "" Then
        txtCampanyaInicio.SetFocus
        Exit Sub
    Else
'        If txtCampanyaInicio < 0 Or txtCampanyaInicio > 17 Then
'            txtCampanyaInicio.SetFocus
'            Exit Sub
'        End If
    End If
    If txtCampanyaFin = "" Then
        txtCampanyaFin.SetFocus
        Exit Sub
    Else
'        If txtCampanyaFin < 0 Or txtCampanyaFin > 17 Then
'            txtCampanyaFin.SetFocus
'            Exit Sub
'        End If
    End If
    If CInt(txtCampanyaInicio) > CInt(txtCampanyaFin) Then
        MsgBox "La campaña de inicio debe ser mayor que la campaña de fin", vbApplicationModal + vbExclamation, "Información"
        txtCampanyaInicio.SetFocus
        Exit Sub
    End If
    If txtAnyo = "" Then
        txtAnyo.SetFocus
        Exit Sub
    End If
    
    If cmbFamiliaGasto.Text = "" Then
        cmbFamiliaGasto.SetFocus
        Exit Sub
    End If

    Screen.MousePointer = 11
    
    Randomize
    sTablaTemp1 = "TGR" & Int((1000000 * Rnd) + 1)
    sTablaTemp2 = "TD" & Int((1000000 * Rnd) + 1)
    
    sSql = "CREATE TEMP TABLE " & sTablaTemp1 & _
           "(Codigo Integer," & _
           "Jerarquia Char(2)," & _
           "Nombre Char(50)," & _
           "Total Decimal(15,2),Matricula Char(15))"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al crear la tabla temporal GASTOS.", vbApplicationModal + vbExclamation, "Error"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    sSql = "CREATE TEMP TABLE " & sTablaTemp2 & _
           "(Codigo Integer," & _
           "Jerarquia Char(2),Matricula Char(15))"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al crear la tabla temporal DIRECTORES.", vbApplicationModal + vbExclamation, "Error"
        sSql = "DROP TABLE " & sTablaTemp1
        gBdGALILEO.bExecute sSql
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    iRegion = CInt(Mid(cmbRegion, 1, InStr(1, cmbRegion, "-") - 2))
    sSql = "INSERT INTO " & sTablaTemp2 & " SELECT DISTINCT Codigo,Jerarquia,NVL(Matricula,'') Matricula " & _
           "From Directores " & _
           "WHERE CodGeografia IN (SELECT DISTINCT CodArea " & _
                                  "FROM Areas " & _
                                  "WHERE CodPais = " & CodPaisTrabajo & _
                                  " AND CodRegion = " & iRegion & ") " & _
           "AND CodPais = " & CodPaisTrabajo & _
           " AND Jerarquia IN ('DA','DP')"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al hacer insert temporal.", vbApplicationModal + vbExclamation, "Información"
        sSql = "DROP TABLE " & sTablaTemp1
        gBdGALILEO.bExecute sSql
        sSql = "DROP TABLE " & sTablaTemp2
        gBdGALILEO.bExecute sSql
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    sSql = "INSERT INTO " & sTablaTemp2 & " SELECT DISTINCT Codigo,Jerarquia,NVL(Matricula,'') Matricula " & _
           "FROM Directores d,Regiones r " & _
           "WHERE d.CodPais = r.CodPais " & _
           "AND d.CodPais = " & CodPaisTrabajo & _
           " AND r.CodRegion = " & iRegion & _
           " AND d.Nombre = r.Responsable " & _
           "AND Jerarquia = 'DR'"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al hacer segundo insert temporal.", vbApplicationModal + vbExclamation, "Información"
        sSql = "DROP TABLE " & sTablaTemp1
        gBdGALILEO.bExecute sSql
        sSql = "DROP TABLE " & sTablaTemp2
        gBdGALILEO.bExecute sSql
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If Trim(cmbTipoGasto) = "Todos" Then
        
        sSql = "SELECT g.Codigo,g.Jerarquia,Nombre,t.Matricula,SUM(Importe * Numero) AS Total " & _
               "FROM Gastos g,Directores d," & sTablaTemp2 & " t " & _
               "WHERE g.Codigo = d.Codigo AND g.Codigo = t.Codigo " & _
               "AND g.Jerarquia = d.Jerarquia AND g.Jerarquia = t.Jerarquia " & _
               "AND g.CodPais = d.CodPais AND g.CodPais = " & CodPaisTrabajo & _
               " AND Campanya BETWEEN " & txtCampanyaInicio & " AND " & txtCampanyaFin & _
               " AND g.TipoGasto = 'K' " & _
               "AND Anyo = " & txtAnyo & _
               " AND Importe <> 0 AND " & _
               "Validado = 'S' " & _
               "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,g.Jerarquia,Nombre,t.Matricula " & _
               "ORDER BY g.Codigo"
        
        gBdGALILEO.bRecordset sSql, afoRegistro
        
        While Not afoRegistro.EOF
        
            sSql = "INSERT INTO " & sTablaTemp1 & "(Codigo,Jerarquia,Nombre,Matricula,Total) " & _
                   "VALUES (" & afoRegistro("Codigo") & ",'" & _
                   Trim(afoRegistro("Jerarquia")) & "','" & _
                   Trim(afoRegistro("Nombre")) & "','" & Trim(afoRegistro("Matricula")) & "',"
                   
            If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                sSql = sSql & dConvierteMoneda(CStr(afoRegistro("Total")), 0, CByte(CodPaisTrabajo)) & ")"
            Else
                sSql = sSql & afoRegistro("Total") & ")"
            End If
            
            If Not gBdGALILEO.bExecute(sSql) Then
                MsgBox "Error al insertar datos.", vbApplicationModal + vbExclamation, "Error"
                sSql = "DROP TABLE " & sTablaTemp1
                gBdGALILEO.bExecute sSql
                sSql = "DROP TABLE " & sTablaTemp2
                gBdGALILEO.bExecute sSql
                Set afoRegistro = Nothing
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
                   
            afoRegistro.MoveNext
            
        Wend
        Set afoRegistro = Nothing
        
        sSql = "SELECT g.Codigo,g.Jerarquia,Nombre,t.Matricula,SUM(Importe) AS Total " & _
               "FROM Gastos g,Directores d," & sTablaTemp2 & " t " & _
               "WHERE g.Codigo = d.Codigo AND g.Codigo = t.Codigo " & _
               "AND g.Jerarquia = d.Jerarquia AND g.Jerarquia = t.Jerarquia " & _
               "AND g.CodPais = d.CodPais AND g.CodPais = " & CodPaisTrabajo & _
               " AND Campanya BETWEEN " & txtCampanyaInicio & " AND " & txtCampanyaFin & _
               " AND g.TipoGasto <> 'K' " & _
               "AND Anyo = " & txtAnyo & _
               " AND Importe <> 0 AND " & _
               "Validado = 'S' " & _
               "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,g.Jerarquia,Nombre,t.Matricula " & _
               "ORDER BY g.Codigo"
        
        gBdGALILEO.bRecordset sSql, afoRegistro
        
        While Not afoRegistro.EOF
        
            sSql = "SELECT * FROM " & sTablaTemp1 & _
                   " WHERE Codigo = " & afoRegistro("Codigo") & _
                   " AND Jerarquia = '" & afoRegistro("Jerarquia") & "'"
            
            Set afoRegAux = Nothing
            If gBdGALILEO.bRecordset(sSql, afoRegAux) Then
                sSql = "UPDATE " & sTablaTemp1 & " SET Total = Total + "
                
                If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                    sSql = sSql & dConvierteMoneda(CStr(afoRegistro("Total")), 0, CByte(CodPaisTrabajo))
                Else
                    sSql = sSql & afoRegistro("Total")
                End If
                
                sSql = sSql & " WHERE Codigo = " & afoRegistro("Codigo") & _
                       " AND Jerarquia = '" & afoRegistro("Jerarquia") & "'"
            Else
                sSql = "INSERT INTO " & sTablaTemp1 & " (Codigo,Jerarquia,Nombre,Matricula,Total) " & _
                       "VALUES (" & afoRegistro("Codigo") & ",'" & _
                       Trim(afoRegistro("Jerarquia")) & "','" & _
                       Trim(afoRegistro("Nombre")) & "','" & Trim(afoRegistro("Matricula")) & "',"
                       
                If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                    sSql = sSql & dConvierteMoneda(CStr(afoRegistro("Total")), 0, CByte(CodPaisTrabajo)) & ")"
                Else
                    sSql = sSql & afoRegistro("Total") & ")"
                End If
            End If
            
            If Not gBdGALILEO.bExecute(sSql) Then
                MsgBox "Error al insertar datos.", vbApplicationModal + vbExclamation, "Error"
                sSql = "DROP TABLE " & sTablaTemp1
                gBdGALILEO.bExecute sSql
                sSql = "DROP TABLE " & sTablaTemp2
                gBdGALILEO.bExecute sSql
                Set afoRegistro = Nothing
                Set afoRegAux = Nothing
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
                   
            afoRegistro.MoveNext
            
        Wend
        Set afoRegistro = Nothing
        Set afoRegAux = Nothing
    
    Else
        
        If Trim(cmbTipoGasto) = "Kilometros" Then
            sSql = "SELECT g.Codigo,g.Jerarquia,Nombre,t.Matricula,SUM(Importe * Numero) AS Total "
        Else
            sSql = "SELECT g.Codigo,g.Jerarquia,Nombre,t.Matricula,SUM(Importe) AS Total "
        End If
        sSql = sSql & "FROM Gastos g,TiposGasto a,Directores d," & sTablaTemp2 & " t " & _
               "WHERE g.Codigo = d.Codigo AND g.Codigo = t.Codigo " & _
               "AND g.Jerarquia = d.Jerarquia AND g.Jerarquia = t.Jerarquia " & _
               "AND g.CodPais = d.CodPais AND g.CodPais = " & CodPaisTrabajo & _
               " AND Campanya BETWEEN " & txtCampanyaInicio & " AND " & txtCampanyaFin & _
               " AND NombreGasto = '" & Trim(cmbTipoGasto) & "' " & _
               "AND g.TipoGasto = a.TipoGasto AND a.Tipo = 'C' " & _
               "AND Anyo = " & txtAnyo & _
               " AND Importe <> 0 AND " & _
               "Validado = 'S' " & _
               "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,g.Jerarquia,Nombre,t.Matricula " & _
               "ORDER BY g.Codigo"
        
        If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
            MsgBox "No existen gastos con las condiciones introducidas.", vbApplicationModal + vbInformation, "Información"
            sSql = "DROP TABLE " & sTablaTemp1
            gBdGALILEO.bExecute sSql
            sSql = "DROP TABLE " & sTablaTemp2
            gBdGALILEO.bExecute sSql
            Set afoRegistro = Nothing
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        While Not afoRegistro.EOF
        
            sSql = "INSERT INTO " & sTablaTemp1 & " (Codigo,Jerarquia,Nombre,Matricula,Total) " & _
                   "VALUES (" & afoRegistro("Codigo") & ",'" & _
                   Trim(afoRegistro("Jerarquia")) & "','" & _
                   Trim(afoRegistro("Nombre")) & "','" & Trim(afoRegistro("Matricula")) & "',"
                   
            If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                sSql = sSql & dConvierteMoneda(CStr(afoRegistro("Total")), 0, CByte(CodPaisTrabajo)) & ")"
            Else
                sSql = sSql & afoRegistro("Total") & ")"
            End If
            
            If Not gBdGALILEO.bExecute(sSql) Then
                MsgBox "Error al insertar datos.", vbApplicationModal + vbExclamation, "Error"
                sSql = "DROP TABLE " & sTablaTemp1
                gBdGALILEO.bExecute sSql
                sSql = "DROP TABLE " & sTablaTemp2
                gBdGALILEO.bExecute sSql
                Set afoRegistro = Nothing
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
                   
            afoRegistro.MoveNext
            
        Wend
        Set afoRegistro = Nothing
    
    End If
    
    sSql = "SELECT * FROM " & sTablaTemp1
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        MsgBox "No existen gastos con las condiciones introducidas.", vbApplicationModal + vbInformation, "Información"
        sSql = "DROP TABLE " & sTablaTemp1
        gBdGALILEO.bExecute sSql
        sSql = "DROP TABLE " & sTablaTemp2
        gBdGALILEO.bExecute sSql
        Set afoRegistro = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    dtrInformeRegion.DataMember = ""
    Set dtrInformeRegion.DataSource = afoRegistro
        
    dtrInformeRegion.Sections("EncabezadoInforme").Controls("lblPeriodo").Caption = "Entre campaña " & txtCampanyaInicio & " y campaña " & txtCampanyaFin & " de " & txtAnyo
    dtrInformeRegion.Sections("EncabezadoInforme").Controls("lblRegion").Caption = "REGIÓN: " & Trim(cmbRegion)
    dtrInformeRegion.Sections("EncabezadoInforme").Controls("etFamiliaGasto").Caption = "Familia gasto: " & cmbFamiliaGasto.Text
    
    dtrInformeRegion.Sections("EncabezadoPagina").Controls("etMatricula").Caption = ""
    If CodPaisTrabajo = giPAIS_PORTUGAL Then dtrInformeRegion.Sections("EncabezadoPagina").Controls("etMatricula").Caption = "Matrícula"
    
    Screen.MousePointer = 0
    dtrInformeRegion.Show vbModal
    
    sSql = "DROP TABLE " & sTablaTemp1
    gBdGALILEO.bExecute sSql
    sSql = "DROP TABLE " & sTablaTemp2
    gBdGALILEO.bExecute sSql
    Set afoRegistro = Nothing
    Screen.MousePointer = vbDefault
    
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

    RellenarComboRegiones
    RellenarComboTipoGastos
    
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

Private Sub RellenarComboRegiones()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    On Error GoTo Error_Rutina
    
    sSql = "SELECT CodRegion,Nombre,Responsable " & _
           "FROM Regiones " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " AND Nombre NOT MATCHES '*LAY*'" & _
           " AND Nombre NOT MATCHES '*CRISTIAN*'" & _
           " AND Nombre NOT MATCHES '*CAJA*'" & _
           " AND Nombre NOT MATCHES '*DIEGO*'" & _
           " AND Nombre NOT MATCHES '*MOROSIDAD*'" & _
           " AND Nombre NOT MATCHES '*PERSONAL*'"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        MsgBox "Error al cargar las regiones.", vbApplicationModal + vbExclamation, "Información"
        Set afoRegistro = Nothing
        Exit Sub
    End If
        
    While Not afoRegistro.EOF
        cmbRegion.AddItem afoRegistro("CodRegion") & " - " & _
            Trim(afoRegistro("Nombre")) & " (" & Trim(afoRegistro("Responsable")) & ")"
        afoRegistro.MoveNext
    Wend
    cmbRegion.ListIndex = 0
    Set afoRegistro = Nothing
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub RellenarComboTipoGastos()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    On Error GoTo Error_Rutina
    
    sSql = "SELECT NombreGasto FROM TiposGasto WHERE Tipo = 'C' ORDER BY NombreGasto"
    
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        MsgBox "Error al cargar tipos de gasto.", vbApplicationModal + vbExclamation, "Información"
        Set afoRegistro = Nothing
        Exit Sub
    End If
    
    cmbTipoGasto.AddItem "Todos"
    While Not afoRegistro.EOF
       cmbTipoGasto.AddItem (afoRegistro("NombreGasto"))
       afoRegistro.MoveNext
    Wend
    cmbTipoGasto.ListIndex = 0
    Set afoRegistro = Nothing
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub txtAnyo_GotFocus()
    TextGotFocus txtAnyo
End Sub

Private Sub txtAnyo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmbFamiliaGasto.SetFocus
        Exit Sub
    End If
    
    EntraNumero KeyAscii, txtAnyo, 4
    
End Sub

Private Sub txtAnyo_LostFocus()
    TextLostFocus txtAnyo
End Sub

Private Sub txtCampanyaFin_GotFocus()
    TextGotFocus txtCampanyaFin
End Sub

Private Sub txtCampanyaFin_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtAnyo.SetFocus
        Exit Sub
    End If
    
    EntraNumero KeyAscii, txtCampanyaFin, 2
    
End Sub

Private Sub txtCampanyaFin_LostFocus()
    TextLostFocus txtCampanyaFin
End Sub

Private Sub txtCampanyaInicio_GotFocus()
    TextGotFocus txtCampanyaInicio
End Sub

Private Sub txtCampanyaInicio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtCampanyaFin.SetFocus
        Exit Sub
    End If
    
    EntraNumero KeyAscii, txtCampanyaInicio, 2
    
End Sub

Private Sub txtCampanyaInicio_LostFocus()
    TextLostFocus txtCampanyaInicio
End Sub
