VERSION 5.00
Begin VB.Form frmInformeAgrupado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de gastos agrupados"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   Icon            =   "frmInformeAgrupado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDeducible 
      Height          =   315
      ItemData        =   "frmInformeAgrupado.frx":08CA
      Left            =   2520
      List            =   "frmInformeAgrupado.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox cmbFamiliaGasto 
      Height          =   315
      ItemData        =   "frmInformeAgrupado.frx":08CE
      Left            =   240
      List            =   "frmInformeAgrupado.frx":08D0
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H8000000B&
      Caption         =   "&Aceptar"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   480
      Picture         =   "frmInformeAgrupado.frx":08D2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtAnyo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtCampanyaFin 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   1
      Top             =   555
      Width           =   375
   End
   Begin VB.TextBox txtCampanyaInicio 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   0
      Top             =   200
      Width           =   375
   End
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "&Volver"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   2400
      Picture         =   "frmInformeAgrupado.frx":0D1C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jerarquía"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   3315
      Begin VB.OptionButton optDirectores 
         Caption         =   "Directores"
         Height          =   195
         Left            =   255
         TabIndex        =   5
         Top             =   285
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optSupervisoras 
         Caption         =   "Supervisoras"
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   280
         Width           =   1335
      End
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Deducible"
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
      Index           =   4
      Left            =   2520
      TabIndex        =   14
      Top             =   960
      Width           =   870
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Familia gasto"
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
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   1125
   End
   Begin VB.Label lblEtiqueta 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Año"
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
      Left            =   2655
      TabIndex        =   12
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Campaña fin"
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
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Campaña inicio"
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
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   1305
   End
End
Attribute VB_Name = "frmInformeAgrupado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub cmdAceptar_Click()

    Dim afoGastos As ADODB.Recordset
    Dim afoRegistro As ADODB.Recordset
    Dim sTablaTemp As String

    On Error GoTo Error_Rutina
    
    If txtCampanyaInicio = "" Then
        txtCampanyaInicio.SetFocus
        Exit Sub
    Else
'        If txtCampanyaInicio = 0 Or txtCampanyaInicio > 17 Then
'            txtCampanyaInicio.SetFocus
'            Exit Sub
'        End If
    End If
    If txtCampanyaFin = "" Then
        txtCampanyaFin.SetFocus
        Exit Sub
    Else
'        If txtCampanyaFin = 0 Or txtCampanyaFin > 17 Then
'            txtCampanyaFin.SetFocus
'            Exit Sub
'        End If
    End If
    If txtAnyo = "" Then
        txtAnyo.SetFocus
        Exit Sub
    Else
        If txtAnyo = 0 Or Len(txtAnyo) < 4 Or txtAnyo > Year(Now) Then
            txtAnyo.SetFocus
            Exit Sub
        End If
    End If
    
    If CInt(txtCampanyaInicio) > CInt(txtCampanyaFin) Then
        MsgBox "La campaña de inicio debe ser mayor que la campaña de fin.", vbApplicationModal + vbExclamation, "Información"
        txtCampanyaInicio.SetFocus
        Exit Sub
    End If
    
    If cmbFamiliaGasto.Text = "" Then
        cmbFamiliaGasto.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    Randomize
    sTablaTemp = "TG" & Int((1000000 * Rnd) + 1)
    
    sSql = "CREATE TEMP TABLE " & sTablaTemp & "(Codigo INTEGER,Nombre CHAR(50),Matricula CHAR(15)," & _
           "TotalKG DECIMAL(15,2),TotalTel DECIMAL(15,2),TotalResto DECIMAL(15,2))"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al crear la tabla temporal.", vbApplicationModal + vbExclamation, "Error"
        Exit Sub
    End If
    
    If optDirectores Then
        
        'calculamos los gastos de TELÉFONO
        sSql = "SELECT g.Codigo,Nombre,NVL(d.Matricula,'') Matricula,SUM(Importe) AS Total " & _
               "FROM Gastos g,Directores d " & _
               "WHERE g.Codigo = d.Codigo AND " & _
               "g.Jerarquia = d.Jerarquia AND " & _
               "g.Jerarquia <> 'SU' AND " & _
               "g.CodPais = d.CodPais AND " & _
               "Campanya BETWEEN " & txtCampanyaInicio & " AND " & txtCampanyaFin & " AND " & _
               "g.TipoGasto = 'F' AND " & _
               "g.CodPais = " & CodPaisTrabajo & " AND " & _
               "Anyo = " & txtAnyo & " AND " & _
               "Importe <> 0 AND " & _
               "Validado = 'S' " & _
               "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,Nombre,d.Matricula " & _
               "ORDER BY g.Codigo"
        
        If gBdGALILEO.bRecordset(sSql, afoGastos) Then
        
            While Not afoGastos.EOF
            
                sSql = "INSERT INTO " & sTablaTemp & "(Codigo,Nombre,Matricula,TotalKG,TotalTel,TotalResto) " & _
                       "VALUES(" & afoGastos("Codigo") & ",'" & _
                       Trim(afoGastos("Nombre")) & "','" & Trim(afoGastos("Matricula")) & "',0,"
                
                If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                    sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo)) & ",0)"
                Else
                    sSql = sSql & afoGastos("Total") & ",0)"
                End If
                
                If Not gBdGALILEO.bExecute(sSql) Then
                    
                    MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                    sSql = "DROP TABLE " & sTablaTemp
                    gBdGALILEO.bExecute (sSql)
                    Exit Sub
                    
                End If
                
                afoGastos.MoveNext
            
            Wend
        
        End If
        Set afoGastos = Nothing
        
        'Calculamos los gastos que NO sean de TELÉFONO,GASOLINA O KILÓMETROS
        sSql = "SELECT g.Codigo,Nombre,NVL(d.Matricula,'') Matricula,SUM(Importe) AS Total " & _
               "FROM Gastos g,Directores d " & _
               "WHERE g.Codigo = d.Codigo AND " & _
               "g.Jerarquia = d.Jerarquia AND " & _
               "g.Jerarquia <> 'SU' AND " & _
               "g.CodPais = d.CodPais AND " & _
               "Campanya BETWEEN " & txtCampanyaInicio & " AND " & txtCampanyaFin & " AND " & _
               "g.TipoGasto NOT IN ('F','K','G') AND " & _
               "g.CodPais = " & CodPaisTrabajo & " AND " & _
               "Anyo = " & txtAnyo & " AND " & _
               "Importe <> 0 AND " & _
               "Validado = 'S' " & _
               " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,Nombre,d.Matricula " & _
               "ORDER BY g.Codigo"
        
        If gBdGALILEO.bRecordset(sSql, afoGastos) Then
        
            While Not afoGastos.EOF
            
                Set afoRegistro = Nothing
                sSql = "SELECT * FROM " & sTablaTemp & _
                       " WHERE Codigo = " & afoGastos("Codigo")
                If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
                
                    sSql = "UPDATE " & sTablaTemp & " SET TotalResto = "
                    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                        sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo))
                    Else
                        sSql = sSql & afoGastos("Total")
                    End If
                    sSql = sSql & " WHERE Codigo = " & afoGastos("Codigo")
                                        
                Else
            
                    sSql = "INSERT INTO " & sTablaTemp & "(Codigo,Nombre,Matricula,TotalKG,TotalTel,TotalResto) " & _
                           "VALUES(" & afoGastos("Codigo") & ",'" & _
                           Trim(afoGastos("Nombre")) & "','" & Trim(afoGastos("Matricula")) & "',0,0,"
                    
                    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                        sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo)) & ")"
                    Else
                        sSql = sSql & afoGastos("Total") & ")"
                    End If
                
                End If
                Set afoRegistro = Nothing
                
                If Not gBdGALILEO.bExecute(sSql) Then
                    
                    MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                    sSql = "DROP TABLE " & sTablaTemp
                    gBdGALILEO.bExecute (sSql)
                    Exit Sub
                    
                End If
                
                afoGastos.MoveNext
            
            Wend
        
        End If
        Set afoGastos = Nothing
        
        'Calculamos los gastos que de KILÓMETROS
        sSql = "SELECT g.Codigo,Nombre,NVL(d.Matricula,'') Matricula,SUM(Importe*Numero) AS Total " & _
               "FROM Gastos g,Directores d " & _
               "WHERE g.Codigo = d.Codigo AND " & _
               "g.Jerarquia = d.Jerarquia AND " & _
               "g.Jerarquia <> 'SU' AND " & _
               "g.CodPais = d.CodPais AND " & _
               "Campanya BETWEEN " & txtCampanyaInicio & " AND " & txtCampanyaFin & " AND " & _
               "g.TipoGasto = 'K' AND " & _
               "g.CodPais = " & CodPaisTrabajo & " AND " & _
               "Anyo = " & txtAnyo & " AND " & _
               "Importe <> 0 AND " & _
               "Validado = 'S' " & _
               " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,Nombre,d.Matricula " & _
               "ORDER BY g.Codigo"
        
        If gBdGALILEO.bRecordset(sSql, afoGastos) Then
        
            While Not afoGastos.EOF
            
                Set afoRegistro = Nothing
                sSql = "SELECT * FROM " & sTablaTemp & _
                       " WHERE Codigo = " & afoGastos("Codigo")
                If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
                
                    sSql = "UPDATE " & sTablaTemp & " SET TotalKG = "
                    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                        sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo))
                    Else
                        sSql = sSql & afoGastos("Total")
                    End If
                    sSql = sSql & " WHERE Codigo = " & afoGastos("Codigo")
                                        
                Else
            
                    sSql = "INSERT INTO " & sTablaTemp & "(Codigo,Nombre,Matricula,TotalKG,TotalTel,TotalResto) " & _
                           "VALUES(" & afoGastos("Codigo") & ",'" & _
                           Trim(afoGastos("Nombre")) & "','" & Trim(afoGastos("Matricula")) & "',"
                    
                    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                        sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo)) & ",0,0)"
                    Else
                        sSql = sSql & afoGastos("Total") & ",0,0)"
                    End If
                
                End If
                Set afoRegistro = Nothing
                
                If Not gBdGALILEO.bExecute(sSql) Then
                    
                    MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                    sSql = "DROP TABLE " & sTablaTemp
                    gBdGALILEO.bExecute (sSql)
                    Exit Sub
                    
                End If
                
                afoGastos.MoveNext
            
            Wend
        
        End If
        Set afoGastos = Nothing
        
        'Calculamos los gastos que de GASOLINA
        sSql = "SELECT g.Codigo,Nombre,NVL(d.Matricula,'') Matricula,SUM(Importe) AS Total " & _
               "FROM Gastos g,Directores d " & _
               "WHERE g.Codigo = d.Codigo AND " & _
               "g.Jerarquia = d.Jerarquia AND " & _
               "g.Jerarquia <> 'SU' AND " & _
               "g.CodPais = d.CodPais AND " & _
               "Campanya BETWEEN " & txtCampanyaInicio & " AND " & txtCampanyaFin & " AND " & _
               "g.TipoGasto = 'G' AND " & _
               "g.CodPais = " & CodPaisTrabajo & " AND " & _
               "Anyo = " & txtAnyo & " AND " & _
               "Importe <> 0 AND " & _
               "Validado = 'S' " & _
               " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,Nombre,d.Matricula " & _
               "ORDER BY g.Codigo"
        
        If gBdGALILEO.bRecordset(sSql, afoGastos) Then
        
            While Not afoGastos.EOF
            
                Set afoRegistro = Nothing
                sSql = "SELECT * FROM " & sTablaTemp & _
                       " WHERE Codigo = " & afoGastos("Codigo")
                If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
                
                    If afoRegistro("TotalKG") = 0 Then
                    
                        sSql = "UPDATE " & sTablaTemp & " SET TotalKG = "
                        If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                            sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo))
                        Else
                            sSql = sSql & afoGastos("Total")
                        End If
                    
                    Else
                    
                        sSql = "UPDATE " & sTablaTemp & " SET TotalKG = TotalKG + "
                        If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                            sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo))
                        Else
                            sSql = sSql & afoGastos("Total")
                        End If
                    
                    End If
                        
                    sSql = sSql & " WHERE Codigo = " & afoGastos("Codigo")
                                        
                Else
            
                    sSql = "INSERT INTO " & sTablaTemp & "(Codigo,Nombre,Matricula,TotalKG,TotalTel,TotalResto) " & _
                           "VALUES(" & afoGastos("Codigo") & ",'" & _
                           Trim(afoGastos("Nombre")) & "','" & Trim(afoGastos("Matricula")) & "',"
                    
                    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                        sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo)) & ",0,0)"
                    Else
                        sSql = sSql & afoGastos("Total") & ",0,0)"
                    End If
                
                End If
                Set afoRegistro = Nothing
                
                If Not gBdGALILEO.bExecute(sSql) Then
                    
                    MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                    sSql = "DROP TABLE " & sTablaTemp
                    gBdGALILEO.bExecute (sSql)
                    Exit Sub
                    
                End If
                
                afoGastos.MoveNext
            
            Wend
        
        End If
        Set afoGastos = Nothing
        
    Else
    'SUPERVISORAS
    
        'calculamos los gastos de TELÉFONO
        sSql = "SELECT g.Codigo,Nombre,SUM(Importe) AS Total " & _
               "FROM Gastos g,Supervisoras s " & _
               "WHERE g.Codigo = s.CodSuperv AND " & _
               "g.Jerarquia = 'SU' AND " & _
               "g.CodPais = s.CodPais AND " & _
               "Campanya BETWEEN " & txtCampanyaInicio & " AND " & txtCampanyaFin & " AND " & _
               "g.TipoGasto = 'T' AND " & _
               "g.CodPais = " & CodPaisTrabajo & " AND " & _
               "Anyo = " & txtAnyo & " AND " & _
               "Importe <> 0 AND " & _
               "Validado = 'S' " & _
               " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,Nombre " & _
               "ORDER BY g.Codigo"
        
        If gBdGALILEO.bRecordset(sSql, afoGastos) Then
        
            While Not afoGastos.EOF
            
                sSql = "INSERT INTO " & sTablaTemp & "(Codigo,Nombre,Matricula,TotalKG,TotalTel,TotalResto) " & _
                       "VALUES(" & afoGastos("Codigo") & ",'" & _
                       Trim(afoGastos("Nombre")) & "','',0,"
                
                If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                    sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo)) & ",0)"
                Else
                    sSql = sSql & afoGastos("Total") & ",0)"
                End If
                
                If Not gBdGALILEO.bExecute(sSql) Then
                    
                    MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                    sSql = "DROP TABLE " & sTablaTemp
                    gBdGALILEO.bExecute (sSql)
                    Exit Sub
                    
                End If
                
                afoGastos.MoveNext
            
            Wend
        
        End If
        Set afoGastos = Nothing
        
        'Calculamos los gastos que NO sean de TELÉFONO,GASOLINA O KILÓMETROS
        sSql = "SELECT g.Codigo,Nombre,SUM(Importe) AS Total " & _
               "FROM Gastos g,Supervisoras s " & _
               "WHERE g.Codigo = s.CodSuperv AND " & _
               "g.Jerarquia = 'SU' AND " & _
               "g.CodPais = s.CodPais AND " & _
               "Campanya BETWEEN " & txtCampanyaInicio & " AND " & txtCampanyaFin & " AND " & _
               "g.TipoGasto NOT IN ('T','K','G') AND " & _
               "g.CodPais = " & CodPaisTrabajo & " AND " & _
               "Anyo = " & txtAnyo & " AND " & _
               "Importe <> 0 AND " & _
               "Validado = 'S' " & _
               " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,Nombre " & _
               "ORDER BY g.Codigo"
        
        If gBdGALILEO.bRecordset(sSql, afoGastos) Then
        
            While Not afoGastos.EOF
            
                Set afoRegistro = Nothing
                sSql = "SELECT * FROM " & sTablaTemp & _
                       " WHERE Codigo = " & afoGastos("Codigo")
                If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
                
                    sSql = "UPDATE " & sTablaTemp & " SET TotalResto = "
                    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                        sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo))
                    Else
                        sSql = sSql & afoGastos("Total")
                    End If
                    sSql = sSql & " WHERE Codigo = " & afoGastos("Codigo")
                                        
                Else
            
                    sSql = "INSERT INTO " & sTablaTemp & "(Codigo,Nombre,Matricula,TotalKG,TotalTel,TotalResto) " & _
                           "VALUES(" & afoGastos("Codigo") & ",'" & _
                           Trim(afoGastos("Nombre")) & "','',0,0,"
                    
                    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                        sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo)) & ")"
                    Else
                        sSql = sSql & afoGastos("Total") & ")"
                    End If
                
                End If
                Set afoRegistro = Nothing
                
                If Not gBdGALILEO.bExecute(sSql) Then
                    
                    MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                    sSql = "DROP TABLE " & sTablaTemp
                    gBdGALILEO.bExecute (sSql)
                    Exit Sub
                    
                End If
                
                afoGastos.MoveNext
            
            Wend
        
        End If
        Set afoGastos = Nothing
        
        'Calculamos los gastos que de KILÓMETROS
        sSql = "SELECT g.Codigo,Nombre,SUM(Importe*Numero) AS Total " & _
               "FROM Gastos g,Supervisoras s " & _
               "WHERE g.Codigo = s.CodSuperv AND " & _
               "g.Jerarquia = 'SU' AND " & _
               "g.CodPais = s.CodPais AND " & _
               "Campanya BETWEEN " & txtCampanyaInicio & " AND " & txtCampanyaFin & " AND " & _
               "g.TipoGasto = 'K' AND " & _
               "g.CodPais = " & CodPaisTrabajo & " AND " & _
               "Anyo = " & txtAnyo & " AND " & _
               "Importe <> 0 AND " & _
               "Validado = 'S' " & _
               " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,Nombre " & _
               "ORDER BY g.Codigo"
        
        If gBdGALILEO.bRecordset(sSql, afoGastos) Then
        
            While Not afoGastos.EOF
            
                Set afoRegistro = Nothing
                sSql = "SELECT * FROM " & sTablaTemp & _
                       " WHERE Codigo = " & afoGastos("Codigo")
                If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
                
                    sSql = "UPDATE " & sTablaTemp & " SET TotalKG = "
                    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                        sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo))
                    Else
                        sSql = sSql & afoGastos("Total")
                    End If
                    sSql = sSql & " WHERE Codigo = " & afoGastos("Codigo")
                                        
                Else
            
                    sSql = "INSERT INTO " & sTablaTemp & "(Codigo,Nombre,Matricula,TotalKG,TotalTel,TotalResto) " & _
                           "VALUES(" & afoGastos("Codigo") & ",'" & _
                           Trim(afoGastos("Nombre")) & "','',"
                    
                    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                        sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo)) & ",0,0)"
                    Else
                        sSql = sSql & afoGastos("Total") & ",0,0)"
                    End If
                
                End If
                Set afoRegistro = Nothing
                
                If Not gBdGALILEO.bExecute(sSql) Then
                    
                    MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                    sSql = "DROP TABLE " & sTablaTemp
                    gBdGALILEO.bExecute (sSql)
                    Exit Sub
                    
                End If
                
                afoGastos.MoveNext
            
            Wend
        
        End If
        Set afoGastos = Nothing
        
        'Calculamos los gastos que de GASOLINA
        sSql = "SELECT g.Codigo,Nombre,SUM(Importe) AS Total " & _
               "FROM Gastos g,Supervisoras s " & _
               "WHERE g.Codigo = s.CodSuperv AND " & _
               "g.Jerarquia = 'SU' AND " & _
               "g.CodPais = s.CodPais AND " & _
               "Campanya BETWEEN " & txtCampanyaInicio & " AND " & txtCampanyaFin & " AND " & _
               "g.TipoGasto = 'G' AND " & _
               "g.CodPais = " & CodPaisTrabajo & " AND " & _
               "Anyo = " & txtAnyo & " AND " & _
               "Importe <> 0 AND " & _
               "Validado = 'S' " & _
               " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,Nombre " & _
               "ORDER BY g.Codigo"
        
        If gBdGALILEO.bRecordset(sSql, afoGastos) Then
        
            While Not afoGastos.EOF
            
                Set afoRegistro = Nothing
                sSql = "SELECT * FROM " & sTablaTemp & _
                       " WHERE Codigo = " & afoGastos("Codigo")
                If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
                
                    If afoRegistro("TotalKG") = 0 Then
                    
                        sSql = "UPDATE " & sTablaTemp & " SET TotalKG = "
                        If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                            sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo))
                        Else
                            sSql = sSql & afoGastos("Total")
                        End If
                    
                    Else
                    
                        sSql = "UPDATE " & sTablaTemp & " SET TotalKG = TotalKG + "
                        If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                            sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo))
                        Else
                            sSql = sSql & afoGastos("Total")
                        End If
                    
                    End If
                        
                    sSql = sSql & " WHERE Codigo = " & afoGastos("Codigo")
                                        
                Else
            
                    sSql = "INSERT INTO " & sTablaTemp & "(Codigo,Nombre,Matricula,TotalKG,TotalTel,TotalResto) " & _
                           "VALUES(" & afoGastos("Codigo") & ",'" & _
                           Trim(afoGastos("Nombre")) & "','',"
                    
                    If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                        sSql = sSql & dConvierteMoneda(CStr(afoGastos("Total")), 0, CByte(CodPaisTrabajo)) & ",0,0)"
                    Else
                        sSql = sSql & afoGastos("Total") & ",0,0)"
                    End If
                
                End If
                Set afoRegistro = Nothing
                
                If Not gBdGALILEO.bExecute(sSql) Then
                    
                    MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                    sSql = "DROP TABLE " & sTablaTemp
                    gBdGALILEO.bExecute (sSql)
                    Exit Sub
                    
                End If
                
                afoGastos.MoveNext
            
            Wend
        
        End If
        Set afoGastos = Nothing
        
    End If
    
    sSql = "SELECT * FROM " & sTablaTemp
    If gBdGALILEO.bRecordset(sSql, afoGastos) Then
        
        dtrInformeAgrupado.DataMember = ""
        Set dtrInformeAgrupado.DataSource = afoGastos
        
        dtrInformeAgrupado.Sections("EncabezadoPagina").Controls("etinicio").Caption = txtCampanyaInicio
        dtrInformeAgrupado.Sections("EncabezadoPagina").Controls("etfin").Caption = txtCampanyaFin
        dtrInformeAgrupado.Sections("EncabezadoPagina").Controls("etanyo").Caption = txtAnyo
        dtrInformeAgrupado.Sections("EncabezadoPagina").Controls("etFamiliaGasto").Caption = cmbFamiliaGasto.Text
        dtrInformeAgrupado.Sections("EncabezadoPagina").Controls("etMatricula").Caption = ""
        If CodPaisTrabajo = giPAIS_PORTUGAL Then dtrInformeAgrupado.Sections("EncabezadoPagina").Controls("etMatricula").Caption = "Matrícula"
        
        Screen.MousePointer = 0
          
        dtrInformeAgrupado.Show vbModal
        
    Else
        MsgBox "No existen datos entre esas campañas.", vbApplicationModal + vbExclamation, "Información"
    End If
    Set afoGastos = Nothing
    
    sSql = "DROP TABLE " & sTablaTemp
    gBdGALILEO.bExecute (sSql)
    
    Screen.MousePointer = 0

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub Form_Load()

    cmbFamiliaGasto.AddItem "Ordinario"
    cmbFamiliaGasto.AddItem "Formación"
    cmbFamiliaGasto.AddItem "Presentación"
    cmbFamiliaGasto.AddItem "Incentivo"
    cmbFamiliaGasto.Text = "Ordinario"
    
    cmbDeducible.AddItem "Sí"
    cmbDeducible.AddItem "No"
    cmbDeducible.ListIndex = 0
    
    lblEtiqueta(4).Visible = False
    cmbDeducible.Visible = False
    If CodPaisTrabajo = 52 Or CodPaisTrabajo = 56 Then
        lblEtiqueta(4).Visible = True
        cmbDeducible.Visible = True
    End If

End Sub

Private Sub optDirectores_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then cmdAceptar.SetFocus

End Sub

Private Sub optSupervisoras_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdAceptar.SetFocus
    
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

Private Sub txtCampanyaFin_GotFocus()
    TextGotFocus txtCampanyaFin
End Sub

Private Sub txtCampanyaFin_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, txtCampanyaFin, 2
End Sub

Private Sub txtCampanyaFin_LostFocus()
    TextLostFocus txtCampanyaFin
End Sub

Private Sub txtCampanyaInicio_GotFocus()
    TextGotFocus txtCampanyaInicio
End Sub

Private Sub txtCampanyaInicio_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, txtCampanyaInicio, 2
End Sub

Private Sub txtCampanyaInicio_LostFocus()
    TextLostFocus txtCampanyaInicio
End Sub
