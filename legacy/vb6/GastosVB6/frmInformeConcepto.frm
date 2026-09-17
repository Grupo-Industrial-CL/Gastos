VERSION 5.00
Begin VB.Form frmInformeConcepto 
   Caption         =   "Informe por Concepto"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   4545
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbDeducible 
      Height          =   315
      ItemData        =   "frmInformeConcepto.frx":0000
      Left            =   2640
      List            =   "frmInformeConcepto.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Width           =   1000
   End
   Begin VB.ComboBox cmbFamiliaGasto 
      Height          =   315
      ItemData        =   "frmInformeConcepto.frx":0004
      Left            =   240
      List            =   "frmInformeConcepto.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.OptionButton optPts 
      Height          =   255
      Left            =   3720
      Picture         =   "frmInformeConcepto.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optEuro 
      Height          =   255
      Left            =   3720
      Picture         =   "frmInformeConcepto.frx":02AB
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optEscudos 
      Height          =   255
      Left            =   3720
      Picture         =   "frmInformeConcepto.frx":057D
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecciona Persona"
      Height          =   705
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   3315
      Begin VB.OptionButton optSuper 
         Caption         =   "Supervisoras"
         Height          =   195
         Left            =   1800
         TabIndex        =   8
         Top             =   280
         Width           =   1335
      End
      Begin VB.OptionButton optDirec 
         Caption         =   "Directores"
         Height          =   195
         Left            =   255
         TabIndex        =   7
         Top             =   285
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox boxAnyo 
      Height          =   285
      Left            =   825
      TabIndex        =   2
      Top             =   855
      Width           =   465
   End
   Begin VB.ComboBox cboxConcepto 
      Height          =   315
      ItemData        =   "frmInformeConcepto.frx":0818
      Left            =   2670
      List            =   "frmInformeConcepto.frx":081A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   885
      Width           =   1620
   End
   Begin VB.CommandButton btAceptar 
      Caption         =   "Aceptar"
      Height          =   420
      Left            =   1140
      TabIndex        =   13
      Top             =   3225
      Width           =   975
   End
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "C&errar"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   2655
      Picture         =   "frmInformeConcepto.frx":081C
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3045
      Width           =   855
   End
   Begin VB.TextBox boxFin 
      Height          =   315
      Left            =   3420
      TabIndex        =   1
      Top             =   180
      Width           =   380
   End
   Begin VB.TextBox boxInicio 
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   180
      Width           =   390
   End
   Begin VB.OptionButton optLiras 
      Height          =   255
      Left            =   3720
      Picture         =   "frmInformeConcepto.frx":0B26
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblDeducible 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deducible"
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia Gasto"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label etAnyo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Año"
      Height          =   255
      Left            =   255
      TabIndex        =   18
      Top             =   840
      Width           =   435
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto"
      Height          =   255
      Left            =   1785
      TabIndex        =   17
      Top             =   885
      Width           =   795
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Campaña Fin"
      Height          =   300
      Left            =   2265
      TabIndex        =   16
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label etCampanya 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Campaña Inicio"
      Height          =   300
      Left            =   270
      TabIndex        =   15
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmInformeConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub boxAnyo_GotFocus()
    TextGotFocus boxAnyo
End Sub

Private Sub boxAnyo_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, boxAnyo, 4
End Sub

Private Sub boxAnyo_LostFocus()
    TextLostFocus boxAnyo
End Sub

Private Sub boxInicio_GotFocus()
    TextGotFocus boxInicio
End Sub

Private Sub boxInicio_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, boxInicio, 2
End Sub

Private Sub boxInicio_LostFocus()
    TextLostFocus boxInicio
End Sub

Private Sub boxFin_GotFocus()
    TextGotFocus boxFin
End Sub
Private Sub boxFin_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, boxFin, 2
End Sub

Private Sub boxFin_LostFocus()
    TextLostFocus boxFin
End Sub

Private Sub btAceptar_Click()
    
    Dim rsGastos As ADODB.Recordset
    Dim Box
    Dim sTablaTemp As String
    
    On Error GoTo ErrorSub
    
    Set Box = Es_Vacio(boxInicio, boxFin, boxAnyo, cboxConcepto)
    If Not (Box Is Nothing) Then
        Box.SetFocus
        Exit Sub
    End If
    
    Set Box = Nothing
    
    If CInt(boxInicio) > CInt(boxFin) Then
        MsgBox "La campaña de inicio debe ser mayor que la campaña de fin", vbApplicationModal + vbCritical
        boxInicio.SetFocus
        Exit Sub
    End If
    
    If cmbFamiliaGasto.Text = "" Then
        cmbFamiliaGasto.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    Randomize
    sTablaTemp = "TG" & Int((1000000 * Rnd) + 1)
    sSql = "CREATE TEMP TABLE " & sTablaTemp & _
           "(Codigo Integer," & _
           "Nombre Char(50)," & _
           "Total Decimal(15,2),Matricula Char(15))"
    
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al crear la tabla temporal.", vbApplicationModal + vbExclamation, "Error"
        Exit Sub
    End If
    
    If optDirec Then
        If Trim(cboxConcepto) <> "Kilometros" Then
            sSql = "SELECT g.codigo,nombre,NVL(d.Matricula,'') Matricula,sum(importe) as total from " & _
                   " gastos g,tiposgasto t,directores d WHERE " & _
                   "g.codigo=d.codigo AND " & _
                   "g.jerarquia=d.jerarquia AND " & _
                   "g.jerarquia <> 'SU' AND " & _
                   "g.codpais=d.codpais AND " & _
                   "campanya between " & boxInicio & " AND " & boxFin & " AND " & _
                   "nombregasto = '" & cboxConcepto & "' AND " & _
                   "g.tipogasto = t.tipogasto AND t.Tipo = 'C' AND " & _
                   "g.codpais = " & CodPaisTrabajo & " AND " & _
                   "anyo = " & boxAnyo & " AND " & _
                   "importe <> 0 AND " & _
                   "Validado = 'S' " & _
                   "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
                   "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
                   "GROUP BY g.codigo,nombre,d.Matricula " & _
                   "ORDER BY g.codigo"
        Else
            sSql = "SELECT g.codigo,nombre,NVL(d.Matricula,'') Matricula,sum(importe*numero) as total from " & _
                   " gastos g,tiposgasto t,directores d WHERE " & _
                   "g.codigo=d.codigo AND " & _
                   "g.jerarquia=d.jerarquia AND " & _
                   "g.jerarquia <> 'SU' AND " & _
                   "g.codpais=d.codpais AND " & _
                   "campanya between " & boxInicio & " AND " & boxFin & " AND " & _
                   "nombregasto = '" & cboxConcepto & "' AND " & _
                   "g.tipogasto = t.tipogasto AND t.Tipo = 'C' AND " & _
                   "g.codpais = " & CodPaisTrabajo & " AND " & _
                   "anyo = " & boxAnyo & " AND " & _
                   "importe <> 0 AND " & _
                   "Validado = 'S' " & _
                   "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
                   "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
                   "GROUP BY g.codigo,nombre,d.Matricula " & _
                   "ORDER BY g.codigo"
        End If
    Else
    
        If Trim(cboxConcepto) <> "Kilometros" Then
            sSql = "SELECT g.codigo,nombre,'' Matricula, sum(importe) as total from " & _
                   " gastos g,tiposgasto t,supervisoras s WHERE " & _
                   "g.codigo=s.codsuperv AND " & _
                   "g.codpais=s.codpais AND " & _
                   "g.jerarquia = 'SU' AND " & _
                   "campanya between " & boxInicio & " AND " & boxFin & " AND " & _
                   "nombregasto = '" & cboxConcepto & "' AND " & _
                   "g.tipogasto = t.tipogasto AND t.Tipo = 'C' AND " & _
                   "g.codpais = " & CodPaisTrabajo & " AND " & _
                   "anyo = " & boxAnyo & " AND " & _
                   "importe <> 0 AND " & _
                   "Validado = 'S' " & _
                   "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
                   "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
                   "GROUP BY g.codigo,nombre " & _
                   "ORDER BY g.codigo"
        Else
            sSql = "SELECT g.codigo,nombre,'' Matricula, sum(importe*numero) as total from " & _
                   " gastos g,tiposgasto t,supervisoras s WHERE " & _
                   "g.codigo=s.codsuperv AND " & _
                   "g.codpais=s.codpais AND " & _
                   "g.jerarquia = 'SU' AND " & _
                   "campanya between " & boxInicio & " AND " & boxFin & " AND " & _
                   "nombregasto = '" & cboxConcepto & "' AND " & _
                   "g.tipogasto = t.tipogasto AND t.Tipo = 'C' AND " & _
                   "g.codpais = " & CodPaisTrabajo & " AND " & _
                   "anyo = " & boxAnyo & " AND " & _
                   "importe <> 0 AND " & _
                   "Validado = 'S' " & _
                   "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
                   "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
                   "GROUP BY g.codigo,nombre " & _
                   "ORDER BY g.codigo"
        End If
    End If
    
    If gBdGALILEO.bRecordset(sSql, rsGastos) Then
        
        While Not rsGastos.EOF
            
            sSql = "INSERT INTO " & sTablaTemp & "(Codigo,Nombre,Matricula,Total) " & _
                   "VALUES(" & rsGastos("Codigo") & ",'" & _
                   Trim(rsGastos("Nombre")) & "','" & Trim(rsGastos("Matricula")) & "',"
            
            If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    sSql = sSql & dConvierteMoneda(CStr(rsGastos("Total")), 0, CByte(CodPaisTrabajo)) & ")"
                Else
                    sSql = sSql & rsGastos("Total") & ")"
                End If
            Else
            
                sSql = sSql & rsGastos("Total") & ")"
            
            End If
            
            If Not gBdGALILEO.bExecute(sSql) Then
                
                MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                sSql = "DROP TABLE " & sTablaTemp
                gBdGALILEO.bExecute (sSql)
                Exit Sub
                
            End If
            
            rsGastos.MoveNext
            
        Wend
        
        Set rsGastos = Nothing
    
    Else
        MsgBox "No existen datos entre esas campañas ", vbApplicationModal + vbExclamation
    End If
    
    sSql = "SELECT * FROM " & sTablaTemp
    If gBdGALILEO.bRecordset(sSql, rsGastos) Then
        
        InformeConcepto.DataMember = ""
        Set InformeConcepto.DataSource = rsGastos
        
        InformeConcepto.Sections("CabezaInforme").Controls("etinicio").Caption = boxInicio
        InformeConcepto.Sections("CabezaInforme").Controls("etfin").Caption = boxFin
        InformeConcepto.Sections("CabezaInforme").Controls("etconcepto").Caption = Trim(cboxConcepto.Text) & " (" & Trim(cmbFamiliaGasto.Text) & ")"
        InformeConcepto.Sections("Sección2").Controls("etMatricula").Caption = ""
        If CodPaisTrabajo = giPAIS_PORTUGAL Then InformeConcepto.Sections("Sección2").Controls("etMatricula").Caption = "Matrícula"
        
        Screen.MousePointer = 0
          
        InformeConcepto.Show vbModal
    Else
        MsgBox "No existen datos entre esas campañas ", vbApplicationModal + vbExclamation
    End If
    
    sSql = "DROP TABLE " & sTablaTemp
    gBdGALILEO.bExecute (sSql)
    
    Set rsGastos = Nothing
    
    Screen.MousePointer = 0

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

Private Sub cboxConcepto_KeyPress(KeyAscii As Integer)
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
    
    Dim rsConcepto As ADODB.Recordset
    
    On Error GoTo ErrorSub
    
    sSql = "SELECT nombregasto FROM tiposgasto WHERE Tipo = 'C' " & _
           "ORDER BY nombregasto"
    
    If gBdGALILEO.bRecordset(sSql, rsConcepto) Then
        While Not rsConcepto.EOF
           cboxConcepto.AddItem (rsConcepto("nombregasto"))
           rsConcepto.MoveNext
        Wend
        cboxConcepto.ListIndex = 0
    Else
        MsgBox "Error al cargar los Conceptos", vbApplicationModal + vbCritical
    End If
    
    Set rsConcepto = Nothing

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

Private Sub optDirec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then btAceptar.SetFocus
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

Private Sub optSuper_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then btAceptar.SetFocus
End Sub
