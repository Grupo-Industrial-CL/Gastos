VERSION 5.00
Begin VB.Form frmInformeLitrosGasolina 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Litros Gasolina"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDeducible 
      Height          =   315
      ItemData        =   "frmInformeLitrosGasolina.frx":0000
      Left            =   3360
      List            =   "frmInformeLitrosGasolina.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   1000
   End
   Begin VB.ComboBox cmbFamiliaGasto 
      Height          =   315
      ItemData        =   "frmInformeLitrosGasolina.frx":0004
      Left            =   960
      List            =   "frmInformeLitrosGasolina.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox boxInicio 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   235
      Width           =   400
   End
   Begin VB.TextBox boxFin 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Top             =   235
      Width           =   400
   End
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "Volver"
      CausesValidation=   0   'False
      Height          =   675
      Left            =   2040
      Picture         =   "frmInformeLitrosGasolina.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton btAceptar 
      Caption         =   "Aceptar"
      Height          =   675
      Left            =   840
      Picture         =   "frmInformeLitrosGasolina.frx":0312
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox boxAnyo 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   600
   End
   Begin VB.Label lblDeducible 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deducible"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia Gasto"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label etCampanya 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Campaña Inicio"
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Campaña Fin"
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label etAnyo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Año"
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   600
   End
End
Attribute VB_Name = "frmInformeLitrosGasolina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btAceptar_Click()
    
    Dim rsGastos As ADODB.Recordset
    Dim rsGastosAux As ADODB.Recordset
    Dim Box
    Dim Media As Double
    Dim sTablaTemp As String
    
    On Error GoTo ErrorSub
    
    Set Box = Es_Vacio(boxInicio, boxFin, boxAnyo)
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
           "Litros Decimal(15,2)," & _
           "TotalLitros Decimal(15,2)," & _
           "Km Decimal(15,2)," & _
           "Media Decimal(15,2),Matricula Char(15))"
    
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al crear la tabla temporal.", vbApplicationModal + vbExclamation, "Error"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    sSql = "SELECT g.Codigo,Nombre,NVL(d.Matricula,'') Matricula,SUM(Numero) AS Litros,SUM(Importe) AS TotalLitros " & _
           "FROM Gastos g,Directores d " & _
           "WHERE g.Codigo = d.Codigo AND " & _
           "g.Jerarquia = d.Jerarquia AND " & _
           "g.Jerarquia IN ('DA','DP','DR','DN') AND " & _
           "g.CodPais = d.CodPais AND " & _
           "g.CodPais = " & CodPaisTrabajo & " AND " & _
           "Campanya BETWEEN " & boxInicio & " AND " & boxFin & " AND " & _
           "TipoGasto = 'G' AND " & _
           "Anyo = " & boxAnyo & " AND " & _
           "Importe <> 0 AND " & _
           "Validado = 'S' " & _
           "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
           "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
           "GROUP BY g.Codigo,Nombre,d.Matricula " & _
           "ORDER BY g.Codigo"
    
    If gBdGALILEO.bRecordset(sSql, rsGastos) Then
        
        While Not rsGastos.EOF
            
            sSql = "INSERT INTO " & sTablaTemp & "(Codigo,Nombre,Matricula,Litros,TotalLitros,Km,Media) " & _
                   "VALUES(" & rsGastos("Codigo") & ",'" & _
                   Trim(rsGastos("Nombre")) & "','" & Trim(rsGastos("Matricula")) & "'," & _
                   rsGastos("Litros") & ","
            
            If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                sSql = sSql & dConvierteMoneda(CStr(rsGastos("TotalLitros")), 0, CByte(CodPaisTrabajo)) & ",0,0)"
            Else
                sSql = sSql & rsGastos("TotalLitros") & ",0,0)"
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
        
        sSql = "SELECT g.Codigo,Nombre,NVL(d.Matricula,'') Matricula,SUM(Numero) AS KM " & _
               "FROM Gastos g,Directores d " & _
               "WHERE g.Codigo = d.Codigo AND " & _
               "g.Jerarquia = d.Jerarquia AND " & _
               "g.Jerarquia IN ('DA','DP','DR','DN') AND " & _
               "g.CodPais = d.CodPais AND " & _
               "g.CodPais = " & CodPaisTrabajo & " AND " & _
               "Campanya BETWEEN " & boxInicio & " AND " & boxFin & " AND " & _
               "TipoGasto = 'K' AND " & _
               "Anyo = " & boxAnyo & " AND " & _
               "Importe <> 0 AND " & _
               "Validado = 'S' " & _
               "AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
               "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' " & _
               "GROUP BY g.Codigo,Nombre,d.Matricula " & _
               "ORDER BY g.Codigo"
        
        If gBdGALILEO.bRecordset(sSql, rsGastos) Then
        
            While Not rsGastos.EOF
            
                Set rsGastosAux = Nothing
                sSql = "SELECT * FROM " & sTablaTemp & _
                       " WHERE Codigo = " & rsGastos("Codigo")
                If gBdGALILEO.bRecordset(sSql, rsGastosAux) Then
                    Media = (rsGastosAux("Litros") / rsGastos("KM")) * 100
                    sSql = "UPDATE " & sTablaTemp & " SET KM = " & rsGastos("KM") & _
                           ",Media = " & Media & _
                           " WHERE Codigo = " & rsGastos("Codigo")
                    
                    If Not gBdGALILEO.bExecute(sSql) Then
                        MsgBox "Error interno al obtener los datos.", vbApplicationModal + vbExclamation, "Error"
                        sSql = "DROP TABLE " & sTablaTemp
                        gBdGALILEO.bExecute (sSql)
                        Exit Sub
                    End If
                End If
                
                rsGastos.MoveNext
                
            Wend
            
        End If
    
    Else
        MsgBox "No existen datos entre esas campañas ", vbApplicationModal + vbExclamation
    End If
    
    sSql = "SELECT * FROM " & sTablaTemp
    If gBdGALILEO.bRecordset(sSql, rsGastos) Then
        
        InformeLitros.DataMember = ""
        Set InformeLitros.DataSource = rsGastos
        
        InformeLitros.Sections("EncabezadoInforme").Controls("lblPeriodo").Caption = "Entre campaña " & boxInicio & " y campaña " & boxFin & " de " & boxAnyo
        InformeLitros.Sections("EncabezadoInforme").Controls("etFamiliaGasto").Caption = "Familia gasto: " & cmbFamiliaGasto.Text
        InformeLitros.Sections("EncabezadoPagina").Controls("etMatricula").Caption = ""
        If CodPaisTrabajo = giPAIS_PORTUGAL Then InformeLitros.Sections("EncabezadoPagina").Controls("etMatricula").Caption = "Matrícula"
        Screen.MousePointer = 0
          
        InformeLitros.Show vbModal
    Else
        MsgBox "No existen datos entre esas campañas ", vbApplicationModal + vbExclamation
    End If
    
    sSql = "DROP TABLE " & sTablaTemp
    gBdGALILEO.bExecute (sSql)
    
    Set rsGastos = Nothing
    Set rsGastosAux = Nothing
    
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
