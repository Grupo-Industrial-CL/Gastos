VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInformeAnual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Anual"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4935
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbDeducible 
      Height          =   315
      ItemData        =   "frmInformeAnual.frx":0000
      Left            =   3360
      List            =   "frmInformeAnual.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1440
      Width           =   1000
   End
   Begin VB.ComboBox cmbFamiliaGasto 
      Height          =   315
      ItemData        =   "frmInformeAnual.frx":0004
      Left            =   1200
      List            =   "frmInformeAnual.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.OptionButton optEscudos 
      Height          =   255
      Left            =   3000
      Picture         =   "frmInformeAnual.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optEuro 
      Height          =   255
      Left            =   3960
      Picture         =   "frmInformeAnual.frx":02A3
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optPts 
      Height          =   255
      Left            =   3000
      Picture         =   "frmInformeAnual.frx":0575
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optLiras 
      Height          =   255
      Left            =   3000
      Picture         =   "frmInformeAnual.frx":0818
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox boxCodigo 
      Height          =   285
      Left            =   885
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
   Begin VB.TextBox boxJerarquia 
      Height          =   285
      Left            =   2955
      TabIndex        =   1
      Top             =   195
      Width           =   375
   End
   Begin VB.CommandButton btAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1095
      TabIndex        =   9
      Top             =   2835
      Width           =   1215
   End
   Begin VB.TextBox boxAnyo 
      Height          =   285
      Left            =   165
      TabIndex        =   2
      Top             =   1440
      Width           =   570
   End
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "C&errar"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   3120
      Picture         =   "frmInformeAnual.frx":0A8A
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2640
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   300
      Left            =   1560
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   9568259
      CurrentDate     =   40415
   End
   Begin VB.Label lblMatricula 
      BackColor       =   &H008080FF&
      Caption         =   "Label3"
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblDeducible 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deducible"
      Height          =   255
      Left            =   3360
      TabIndex        =   20
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia Gasto"
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fecha del informe"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2205
      Width           =   1260
   End
   Begin VB.Label etCodigo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Código"
      Height          =   255
      Left            =   165
      TabIndex        =   16
      Top             =   180
      Width           =   615
   End
   Begin VB.Label etNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1005
      TabIndex        =   15
      Top             =   690
      Width           =   3135
   End
   Begin VB.Label etJerarquia 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jerarquía"
      Height          =   255
      Left            =   2085
      TabIndex        =   14
      Top             =   195
      Width           =   780
   End
   Begin VB.Label etAyuda 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1:DN  2:DR  3:DP  4:DA  5:SU"
      Height          =   465
      Left            =   3420
      TabIndex        =   13
      Top             =   105
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   255
      Left            =   165
      TabIndex        =   12
      Top             =   690
      Width           =   615
   End
   Begin VB.Label etAnyo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Año "
      Height          =   255
      Left            =   165
      TabIndex        =   11
      Top             =   1215
      Width           =   405
   End
End
Attribute VB_Name = "frmInformeAnual"
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
Dim rsGastos As ADODB.Recordset
Dim rsTabla As ADODB.Recordset
Dim rsCampanya As ADODB.Recordset
Dim Contador As Integer
Dim TotGaso As Double
Dim TotKm As Double
Dim TotBil As Double
Dim TotInv As Double
Dim TotHot As Double
Dim TotSal As Double
Dim TotTel As Double
Dim TotTra As Double
Dim TotPea As Double
Dim TotVar As Double
Dim TotCom As Double
Dim TotCen As Double
Dim TotCaf As Double
Dim TotPar As Double
Dim Total As Double
Dim Box
Dim Jerarquia As String
Dim rsFecha As ADODB.Recordset
Dim fecIni As Date
Dim sTablaTemp As String

Dim sFecIni As String
Dim sFecFin As String

On Error GoTo ErrorSub

Set Box = Es_Vacio(boxCodigo, etNombre, boxJerarquia, boxAnyo)

If Not (Box Is Nothing) Then
    Box.SetFocus
    Exit Sub

Else
  
Set Box = Nothing

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
 'gBdGALILEO.ModoCapturarErrores = False
 
    If cmbFamiliaGasto.Text = "" Then
        cmbFamiliaGasto.SetFocus
        Exit Sub
    End If
 
Randomize
sTablaTemp = "TG" & Int((1000000 * Rnd) + 1)

 sSql = "CREATE temp table " & sTablaTemp & _
    "(Campanya integer," & _
    "Gasolina decimal(15,2)," & _
    "Km integer," & _
    "Invitaciones decimal(15,2)," & _
    "Hotel decimal(15,2)," & _
    "Salones decimal(15,2)," & _
    "Telefono decimal(15,2)," & _
    "Billetes decimal(15,2)," & _
    "Transporte decimal(15,2)," & _
    "Peaje decimal(15,2)," & _
    "Varios decimal(15,2)," & _
    "Comida decimal(15,2)," & _
    "Cena decimal(15,2)," & _
    "Cafeteria decimal(15,2)," & _
    "Parking decimal(15,2)," & _
    "Total decimal(15,2))"
    
 If Not gBdGALILEO.bExecute(sSql) Then
    MsgBox "Error al crear la tabla temporal", vbApplicationModal + vbExclamation
    Exit Sub
 Else
 Me.MousePointer = 11
    sSql = "SELECT max(campanya) as maxcampanya FROM gastos " & _
      "WHERE codpais=" & CodPaisTrabajo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & " AND " & _
        "anyo=" & boxAnyo & " AND " & _
        "codigo=" & boxCodigo & " AND Jerarquia = '" & Jerarquia & "' AND " & _
        "Validado = 'S'"

  If Not gBdGALILEO.bRecordset(sSql, rsCampanya) Then
       Me.MousePointer = 0
       MsgBox "Error: No se ha podido realizar la operación correctamente", vbApplicationModal + vbExclamation
       sSql = "DROP TABLE " & sTablaTemp
       gBdGALILEO.bExecute sSql
       Exit Sub
  Else
     
     If IsNull(rsCampanya("maxcampanya")) Then
        Me.MousePointer = 0
        MsgBox "No hay gastos en el año seleccionado", vbApplicationModal + vbExclamation
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute sSql
        Exit Sub
     End If
     
     For Contador = 1 To rsCampanya("maxcampanya") Step 1
       sSql = "INSERT INTO " & sTablaTemp & " (Campanya ,Gasolina ,Km ,Invitaciones,Hotel,Salones,Telefono,Billetes,Transporte,Peaje,Varios,Comida,Cena,Cafeteria,Parking,Total) " & _
          "values( " & Contador & " ,'','','','','','','','','','','','','','','')"
      
      If Not gBdGALILEO.bExecute(sSql) Then
        Me.MousePointer = 0
        MsgBox "Error al insertar en la tabla temporal", vbApplicationModal + vbExclamation
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute sSql
        Exit Sub
      End If
     Next Contador

 
 For Contador = 1 To rsCampanya("maxcampanya") Step 1

'      sSql = "SELECT unique fecinicio,fecfin from campanyas WHERE " & _
'            "codpais=" & CodPaisTrabajo & " AND " & _
'            "anyo=" & boxAnyo & " AND " & _
'            "codcampanya=" & Contador
'
'      If Not gBdGALILEO.bRecordset(sSql, rsFecha) Then
'        MsgBox "Error al seleccionar las fechas", vbApplicationModal + vbCritical
'        sSql = "DROP TABLE " & sTablaTemp
'        gBdGALILEO.bExecute sSql
'      End If
     
    If (boxAnyo <> "") And (etNombre <> "") And (boxJerarquia <> "") Then
    
        If Not bDameFechasCampanya(CodPaisTrabajo, boxAnyo, CByte(Contador), boxCodigo, boxJerarquia, sFecIni, sFecFin) Then
            MsgBox "Error al seleccionar las fechas", vbApplicationModal + vbCritical
            sSql = "DROP TABLE " & sTablaTemp
            gBdGALILEO.bExecute sSql
            Exit Sub
        End If
    
    Else
        MsgBox "Error al seleccionar las fechas", vbApplicationModal + vbCritical
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute sSql
        Exit Sub
    End If

     
     'Case "G"
      sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
        "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
        "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
        "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
        "codpais=" & CodPaisTrabajo & " AND " & _
        "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
        "tipogasto = 'G' AND " & _
        "Validado = 'S'"

      If gBdGALILEO.bRecordset(sSql, rsGastos) Then
       If Not IsNull(rsGastos("cantidad")) Then
          TotGaso = CDbl(rsGastos("cantidad"))
          
          If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
            If optEuro Then
                TotGaso = dConvierteMoneda(CStr(TotGaso), 0, CByte(CodPaisTrabajo))
            End If
          End If
          
       Else
          TotGaso = 0
        End If
      End If 'fin de gasolina
             '   Case "k"
            
          sSql = "SELECT sum(numero*importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'K' AND " & _
            "Validado = 'S'"

        If gBdGALILEO.bRecordset(sSql, rsGastos) Then
         If Not IsNull(rsGastos("cantidad")) Then
             TotKm = CDbl(rsGastos("cantidad"))
             
             If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotKm = dConvierteMoneda(CStr(TotKm), 0, CByte(CodPaisTrabajo))
                End If
             End If
             
         Else
             TotKm = 0
         End If
        End If 'fin km
      
             '   Case "I"
               
           sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'I' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
            If Not IsNull(rsGastos("cantidad")) Then
              TotInv = CDbl(rsGastos("cantidad"))
              
              If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotInv = dConvierteMoneda(CStr(TotInv), 0, CByte(CodPaisTrabajo))
                End If
              End If

            Else
               TotInv = 0
            End If
         End If 'Invitaciones
       
          '   Case "H"
          sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'H' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
            If Not IsNull(rsGastos("cantidad")) Then
             TotHot = CDbl(rsGastos("cantidad"))
             
             If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotHot = dConvierteMoneda(CStr(TotHot), 0, CByte(CodPaisTrabajo))
                End If
             End If

            Else
             TotHot = 0
            End If
         End If 'Hotel
         
             '   Case "S"
                    
         sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'S' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
           If Not IsNull(rsGastos("cantidad")) Then
              TotSal = CDbl(rsGastos("cantidad"))
              
              If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotSal = dConvierteMoneda(CStr(TotSal), 0, CByte(CodPaisTrabajo))
                End If
              End If

           Else
              TotSal = 0
           End If
         End If 'Salones
        
             '   Case "F"Telefono
         sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'F' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
            If Not IsNull(rsGastos("cantidad")) Then
              TotTel = CDbl(rsGastos("cantidad"))
              
              If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotTel = dConvierteMoneda(CStr(TotTel), 0, CByte(CodPaisTrabajo))
                End If
              End If

            Else
              TotTel = 0
            End If
         End If 'Telefono
         
           '   Case "B"
                   
         sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'B' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
            If Not IsNull(rsGastos("cantidad")) Then
             TotBil = CDbl(rsGastos("cantidad"))
             
             If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotBil = dConvierteMoneda(CStr(TotBil), 0, CByte(CodPaisTrabajo))
                End If
             End If

            Else
                TotBil = 0
            End If
         End If 'Billetes
         
             '   Case "T"
                  
          sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto= 'T' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
            If Not IsNull(rsGastos("cantidad")) Then
              TotTra = CDbl(rsGastos("cantidad"))
              
              If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotTra = dConvierteMoneda(CStr(TotTra), 0, CByte(CodPaisTrabajo))
                End If
              End If

            Else
              TotTra = 0
            End If
         End If 'Transporte
        
           '   Case "P"
                    
          sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'P' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
            If Not IsNull(rsGastos("cantidad")) Then
              TotPea = CDbl(rsGastos("cantidad"))
              
              If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotPea = dConvierteMoneda(CStr(TotPea), 0, CByte(CodPaisTrabajo))
                End If
              End If

            Else
              TotPea = 0
            End If
         End If 'Peaje
             '   Case "V"
                    
          sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'V' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
            If Not IsNull(rsGastos("cantidad")) Then
             TotVar = CDbl(rsGastos("cantidad"))
             
             If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotVar = dConvierteMoneda(CStr(TotVar), 0, CByte(CodPaisTrabajo))
                End If
             End If

            Else
              TotVar = 0
            End If
         End If 'Varios
         
             '   Case "O"
               
           sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'O' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
            If Not IsNull(rsGastos("cantidad")) Then
              TotCom = CDbl(rsGastos("cantidad"))
              
              If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotCom = dConvierteMoneda(CStr(TotCom), 0, CByte(CodPaisTrabajo))
                End If
              End If

            Else
               TotCom = 0
            End If
         End If 'Comidas
         
             '   Case "E"
               
           sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'E' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
            If Not IsNull(rsGastos("cantidad")) Then
              TotCen = CDbl(rsGastos("cantidad"))
              
              If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotCen = dConvierteMoneda(CStr(TotCen), 0, CByte(CodPaisTrabajo))
                End If
              End If

            Else
               TotCen = 0
            End If
         End If 'Cenas
         
             '   Case "A"
               
           sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'A' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
            If Not IsNull(rsGastos("cantidad")) Then
              TotCaf = CDbl(rsGastos("cantidad"))
              
              If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotCaf = dConvierteMoneda(CStr(TotCaf), 0, CByte(CodPaisTrabajo))
                End If
              End If

            Else
               TotCaf = 0
            End If
         End If 'Cafeteria

             '   Case "X"
             
           sSql = "SELECT sum(importe) as cantidad FROM gastos WHERE " & _
            "codigo=" & boxCodigo & " AND jerarquia = '" & Jerarquia & "' AND " & _
            "anyo=" & boxAnyo & " AND NVL(FamiliaGasto,'O') = '" & Mid(cmbFamiliaGasto.Text, 1, 1) & "' " & _
            "AND NVL(Deducible,'S') = '" & Mid(cmbDeducible.Text, 1, 1) & "' AND " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "fecha between '" & sFecIni & "' AND '" & sFecFin & "' AND " & _
            "tipogasto = 'X' AND " & _
            "Validado = 'S'"

         If gBdGALILEO.bRecordset(sSql, rsGastos) Then
            If Not IsNull(rsGastos("cantidad")) Then
              TotPar = CDbl(rsGastos("cantidad"))
              
              If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
                If optEuro Then
                    TotPar = dConvierteMoneda(CStr(TotPar), 0, CByte(CodPaisTrabajo))
                End If
              End If

            Else
               TotPar = 0
            End If
         End If 'Parking



         
             Total = CDbl(TotGaso + TotKm + TotInv + TotHot + TotSal + TotTel + TotBil + TotTra + TotPea + TotVar + TotCom + TotCen + TotCaf + TotPar)
         
    sSql = "UPDATE " & sTablaTemp & " SET " & _
        "gasolina=" & TotGaso & "," & _
        "km=" & TotKm & "," & _
        "invitaciones=" & TotInv & "," & _
        "hotel=" & TotHot & "," & _
        "salones=" & TotSal & "," & _
        "telefono=" & TotTel & "," & _
        "billetes=" & TotBil & "," & _
        "transporte=" & TotTra & "," & _
        "peaje=" & TotPea & "," & _
        "varios=" & TotVar & "," & _
        "comida=" & TotCom & "," & _
        "cena=" & TotCen & "," & _
        "cafeteria=" & TotCaf & "," & _
        "parking=" & TotPar & "," & _
        "total=" & Total & " WHERE " & _
        "campanya=" & Contador
    
   If Not gBdGALILEO.bExecute(sSql) Then
    Me.MousePointer = 0
    MsgBox "Error: No se ha podido realizar la operación correctamente", vbApplicationModal + vbExclamation
    sSql = "DROP TABLE " & sTablaTemp
    gBdGALILEO.bExecute sSql
   End If
  
  Next Contador
  
    sSql = "SELECT campanya,gasolina,km,invitaciones,hotel," & _
        "salones,telefono,billetes,transporte,peaje,varios,comida,cena,cafeteria,parking,total" & _
        " FROM " & sTablaTemp & " "
        
   If gBdGALILEO.bRecordset(sSql, rsTabla) Then
   
        InformeAnual.Sections("Encabezado").Controls("etNombre").Caption = etNombre
        If CodPaisTrabajo = 34 And boxCodigo = "1" And boxJerarquia = "1" Then InformeAnual.Sections("Encabezado").Controls("etNombre").Caption = "JULIAN CARBALLO COBO"
        InformeAnual.Sections("Encabezado").Controls("etFecha").Caption = Format(dtpFecha, "dd/MM/yy")
        InformeAnual.Sections("Encabezado").Controls("etDirector").Caption = boxCodigo
        InformeAnual.Sections("Encabezado").Controls("etAnyo").Caption = boxAnyo
        InformeAnual.Sections("Encabezado").Controls("etFamiliaGasto").Caption = cmbFamiliaGasto.Text
        If Jerarquia = "SU" Then
            InformeAnual.Sections("Encabezado").Controls("etJerar").Caption = "Supervisora:"
        Else
            InformeAnual.Sections("Encabezado").Controls("etJerar").Caption = "Director:"
        End If
        InformeAnual.Sections("Encabezado").Controls("etMatricula").Caption = ""
        If CodPaisTrabajo = giPAIS_PORTUGAL Then InformeAnual.Sections("Encabezado").Controls("etMatricula").Caption = "Matrícula: " & lblMatricula
        InformeAnual.DataMember = ""
        Set InformeAnual.DataSource = rsTabla
        InformeAnual.Orientation = rptOrientLandscape
        InformeAnual.Show vbModal
        
   End If
        
        sSql = "DROP TABLE " & sTablaTemp
        gBdGALILEO.bExecute sSql
        
     Set rsGastos = Nothing
     Set rsCampanya = Nothing
    Set rsTabla = Nothing
 
End If 'campanyas
End If 'create
End If ' control de vacios

Me.MousePointer = 0
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

