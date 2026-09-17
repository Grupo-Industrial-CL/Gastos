VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIntroGastos 
   Caption         =   "Alta / Modificación de Gastos"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   14865
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame etTiposNoDeducible 
      Height          =   1065
      Left            =   12120
      TabIndex        =   66
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Label Label7 
         Caption         =   "0 ó N->No deducible   Vacio, 1 ó S->Deducible"
         ForeColor       =   &H000000FF&
         Height          =   660
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   1740
      End
   End
   Begin VB.TextBox boxDeducible 
      Height          =   300
      Left            =   13440
      TabIndex        =   16
      ToolTipText     =   "No Deducible por temas fiscales"
      Top             =   7635
      Width           =   555
   End
   Begin VB.TextBox boxImagen 
      BackColor       =   &H008080FF&
      Height          =   285
      Left            =   14640
      TabIndex        =   64
      Top             =   7560
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox boxPersonas 
      Height          =   300
      Left            =   11160
      TabIndex        =   15
      Top             =   7635
      Width           =   2295
   End
   Begin VB.TextBox boxProveedor 
      Height          =   300
      Left            =   9000
      MaxLength       =   50
      TabIndex        =   14
      Top             =   7635
      Width           =   2175
   End
   Begin VB.Frame etQueFamiliaGasto 
      Height          =   1065
      Left            =   75
      TabIndex        =   60
      Top             =   8000
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Label Label9 
         Caption         =   "O->Ordinario       F->Formación          P->Presentación    I ->Incentivo"
         ForeColor       =   &H000000FF&
         Height          =   780
         Left            =   135
         TabIndex        =   61
         Top             =   200
         Width           =   1380
      End
   End
   Begin VB.TextBox boxFamiliaGasto 
      Height          =   300
      Left            =   7980
      TabIndex        =   13
      Top             =   7635
      Width           =   1035
   End
   Begin VB.TextBox boxHora 
      Height          =   300
      Left            =   7440
      MaxLength       =   2
      TabIndex        =   12
      Top             =   7635
      Width           =   555
   End
   Begin VB.CommandButton cmdChequearLinea 
      Caption         =   "Chequear línea"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5760
      Picture         =   "frmIntroGastos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton cmdChequear 
      Caption         =   "Chequear todos"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4560
      Picture         =   "frmIntroGastos.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton btInicio 
      Caption         =   "Inicio"
      Height          =   450
      Left            =   8265
      TabIndex        =   39
      Top             =   645
      Width           =   645
   End
   Begin VB.TextBox boxContador 
      BackColor       =   &H008080FF&
      Height          =   285
      Left            =   0
      TabIndex        =   38
      Top             =   7275
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CommandButton btSubirDatos 
      Caption         =   "[^]"
      Height          =   330
      Left            =   14325
      TabIndex        =   17
      Top             =   7605
      Width           =   330
   End
   Begin VB.CommandButton btBorrarDatos 
      Caption         =   "X"
      Height          =   345
      Left            =   14325
      TabIndex        =   30
      Top             =   7215
      Width           =   315
   End
   Begin VB.TextBox boxNumero 
      Height          =   300
      Left            =   6885
      TabIndex        =   11
      Top             =   7635
      Width           =   555
   End
   Begin VB.TextBox boxImporte 
      Height          =   300
      Left            =   5730
      TabIndex        =   10
      Top             =   7635
      Width           =   1170
   End
   Begin VB.TextBox boxDonde 
      Height          =   300
      Left            =   2835
      TabIndex        =   9
      Top             =   7635
      Width           =   2895
   End
   Begin VB.TextBox boxTipoPago 
      Height          =   300
      Left            =   2010
      TabIndex        =   8
      Top             =   7635
      Width           =   810
   End
   Begin VB.TextBox boxTipoGasto 
      Height          =   300
      Left            =   30
      TabIndex        =   6
      Top             =   7635
      Width           =   975
   End
   Begin VB.TextBox boxFecha 
      Height          =   300
      Left            =   1020
      TabIndex        =   7
      Top             =   7635
      Width           =   975
   End
   Begin VB.CommandButton btOk 
      Caption         =   "OK"
      Height          =   450
      Left            =   8265
      TabIndex        =   4
      Top             =   120
      Width           =   645
   End
   Begin MSFlexGridLib.MSFlexGrid gridGastos 
      Height          =   5805
      Left            =   75
      TabIndex        =   5
      Top             =   1395
      Width           =   14670
      _ExtentX        =   25876
      _ExtentY        =   10239
      _Version        =   393216
   End
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "C&errar"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   6960
      Picture         =   "frmIntroGastos.frx":39F4
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8520
      Width           =   855
   End
   Begin VB.TextBox boxCampanya 
      Height          =   285
      Left            =   2070
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox boxJerarquia 
      Height          =   285
      Left            =   2955
      TabIndex        =   1
      Top             =   105
      Width           =   375
   End
   Begin VB.TextBox boxAnyo 
      Height          =   285
      Left            =   555
      TabIndex        =   2
      Top             =   690
      Width           =   495
   End
   Begin VB.TextBox boxCodigo 
      Height          =   285
      Left            =   870
      TabIndex        =   0
      Top             =   105
      Width           =   735
   End
   Begin VB.CommandButton cmdEuro 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Picture         =   "frmIntroGastos.frx":3CFE
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   8760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdEscudo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Picture         =   "frmIntroGastos.frx":3FD0
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   8760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdLira 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Picture         =   "frmIntroGastos.frx":426B
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   8760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdPts 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Picture         =   "frmIntroGastos.frx":44DD
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   8760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox boxImporte2 
      Height          =   210
      Left            =   6360
      TabIndex        =   55
      Top             =   7680
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Frame fraMoneda 
      Height          =   495
      Left            =   4560
      TabIndex        =   48
      Top             =   7920
      Width           =   2535
      Begin VB.OptionButton optEuro 
         Caption         =   "Euro"
         Height          =   195
         Left            =   1680
         TabIndex        =   50
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optLocal 
         Caption         =   "Moneda Local"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame etQueTipoGasto 
      Height          =   1305
      Left            =   75
      TabIndex        =   42
      Top             =   7920
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label Label6 
         Caption         =   "11 u O->Comida  12 ó E->Cena       13 ó A->Cafeteria  14 ó X->Parking"
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   2880
         TabIndex        =   57
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "2 ó G->Gasolina  4 ó K->Kms         6 ó P->Peaje       8 ó F->Telefono 10 ó V->Varios"
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   1500
         TabIndex        =   44
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "1 ó B->Billetes       3 ó H->Hotel          5 ó I->Invitaciones 7 ó S->Salones     9 ó T->Transporte "
         ForeColor       =   &H000000FF&
         Height          =   1020
         Left            =   135
         TabIndex        =   43
         Top             =   225
         Width           =   1380
      End
   End
   Begin VB.Label etQueTipoPago 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1 ó M->Metálico    2 ó V->Visa         3 ó S->Solred      4 ó X->Varios"
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   1560
      TabIndex        =   45
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label etNoDeducible 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deducible"
      Height          =   255
      Left            =   13440
      TabIndex        =   65
      ToolTipText     =   "No Deducible por temas fiscales"
      Top             =   7320
      Width           =   780
   End
   Begin VB.Image imgPortapapeles 
      Height          =   480
      Left            =   13800
      Picture         =   "frmIntroGastos.frx":4780
      ToolTipText     =   "Copiar datos al Portapapeles"
      Top             =   840
      Width           =   480
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   14280
      Picture         =   "frmIntroGastos.frx":504A
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgConsultar 
      Height          =   480
      Left            =   14280
      Picture         =   "frmIntroGastos.frx":5354
      ToolTipText     =   "Ver ticket asociado al gasto"
      Top             =   8040
      Width           =   480
   End
   Begin VB.Label etPersonas 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Personas"
      Height          =   255
      Left            =   11640
      TabIndex        =   63
      Top             =   7320
      Width           =   720
   End
   Begin VB.Label etProveedor 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   9480
      TabIndex        =   62
      Top             =   7320
      Width           =   795
   End
   Begin VB.Label etFamiliaGasto 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F. Gasto"
      Height          =   255
      Left            =   8160
      TabIndex        =   59
      Top             =   7290
      Width           =   660
   End
   Begin VB.Label etHora 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hora"
      Height          =   255
      Left            =   7515
      TabIndex        =   58
      Top             =   7290
      Width           =   405
   End
   Begin VB.Label etAyudaNum 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   7830
      TabIndex        =   47
      Top             =   8040
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label etAyuda 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1:DN  2:DR  3:DP  4:DA  5:SU"
      Height          =   270
      Left            =   2070
      TabIndex        =   46
      Top             =   420
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Label etTotalCamp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   6480
      TabIndex        =   41
      Top             =   705
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "    Total         Campaña"
      Height          =   435
      Left            =   5535
      TabIndex        =   40
      Top             =   615
      Width           =   825
   End
   Begin VB.Label etFechaAyuda 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha en Formato ""ddmmaaaa"""
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   540
      TabIndex        =   37
      Top             =   8265
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label etNumero 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Num"
      Height          =   255
      Left            =   6960
      TabIndex        =   36
      Top             =   7290
      Width           =   390
   End
   Begin VB.Label etImporte 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe"
      Height          =   270
      Left            =   5940
      TabIndex        =   35
      Top             =   7290
      Width           =   600
   End
   Begin VB.Label etDonde 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Donde"
      Height          =   255
      Left            =   4020
      TabIndex        =   34
      Top             =   7320
      Width           =   540
   End
   Begin VB.Label etTipoPago 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T.Pago"
      Height          =   240
      Left            =   2115
      TabIndex        =   33
      Top             =   7320
      Width           =   630
   End
   Begin VB.Label etTipoGasto 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T.Gasto"
      Height          =   255
      Left            =   225
      TabIndex        =   32
      Top             =   7305
      Width           =   645
   End
   Begin VB.Label etFecha 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      Height          =   255
      Left            =   1260
      TabIndex        =   31
      Top             =   7305
      Width           =   555
   End
   Begin VB.Label etAnyo1 
      Caption         =   "Año"
      Height          =   270
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   330
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   210
      Left            =   -1095
      TabIndex        =   28
      Top             =   3210
      Width           =   585
   End
   Begin VB.Label etAnyo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Año"
      Height          =   255
      Left            =   -525
      TabIndex        =   27
      Top             =   495
      Width           =   375
   End
   Begin VB.Label etFecFin 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3990
      TabIndex        =   26
      Top             =   720
      Width           =   945
   End
   Begin VB.Label etFecIni 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2700
      TabIndex        =   25
      Top             =   720
      Width           =   945
   End
   Begin VB.Label etCampanya 
      Caption         =   "Campaña"
      Height          =   255
      Left            =   1230
      TabIndex        =   24
      Top             =   720
      Width           =   735
   End
   Begin VB.Label etJerarquia 
      Caption         =   "Jerarquía"
      Height          =   255
      Left            =   2055
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.Label etCodigo 
      Caption         =   "Código"
      Height          =   255
      Left            =   165
      TabIndex        =   22
      Top             =   105
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   3675
      TabIndex        =   21
      Top             =   105
      Width           =   615
   End
   Begin VB.Label etNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmIntroGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PriGridGastos As New CFlexGrid
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

 Else

    If etNombre = "" Then
        
        MsgBox "No existe el código introducido o está dado de baja.", vbApplicationModal + vbInformation, "Información"
        btInicio_Click
        Exit Sub
        
    End If

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

Private Sub boxDonde_GotFocus()
TextGotFocus boxDonde
End Sub

Private Sub boxDonde_LostFocus()
TextLostFocus boxDonde
End Sub

Private Sub boxDonde_KeyPress(KeyAscii As Integer)
EntraTexto KeyAscii, boxDonde, 40, LetrasTodo
If KeyAscii = 27 Then
    Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
    PriGridGastos.NoResaltarLineas
    boxTipoGasto.SetFocus
End If
End Sub

Private Sub boxFecha_GotFocus()
etFechaAyuda.Visible = True
FechaGotFocus boxFecha
End Sub

Private Sub boxFecha_KeyPress(KeyAscii As Integer)
EntraTexto KeyAscii, boxFecha, 8, Letras09
If KeyAscii = 27 Then
    Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
    PriGridGastos.NoResaltarLineas
    boxTipoGasto.SetFocus
End If
End Sub

Private Sub boxFecha_LostFocus()
Dim sFechaBaja As String

etFechaAyuda.Visible = False
FechaLostFocus boxFecha

sFechaBaja = sObtenerFechaBaja

If Format(boxFecha, "yyyymmdd") > sFechaBaja Then

    MsgBox "La fecha de baja es anterior a la fecha introducida." & Chr(13) & "   No se pueden realizar cambios con esa fecha.", vbApplicationModal + vbInformation, "Información"
    boxFecha = ""
    boxFecha.SetFocus
    
End If

If Format(boxFecha, "yyyymmdd") > Format(Date, "yyyymmdd") Then
    
    MsgBox "La fecha del gasto no puede ser mayor que la fecha actual.", vbApplicationModal + vbInformation, "Información"
    boxFecha = ""
    boxFecha.SetFocus
    
End If

End Sub

Private Sub boxHora_GotFocus()
    TextGotFocus boxHora
End Sub

Private Sub boxHora_KeyPress(KeyAscii As Integer)

    EntraNumero KeyAscii, boxHora, 2
    If KeyAscii = 27 Then
        Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
        PriGridGastos.NoResaltarLineas
        boxTipoGasto.SetFocus
    End If

End Sub

Private Sub boxHora_LostFocus()
    TextLostFocus boxHora
End Sub

Private Sub boxImporte_GotFocus()
TextGotFocus boxImporte
End Sub

Private Sub boxImporte_KeyPress(KeyAscii As Integer)

If optEuro Then
    EntraNumero KeyAscii, boxImporte, 7, 2, "###0.00", "S"
Else
    EntraNumero KeyAscii, boxImporte, 7, , "###0.00", "S"
End If

If KeyAscii = 27 Then
    Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
    PriGridGastos.NoResaltarLineas
    boxTipoGasto.SetFocus
End If
End Sub

Private Sub boxImporte_LostFocus()
TextLostFocus boxImporte
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
    
 Else
 
    If etNombre = "" Then
        
        MsgBox "No existe el código introducido o está dado de baja.", vbApplicationModal + vbInformation, "Información"
        btInicio_Click
        Exit Sub
        
    End If
    
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


Private Sub boxNumero_GotFocus()
etAyudaNum.Visible = True
Select Case boxTipoGasto
    Case "Hotel"
        etAyudaNum.Caption = "Noches"
    Case "Invitaciones"
        etAyudaNum.Caption = "Personas"
    Case "Peaje"
        etAyudaNum.Caption = "Tickets"
    Case "Gasolina"
        etAyudaNum.Caption = "Litros"
    Case Else
        etAyudaNum.Visible = False
End Select

TextGotFocus boxNumero
End Sub

Private Sub boxNumero_KeyPress(KeyAscii As Integer)

If boxTipoGasto = "Gasolina" Then
    EntraNumero KeyAscii, boxNumero, 5, 2, "#0.00"
Else
    EntraNumero KeyAscii, boxNumero, 4
End If

If KeyAscii = 27 Then
    Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
    PriGridGastos.NoResaltarLineas
    boxTipoGasto.SetFocus
End If
End Sub

Private Sub boxNumero_LostFocus()
If etAyudaNum.Visible Then etAyudaNum.Visible = False
TextLostFocus boxNumero
End Sub

Private Sub boxPersonas_GotFocus()
    
    If Trim(boxTipoGasto.Text) <> "Invitaciones" And Trim(boxTipoGasto.Text) <> "Comida" And Trim(boxTipoGasto.Text) <> "Cena" And Trim(boxTipoGasto.Text) <> "Cafeteria" Then
        SendKeys "{TAB}"
        DoEvents
        Exit Sub
    End If
    If Not IsNumeric(Trim(boxNumero.Text)) Then
        SendKeys "{TAB}"
        DoEvents
        Exit Sub
    End If
    If CDbl(Trim(boxNumero.Text)) <= 1 Then
        SendKeys "{TAB}"
        DoEvents
        Exit Sub
    End If
    
    TextGotFocus boxPersonas
End Sub

Private Sub boxPersonas_KeyPress(KeyAscii As Integer)
    
    EntraTexto KeyAscii, boxPersonas, 40, LetrasTodo
    If KeyAscii = 27 Then
        Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
        PriGridGastos.NoResaltarLineas
        boxTipoGasto.SetFocus
    End If

End Sub

Private Sub boxPersonas_LostFocus()
    TextLostFocus boxPersonas
End Sub

Private Sub boxProveedor_GotFocus()
    
    If Trim(boxTipoPago.Text) <> "Varios" Then
        SendKeys "{TAB}"
        DoEvents
        Exit Sub
    End If
    
    TextGotFocus boxProveedor
End Sub

Private Sub boxProveedor_KeyPress(KeyAscii As Integer)
    
    EntraTexto KeyAscii, boxProveedor, 40, LetrasTodo
    If KeyAscii = 27 Then
        Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
        PriGridGastos.NoResaltarLineas
        boxTipoGasto.SetFocus
    End If

End Sub

Private Sub boxProveedor_LostFocus()
    TextLostFocus boxProveedor
End Sub

Private Sub boxTipoGasto_GotFocus()

boxDonde.Enabled = True
boxNumero.Enabled = True

Select Case boxTipoGasto
    Case "Gasolina"
        'boxNumero.Enabled = False
    Case "Salones"
        boxNumero.Enabled = False
    Case "Telefono"
        boxDonde.Enabled = False
        boxNumero.Enabled = False
    Case "Transporte"
        boxNumero.Enabled = False
End Select

If boxContador = "" Then _
    Limpia_Boxes boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
etQueTipoGasto.Visible = True

TextGotFocus boxTipoGasto
End Sub

Private Sub boxTipoGasto_KeyPress(KeyAscii As Integer)
EntraTexto KeyAscii, boxTipoGasto, 2, Letras09 & "bBgGhHiIkKpPsSfFtTvVoOeEaAxX", True
If KeyAscii = 27 Then
    Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
    PriGridGastos.NoResaltarLineas
End If
End Sub

Private Sub boxTipoGasto_LostFocus()
etQueTipoGasto.Visible = False
TextLostFocus boxTipoGasto

If IsNumeric(boxTipoGasto) Then
 If Val(boxTipoGasto) < 1 Or Val(boxTipoGasto) > 14 Then Exit Sub
Else
 If Len(boxTipoGasto) <> 1 Then Exit Sub
End If

boxNumero.Enabled = True
boxDonde.Enabled = True

Select Case boxTipoGasto
    Case "1", "B"
        boxTipoGasto = "Billetes"
    Case "2", "G"
        boxTipoGasto = "Gasolina"
        'boxNumero.Enabled = False
    Case "3", "H"
        boxTipoGasto = "Hotel"
    Case "4", "K"
        boxTipoGasto = "Kilometros"
    Case "5", "I"
        boxTipoGasto = "Invitaciones"
    Case "6", "P"
        boxTipoGasto = "Peaje"
    Case "7", "S"
        boxTipoGasto = "Salones"
        boxNumero.Enabled = False
    Case "8", "F"
        boxTipoGasto = "Telefono"
        boxDonde.Enabled = False
        boxNumero.Enabled = False
    Case "9", "T"
        boxTipoGasto = "Transporte"
        boxNumero.Enabled = False
    Case "10", "V"
        boxTipoGasto = "Varios"
    Case "11", "O"
        boxTipoGasto = "Comida"
    Case "12", "E"
        boxTipoGasto = "Cena"
    Case "13", "A"
        boxTipoGasto = "Cafeteria"
    Case "14", "X"
        boxTipoGasto = "Parking"
End Select
End Sub

Private Sub boxTipoPago_GotFocus()
etQueTipoPago.Visible = True
TextGotFocus boxTipoPago
Select Case boxTipoPago
    Case "Metalico"
        boxTipoPago = 1
    Case "Visa"
        boxTipoPago = 2
    Case "Solred"
        boxTipoPago = 3
    Case "Varios"
        boxTipoPago = 4
    Case Else
        boxTipoPago = ""
End Select
End Sub

Private Sub boxTipoPago_KeyPress(KeyAscii As Integer)
EntraTexto KeyAscii, boxTipoPago, 1, "1234mMvVsSxX", True
If KeyAscii = 27 Then
    Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
    PriGridGastos.NoResaltarLineas
    boxTipoGasto.SetFocus
End If
End Sub

Private Sub boxTipoPago_LostFocus()
etQueTipoPago.Visible = False
TextLostFocus boxTipoPago
Select Case boxTipoPago
    Case "1", "M"
        boxTipoPago = "Metalico"
    Case "2", "V"
        boxTipoPago = "Visa"
    Case "3", "S"
        boxTipoPago = "Solred"
    Case "4", "X"
        boxTipoPago = "Varios"
    Case Else
        boxTipoPago = ""
End Select

End Sub

Private Sub boxDeducible_GotFocus()
etTiposNoDeducible.Visible = True
TextGotFocus boxDeducible
Select Case boxDeducible
    Case "No"
        boxDeducible = 0
    Case Else
        boxDeducible = 1
End Select
boxDeducible.SelStart = 0
boxDeducible.SelLength = Len(boxDeducible.Text)
End Sub

Private Sub boxDeducible_KeyPress(KeyAscii As Integer)
EntraTexto KeyAscii, boxDeducible, 1, "01sSnN", True
If KeyAscii = 27 Then
    Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
    PriGridGastos.NoResaltarLineas
    boxTipoGasto.SetFocus
End If
End Sub

Private Sub boxDeducible_LostFocus()
etTiposNoDeducible.Visible = False
TextLostFocus boxDeducible
Select Case boxDeducible
    Case "0", "N", "n"
        boxDeducible = "No"
    Case Else
        boxDeducible = "Sí"
End Select

End Sub

Private Sub btBorrarDatos_Click()
Dim Contador As Long
Dim Linea
Dim rsCamp As ADODB.Recordset

On Error GoTo ErrorSub

If PriGridGastos.LineaSeleccionada <> -1 Then
    
    If MsgBox("¿Realmente desea eliminar la linea selecionada?", vbApplicationModal + vbDefaultButton1 + vbYesNo) = vbNo Then Exit Sub
    
    Contador = Val(boxContador)
    
    sSql = "DELETE FROM Gastos WHERE " & _
    "contador=" & Contador

    If gBdGALILEO.bExecute(sSql) Then
        Linea = PriGridGastos.ObtenerTextoFila(PriGridGastos.LineaSeleccionada)
        PriGridGastos.Vaciar
        ActualizarGrid
        
        sSql = "SELECT * from Campgastos WHERE " & _
            "codpais =" & CodPaisTrabajo & " AND " & _
            "anyo=" & boxAnyo & " AND " & _
            "campanya=" & boxCampanya
        
        If Not gBdGALILEO.bRecordset(sSql, rsCamp) Then
            
            sSql = "INSERT INTO campgastos (codpais,anyo,campanya) VALUES (" & _
                CodPaisTrabajo & ", " & _
                boxAnyo & "," & _
                boxCampanya & ")"
                
            If Not gBdGALILEO.bExecute(sSql) Then _
                MsgBox "Error al introducir la campaña. Consulte al administrador", vbApplicationModal + vbExclamation
        End If
        
    Else
        MsgBox "Error al eliminar el registro", vbApplicationModal + vbCritical
    End If
End If

boxContador = ""
boxTipoGasto.SetFocus

ExitSub:
    Exit Sub
ErrorSub:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
        End If
    End With
    GoTo ExitSub
    Resume Next
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

'    sSql = "SELECT fecinicio,fecfin FROM campanyas " & _
'        " WHERE codpais = " & CodPaisTrabajo & " AND " & _
'        " codcampanya = " & boxCampanya & " AND " & _
'        "anyo = " & boxAnyo
'
'        Select Case boxJerarquia
'        Case 1
'            sSql = sSql & _
'                  " AND CodDirVtas = (SELECT CodGeografia FROM Directores " & _
'                                     "WHERE CodPais = " & CodPaisTrabajo & _
'                                     " AND Codigo = " & boxCodigo & _
'                                     " AND Jerarquia = 'DN')"
'
'        Case 2
'            sSql = sSql & _
'                  " AND CodRegion = (SELECT CodGeografia FROM Directores " & _
'                                    "WHERE CodPais = " & CodPaisTrabajo & _
'                                    " AND Codigo = " & boxCodigo & _
'                                    " AND Jerarquia = 'DR')"
'        Case 3
'            sSql = sSql & _
'                  " AND CodRegion = (SELECT CodRegion FROM Areas " & _
'                                    "WHERE CodPais = " & CodPaisTrabajo & _
'                                    "AND CodArea = (SELECT MIN(CodArea) FROM DirApoyo " & _
'                                                   "WHERE CodPais = " & CodPaisTrabajo & _
'                                                   " AND Codigo = " & boxCodigo & "))"
'        Case 4
'            sSql = sSql & _
'                  " AND CodRegion = (SELECT CodRegion FROM Areas " & _
'                                    "WHERE CodPais = " & CodPaisTrabajo & _
'                                    "AND CodArea = (SELECT CodGeografia FROM Directores " & _
'                                                   "WHERE CodPais = " & CodPaisTrabajo & _
'                                                   " AND Codigo = " & boxCodigo & _
'                                                   " AND Jerarquia = 'DA'))"
'
'        Case 5
'            sSql = sSql & _
'                  " AND CodRegion = (SELECT CodRegion FROM Supervisoras " & _
'                                    "WHERE CodPais = " & CodPaisTrabajo & _
'                                    " AND CodSuperv = " & boxCodigo & ")"
'        Case Else
'
'        End Select
'
'        '& " AND FecFin <= '" & _
'        Format(Date, "ddmmyyyy") & "'"
'
'    If gBdGALILEO.bRecordset(sSql, rsCampanyas) Then
'      etFecIni = rsCampanyas("fecinicio")
'      etFecFin = rsCampanyas("fecfin")
'    End If
'End If
'
'Set rsCampanyas = Nothing

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

If boxAnyo <> "" Then
    If CInt(boxAnyo) > Year(Date) Then
        boxAnyo = ""
        boxAnyo.SetFocus
        Exit Sub
    End If
End If

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
'        "anyo = " & boxAnyo '& " AND FecFin <= '" & _
'        Format(Date, "ddmmyyyy") & "'"
'
'    If gBdGALILEO.bRecordset(sSql, rsCampanyas) Then
'        etFecIni = rsCampanyas("fecinicio")
'        etFecFin = rsCampanyas("fecfin")
'    End If
'End If
'
'Set rsCampanyas = Nothing

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


Private Sub btInicio_Click()
PriGridGastos.Vaciar
etTotalCamp = ""
boxCodigo.Enabled = True
boxJerarquia.Enabled = True
boxAnyo.Enabled = True
boxCampanya.Enabled = True
btOk.Enabled = True

DesactivarCampos

Limpia_Boxes boxCodigo, boxJerarquia, etNombre, boxAnyo, boxCampanya, etFecIni, etFecFin
Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
boxCodigo.SetFocus
End Sub

Private Sub btOk_Click()
Dim Box

On Error GoTo ErrorSub
 
Set Box = Es_Vacio(boxCodigo, etNombre, boxJerarquia, boxAnyo, etFecIni, etFecFin, boxCampanya)
If Not (Box Is Nothing) Then
    Box.SetFocus
    Exit Sub
End If

Set Box = Nothing

ActivarCampos

ActualizarGrid

If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then cmdEuro_Click

boxCodigo.Enabled = False
boxJerarquia.Enabled = False
boxAnyo.Enabled = False
boxCampanya.Enabled = False
btOk.Enabled = False

ExitSub:
    Exit Sub
ErrorSub:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
            Resume
        End If
    End With
    GoTo ExitSub
    Resume Next
End Sub

Private Sub btSubirDatos_Click()
Dim Box
Dim Jerarquia, TipoPago, TipoGasto, Fecha As String
Dim rsCamp As ADODB.Recordset
Dim iHora As Integer
Dim sFamiliaGasto As String
Dim sDeducible As String

On Error GoTo ErrorSub

TipoGasto = ""

boxDonde.Text = Trim(boxDonde.Text)
boxDonde.Text = Replace(boxDonde.Text, "'", "`")
boxDonde.Text = Replace(boxDonde.Text, """", "`")
boxPersonas.Text = Trim(boxPersonas.Text)
boxPersonas.Text = Replace(boxPersonas.Text, "'", "`")
boxPersonas.Text = Replace(boxPersonas.Text, """", "`")
boxProveedor.Text = Trim(boxProveedor.Text)
boxProveedor.Text = Replace(boxProveedor.Text, "'", "`")
boxProveedor.Text = Replace(boxProveedor.Text, """", "`")

Select Case boxTipoGasto
    Case "Billetes"
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero)
        TipoGasto = "B"
    Case "Gasolina"
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero, boxHora)
        'boxNumero = 0
        TipoGasto = "G"
    Case "Hotel"
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero)
        TipoGasto = "H"
    Case "Invitaciones"
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero, boxHora)
        TipoGasto = "I"
    Case "Kilometros"
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero)
        TipoGasto = "K"
    Case "Peaje"
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero, boxHora)
        TipoGasto = "P"
    Case "Salones"
        boxNumero = 0
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte)
        TipoGasto = "S"
    Case "Telefono"
        boxNumero = 0: boxDonde = ""
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxImporte)
        TipoGasto = "F"
    Case "Transporte"
        boxNumero = 0
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte)
        TipoGasto = "T"
    Case "Varios"
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero)
        TipoGasto = "V"
    Case "Comida"
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero, boxHora)
        TipoGasto = "O"
    Case "Cena"
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero, boxHora)
        TipoGasto = "E"
    Case "Cafeteria"
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero, boxHora)
        TipoGasto = "A"
    Case "Parking"
        Set Box = Es_Vacio(boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero, boxHora)
        TipoGasto = "X"
    Case Else
        boxTipoGasto.SetFocus
        Exit Sub
End Select

    If Not (Box Is Nothing) Then
        Box.SetFocus
        Exit Sub
    End If

    Set Box = Nothing
    
    iHora = -1
    If boxHora <> "" Then
        If boxHora < 1 Or boxHora > 24 Then
            boxHora = ""
            boxHora.SetFocus
            Exit Sub
        End If
        iHora = boxHora
    End If
    
    If boxFamiliaGasto = "" Then
        boxFamiliaGasto.SetFocus
        Exit Sub
    Else
        Select Case boxFamiliaGasto
        Case "Ordinario"
            sFamiliaGasto = "O"
        Case "Formación"
            sFamiliaGasto = "F"
        Case "Presentación"
            sFamiliaGasto = "P"
        Case "Incentivo"
            sFamiliaGasto = "I"
        Case Else
            boxFamiliaGasto.SetFocus
            Exit Sub
        End Select
    End If

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

    Select Case boxTipoPago
        Case "Metalico"
            TipoPago = "M"
        Case "Visa"
            TipoPago = "V"
        Case "Solred"
            TipoPago = "S"
        Case "Varios"
            TipoPago = "X"
    End Select

If IsNumeric(boxTipoGasto) Or TipoGasto = "" Then
    boxTipoGasto.SetFocus
    Exit Sub
End If

    If TipoPago = "X" Then
        boxProveedor.Text = Trim(boxProveedor.Text)
        If boxProveedor.Text = "" Then
            boxProveedor.SetFocus
            Exit Sub
        End If
    Else
        boxProveedor.Text = ""
    End If
    If (TipoGasto = "I" Or TipoGasto = "O" Or TipoGasto = "E" Or TipoGasto = "A") Then
        If boxNumero.Text > 1 Then
            boxPersonas.Text = Trim(boxPersonas.Text)
            If boxPersonas.Text = "" Then
                boxPersonas.SetFocus
                Exit Sub
            End If
        Else
            boxPersonas.Text = ""
        End If
    Else
        boxPersonas.Text = ""
    End If
    
    sDeducible = "NULL"
    Select Case boxDeducible.Text
    Case "Sí", "", "S", "s", "1"
        sDeducible = "'S'"
    Case Else
        sDeducible = "'N'"
    End Select

Fecha = Mid(boxFecha, 1, 2) + Mid(boxFecha, 4, 2) + Mid(boxFecha, 7, 4)

If boxContador = "" Then 'Añadir Gasto
    'MsgBox "insertar"
        
    sSql = "INSERT INTO gastos (codpais,codigo,jerarquia,tipogasto,fecha,anyo,campanya," & _
     "importe,numero,donde,tipopago,identificador,chequeado,validado,validacion,pagado,hora,familiagasto,Proveedor,Personas,Deducible) VALUES (" & _
     CodPaisTrabajo & ", " & _
     boxCodigo & "," & _
     "'" & Jerarquia & "'," & _
     "'" & TipoGasto & "'," & _
     "'" & Fecha & "'," & _
     boxAnyo & "," & _
     boxCampanya & ","
     
    If (CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51) And _
    optEuro Then
        sSql = sSql & dConvierteMoneda(boxImporte, 1, CByte(CodPaisTrabajo)) & ","
    Else
        sSql = sSql & boxImporte & ","
    End If
    
    sSql = sSql & boxNumero & "," & _
     "'" & boxDonde & "'," & _
     "'" & TipoPago & "'," & _
     "0,'S','S','Aplicacion Gastos','N'," & iHora & ",'" & sFamiliaGasto & "','" & boxProveedor.Text & "','" & boxPersonas.Text & "'," & sDeducible & ")"
    
    If gBdGALILEO.bExecute(sSql) Then
        PriGridGastos.Vaciar
        ActualizarGrid
        
        sSql = "SELECT * from Campgastos WHERE " & _
            "codpais =" & CodPaisTrabajo & " AND " & _
            "anyo=" & boxAnyo & " AND " & _
            "campanya=" & boxCampanya
        
        If Not gBdGALILEO.bRecordset(sSql, rsCamp) Then
            
            sSql = "INSERT INTO campgastos (codpais,anyo,campanya) VALUES (" & _
                CodPaisTrabajo & ", " & _
                boxAnyo & "," & _
                boxCampanya & ")"
                
            If Not gBdGALILEO.bExecute(sSql) Then _
                MsgBox "Error al introducir la campaña. Consulte al administrador", vbApplicationModal + vbExclamation
        End If
        
    Else
        MsgBox "Error al introducir el gasto", vbApplicationModal + vbCritical
    End If

Else 'Actualizar Gasto
    'MsgBox "actualizar"
    
    sSql = "UPDATE gastos SET " & _
        " tipogasto ='" & TipoGasto & "'," & _
        " fecha = '" & Fecha & "'," & _
        " tipopago='" & TipoPago & "'," & _
        " donde = '" & Trim(boxDonde) & "'," & _
        " hora = " & iHora & "," & _
        " familiagasto = '" & sFamiliaGasto & "'," & _
        " Proveedor = '" & boxProveedor.Text & "'," & _
        " Personas = '" & boxPersonas.Text & "'," & _
        " Deducible = " & sDeducible & ","
    
    If (CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51) And _
    optEuro Then
        sSql = sSql & " importe = " & dConvierteMoneda(boxImporte, 1, CByte(CodPaisTrabajo)) & ","
    Else
        sSql = sSql & " importe = " & boxImporte & ","
    End If
    
    sSql = sSql & " numero = " & boxNumero & _
        " WHERE contador = " & boxContador
    
    If gBdGALILEO.bExecute(sSql) Then
        PriGridGastos.Vaciar
        ActualizarGrid
        
        sSql = "SELECT * from Campgastos WHERE " & _
            "codpais =" & CodPaisTrabajo & " AND " & _
            "anyo=" & boxAnyo & " AND " & _
            "campanya=" & boxCampanya
        
        If Not gBdGALILEO.bRecordset(sSql, rsCamp) Then
            
            sSql = "INSERT INTO campgastos (codpais,anyo,campanya) VALUES (" & _
                CodPaisTrabajo & ", " & _
                boxAnyo & "," & _
                boxCampanya & ")"
                
            If Not gBdGALILEO.bExecute(sSql) Then _
                MsgBox "Error al introducir la campaña. Consulte al administrador", vbApplicationModal + vbExclamation
        End If
    Else
        MsgBox "Error al actualizar el gasto", vbApplicationModal + vbCritical
    End If
    
End If

boxContador = ""
boxTipoGasto.SetFocus

ExitSub:
    Exit Sub
ErrorSub:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
        End If
    End With
    GoTo ExitSub
    Resume Next
End Sub



Private Sub cmdChequear_Click()

    Dim sJerarquia As String
    
    Select Case boxJerarquia
    Case "1"
        sJerarquia = "DN"
    Case "2"
        sJerarquia = "DR"
    Case "3"
        sJerarquia = "DP"
    Case "4"
        sJerarquia = "DA"
    Case "5"
        sJerarquia = "SU"
    Case Else
        Exit Sub
    End Select
    
    gBdGALILEO.BeginTrans False
    
    sSql = "UPDATE Gastos SET Chequeado = 'S' " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " AND Anyo = " & boxAnyo & _
           " AND Campanya = " & boxCampanya & _
           " AND Codigo = " & boxCodigo & _
           " AND Jerarquia = '" & sJerarquia & _
           "' AND Identificador <> '0' " & _
           "AND Validado = 'S'"
           
    If gBdGALILEO.bExecute(sSql) Then
    
        gBdGALILEO.CommitTrans
        MsgBox "Gastos chequeados correctamente.", vbApplicationModal + vbInformation, "Información"
        btInicio_Click
    
    Else
    
        gBdGALILEO.RollbackTrans
        MsgBox "Error al chequear los gastos.", vbApplicationModal + vbExclamation, "Error"
    
    End If

End Sub

Private Sub cmdChequearLinea_Click()

    Dim sSql As String

    On Error GoTo Error_Rutina
    
    If boxContador = "" Then
        Exit Sub
    End If
    
    sSql = "UPDATE Gastos SET Chequeado = 'S' " & _
           "WHERE Contador = " & boxContador & _
           " AND Identificador <> '0' " & _
           "AND Validado = 'S'"
    
    If gBdGALILEO.bExecute(sSql) Then
        
        PriGridGastos.Vaciar
        ActualizarGrid
    
    Else
        MsgBox "Error al actualizar el gasto", vbApplicationModal + vbInformation, "Información"
    End If
    
    Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
    PriGridGastos.NoResaltarLineas
    boxTipoGasto.SetFocus
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdEscudo_Click()
    cmdEscudo.Visible = False
    cmdEuro.Visible = True
    PriGridGastos.CambioColumnaMonedaLocal 5, CodPaisTrabajo, 2
    'etTotalCamp = dConvierteMoneda(CStr(etTotalCamp), 1, CByte(CodPaisTrabajo))
    etTotalCamp = dTotalCampanya
End Sub

Private Sub cmdEuro_Click()

    cmdEuro.Visible = False
    Select Case CodPaisTrabajo
    Case 34
        cmdPts.Visible = True
    Case 39
        cmdLira.Visible = True
    Case 51
        cmdEscudo.Visible = True
    End Select
    
    PriGridGastos.CambioColumnaEuros 5, CodPaisTrabajo
    'lblTotalCampAux = etTotalCamp
    'etTotalCamp = dConvierteMoneda(CStr(etTotalCamp), 0, CByte(CodPaisTrabajo))
    etTotalCamp = dTotalCampanya

End Sub

Private Sub cmdLira_Click()
    cmdLira.Visible = False
    cmdEuro.Visible = True
    PriGridGastos.CambioColumnaMonedaLocal 5, CodPaisTrabajo, 2
    'etTotalCamp = dConvierteMoneda(CStr(etTotalCamp), 1, CByte(CodPaisTrabajo))
    'etTotalCamp = lblTotalCampAux
    etTotalCamp = dTotalCampanya
End Sub

Private Sub cmdPts_Click()
    cmdPts.Visible = False
    cmdEuro.Visible = True
    PriGridGastos.CambioColumnaMonedaLocal 5, CodPaisTrabajo, 2
    'etTotalCamp = dConvierteMoneda(CStr(etTotalCamp), 1, CByte(CodPaisTrabajo))
    'etTotalCamp = lblTotalCampAux
    etTotalCamp = dTotalCampanya
End Sub

Private Sub Form_Load()

PriGridGastos.Create gridGastos, 14, 2

PriGridGastos.Formatear 0, 0, ""
PriGridGastos.Formatear 1, 1000, "T.Gasto"
PriGridGastos.Formatear 2, 1000, "Fecha"
PriGridGastos.Formatear 3, 800, "T.Pago"
PriGridGastos.Formatear 4, 3400, "Donde"
PriGridGastos.Formatear 5, 1200, "Importe", 7, 7
PriGridGastos.Formatear 6, 700, "Numero", 1, 1
PriGridGastos.Formatear 7, 0, "Importe2"
PriGridGastos.Formatear 8, 600, "Hora", 1, 1
PriGridGastos.Formatear 9, 1000, "F.Gasto", 1, 1
PriGridGastos.Formatear 10, 2000, "Proveedor", 1, 1
PriGridGastos.Formatear 11, 2500, "Personas", 1, 1
PriGridGastos.Formatear 12, 0, "Ruta_Imagen"
PriGridGastos.Formatear 13, 1500, "Deducible", 1, 1

DesactivarCampos

Select Case CodPaisTrabajo
Case 34, 51, 39

    fraMoneda.Visible = True
    'optLocal.Value = True
    optEuro.Value = True
    Select Case CodPaisTrabajo
    Case 34
        cmdPts.Visible = True
    Case 39
        cmdLira.Visible = True
    Case 51
        cmdEscudo.Visible = True
    End Select
    
Case Else

    fraMoneda.Visible = False
    optEuro.Value = True
    cmdPts.Visible = False
    cmdLira.Visible = False
    cmdEscudo.Visible = False
    cmdEuro.Visible = False

End Select

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub gridGastos_Click()
    
    Dim Linea

    Dim sRuta As String

If gridGastos.Row = 0 Then Exit Sub
If PriGridGastos.bLineaActualVacia Then Exit Sub
If PriGridGastos.ColorDeFila(gridGastos.Row) = &HC0FFC0 Then
    If MsgBox("¿Desea ver el ticket asociado al gasto?", vbApplicationModal + vbQuestion + vbYesNo, "Confirmación") = vbYes Then
    
        sRuta = Trim(PriGridGastos.TextoLinCol(gridGastos.Row, 12))
    
        If sRuta = "" Then
            MsgBox "No hay asignado ningún ticket al gasto.", vbApplicationModal + vbInformation, "Información"
        Else
            sRuta = Mid(sRuta, 15)
            If LCase(Mid(sRuta, Len(sRuta) - 3)) = ".pdf" Then
                sRuta = "visortickets.cristianlay.com/" & sRuta
            Else
                sRuta = "visortickets.cristianlay.com/VisorTicket.asp?I=" & sRuta
            End If
            frmExplorer.Create sRuta
            frmExplorer.Show MODAL
        End If
    End If
    PriGridGastos.SituarEn 0, 1
    Exit Sub
End If

PriGridGastos.ResaltarLinea

If (CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51) Then
    If cmdEuro.Visible Then
        optLocal = True
    Else
        optEuro = True
    End If
End If

Linea = PriGridGastos.ObtenerTextoFila(PriGridGastos.LineaSeleccionada)

If Linea(0) = "" Then
    PriGridGastos.NoResaltarLineas
    Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
    boxTipoGasto = Trim(boxTipoGasto)
    boxTipoGasto.SetFocus
    Exit Sub
End If

Baja_Linea Linea, boxContador, boxTipoGasto, boxFecha, boxTipoPago, boxDonde, boxImporte, boxNumero, boxImporte2, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible

'boxFecha = Format(boxFecha, "ddmmyyyy")

boxTipoGasto = Trim(boxTipoGasto)

Select Case boxTipoGasto
    Case "Gasolina"
        boxNumero.Enabled = False
    Case "Salones"
        boxNumero.Enabled = False
    Case "Telefono"
        boxDonde.Enabled = False
        boxNumero.Enabled = False
    Case "Transporte"
        boxNumero.Enabled = False
End Select

If boxImagen.Text <> "" Then
    boxImagen.Text = Mid(boxImagen.Text, 15)
End If

boxTipoGasto.SetFocus
End Sub

Private Sub ActualizarGrid()

Dim rsGastos As ADODB.Recordset
Dim Jerarquia As String
Dim totalImporte, totalCamp As Double
Dim totalNumero As Double
Dim TipoGasto, strImporte, strNumero, TipoPago, sFamiliaGasto, sDeducible As String


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

sSql = "SELECT contador,fecha,nombregasto,g.tipogasto,tipopago,donde,importe,numero,chequeado,validado,hora,NVL(FamiliaGasto,'O') FamiliaGasto,NVL(Proveedor,'') Proveedor,NVL(Personas,'') Personas,NVL(Ruta_Ticket,'') Ruta_Imagen,NVL(Deducible,'S') Deducible " & _
       "FROM gastos g,tiposgasto t " & _
       "WHERE codigo=" & boxCodigo & " AND " & _
       "codpais=" & CodPaisTrabajo & " AND " & _
       "jerarquia='" & Jerarquia & "' AND " & _
       "anyo=" & boxAnyo & " AND " & _
       "campanya = " & boxCampanya & " AND " & _
       "g.tipogasto=t.tipogasto AND t.Tipo = 'C' " & _
       " ORDER BY nombregasto,fecha"

If gBdGALILEO.bRecordset(sSql, rsGastos) Then
    
    Screen.MousePointer = vbHourglass
    
    TipoGasto = rsGastos("nombregasto")
    totalNumero = 0
    totalImporte = 0
    totalCamp = 0
    strImporte = ""
    strNumero = ""
    While Not rsGastos.EOF
        Select Case rsGastos("tipopago")
            Case "M"
        TipoPago = "Metalico"
            Case "V"
        TipoPago = "Visa"
            Case "S"
        TipoPago = "Solred"
            Case "X"
        TipoPago = "Varios"
            Case Else
        TipoPago = ""
        End Select
        
        Select Case rsGastos("FamiliaGasto")
        Case "O"
            sFamiliaGasto = "Ordinario"
        Case "F"
            sFamiliaGasto = "Formación"
        Case "P"
            sFamiliaGasto = "Presentación"
        Case "I"
            sFamiliaGasto = "Incentivo"
        Case Else
            sFamiliaGasto = ""
        End Select
        sDeducible = "Sí"
        If Trim(rsGastos("Deducible")) = "N" Then sDeducible = "No"
        
        If TipoGasto <> rsGastos("nombregasto") Then
            If totalImporte <> 0 Then strImporte = FNumero(totalImporte, gDecimales)
            If totalNumero <> 0 Then strNumero = FNumero(totalNumero, gDecimales)
            
            PriGridGastos.TextoFila PriGridGastos.LineaLibre, _
            "", _
            "-------------------------------", _
            "--------------------------------------", _
            "--------------------------------------------", _
            "-------------  Subtotal  " & Trim(TipoGasto) & "  ------------------->", strImporte, strNumero, strImporte, ""
            PriGridGastos.ColorFila PriGridGastos.LineaLibre - 1, vbCyan
            
            TipoGasto = rsGastos("nombregasto")
                        
            totalCamp = totalCamp + totalImporte
            
            strImporte = ""
            strNumero = ""
            totalNumero = 0
            totalImporte = 0
        End If
        
        PriGridGastos.TextoFila PriGridGastos.LineaLibre, _
            rsGastos("contador"), _
            rsGastos("nombregasto"), _
            NoNull(rsGastos("fecha"), "F"), _
            TipoPago, rsGastos("donde"), _
            rsGastos("importe"), _
            rsGastos("numero"), _
            rsGastos("importe"), rsGastos("Hora"), sFamiliaGasto, Trim(rsGastos("Proveedor")), Trim(rsGastos("Personas")), Trim(rsGastos("Ruta_Imagen")), sDeducible
        
        If rsGastos("Chequeado") = "N" Then
            PriGridGastos.ColorFila PriGridGastos.LineaLibre - 1, &HC0C0FF
        End If
        If rsGastos("Validado") = "N" Then
            PriGridGastos.ColorFila PriGridGastos.LineaLibre - 1, &HC0FFC0
        End If
        
        totalNumero = totalNumero + rsGastos("numero")
        
        If rsGastos("TipoGasto") <> "K" Then
            totalImporte = totalImporte + rsGastos("importe")
        Else
            totalImporte = totalImporte + (rsGastos("importe") * rsGastos("numero"))
        End If
                        
        rsGastos.MoveNext
    Wend

totalCamp = totalCamp + totalImporte

If totalImporte <> 0 Then strImporte = FNumero(totalImporte, gDecimales)
If totalNumero <> 0 Then strNumero = FNumero(totalNumero, gDecimales)
  PriGridGastos.TextoFila PriGridGastos.LineaLibre, _
    "", _
    "-------------------------------", _
    "--------------------------------------", _
    "--------------------------------------------", _
    "-------------  Subtotal  " & Trim(TipoGasto) & "  ------------------->", strImporte, strNumero, strImporte, ""
    PriGridGastos.ColorFila PriGridGastos.LineaLibre - 1, vbCyan
    PriGridGastos.NoResaltarLineas

If (Not cmdEuro.Visible) And _
(CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51) Then
    PriGridGastos.CambioColumnaEuros 5, CodPaisTrabajo
    totalCamp = dConvierteMoneda(CStr(totalCamp), 0, CByte(CodPaisTrabajo))
End If

If totalCamp <> 0 Then _
    etTotalCamp = FNumero(totalCamp, gDecimales)
Else
    MsgBox "No existen datos de esa Campaña", vbApplicationModal + vbCritical
    etTotalCamp = ""
End If

boxTipoGasto.SetFocus
Set rsGastos = Nothing

Screen.MousePointer = DEFAULT

End Sub

Private Sub DesactivarCampos()
gridGastos.Enabled = False
btBorrarDatos.Enabled = False
btSubirDatos.Enabled = False
boxTipoGasto.Enabled = False
boxTipoPago.Enabled = False
boxFecha.Enabled = False
boxDonde.Enabled = False
boxImporte.Enabled = False
boxNumero.Enabled = False
boxHora.Enabled = False
boxFamiliaGasto.Enabled = False
boxProveedor.Enabled = False
boxPersonas.Enabled = False
boxDeducible.Enabled = False
cmdChequear.Enabled = False
cmdChequearLinea.Enabled = False
fraMoneda.Enabled = False
optLocal.Enabled = False
optEuro.Enabled = False
cmdEuro.Enabled = False
If CodPaisTrabajo = 34 Or CodPaisTrabajo = 39 Or CodPaisTrabajo = 51 Then
    cmdEuro.Visible = True
End If
cmdLira.Enabled = False
cmdLira.Visible = False
cmdEscudo.Enabled = False
cmdEscudo.Visible = False
cmdPts.Enabled = False
cmdPts.Visible = False
End Sub

Private Sub ActivarCampos()
gridGastos.Enabled = True
btBorrarDatos.Enabled = True
btSubirDatos.Enabled = True
boxTipoGasto.Enabled = True
boxTipoPago.Enabled = True
boxFecha.Enabled = True
boxDonde.Enabled = True
boxImporte.Enabled = True
boxNumero.Enabled = True
boxHora.Enabled = True
boxFamiliaGasto.Enabled = True
boxProveedor.Enabled = True
boxPersonas.Enabled = True
If CodPaisTrabajo = 52 Or CodPaisTrabajo = 56 Then boxDeducible.Enabled = True
cmdChequear.Enabled = True
cmdChequearLinea.Enabled = True
fraMoneda.Enabled = True
optLocal.Enabled = True
optEuro.Enabled = True
cmdEuro.Enabled = True
cmdLira.Enabled = True
cmdEscudo.Enabled = True
cmdPts.Enabled = True
End Sub

Private Function sObtenerFechaBaja() As String
    
    Dim sJerarquia As String
    Dim sTabla As String
    Dim afoRegistro As ADODB.Recordset
    
    On Error GoTo Error_Rutina
    
    sObtenerFechaBaja = "19000101"
    
    
    sSql = "SELECT FechaBaja FROM "
    
    Select Case boxJerarquia
    Case "1"
        sSql = sSql & "Directores " & _
               "WHERE Codigo = " & boxCodigo & _
               " AND Jerarquia = 'DN'"
    Case "2"
        sSql = sSql & "Directores " & _
               "WHERE Codigo = " & boxCodigo & _
               " AND Jerarquia = 'DR'"
    Case "3"
        sSql = sSql & "Directores " & _
               "WHERE Codigo = " & boxCodigo & _
               " AND Jerarquia IN ('DP','DS')"
    Case "4"
        sSql = sSql & "Directores " & _
               "WHERE Codigo = " & boxCodigo & _
               " AND Jerarquia = 'DA'"
    Case "5"
        sSql = sSql & "Supervisoras " & _
               "WHERE CodSuperv = " & boxCodigo
    Case Else
        Exit Function
    End Select

    sSql = sSql & " AND CodPais=" & CodPaisTrabajo

    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        MsgBox "Error al consultar la fecha de baja.", vbApplicationModal + vbExclamation
        Exit Function
    End If
    
    If IsNull(afoRegistro("FechaBaja")) Or _
    afoRegistro("FechaBaja") = "" Then
        
        sObtenerFechaBaja = "30001231"
        
    Else
    
        sObtenerFechaBaja = Format(afoRegistro("FechaBaja"), "yyyymmdd")
    
    End If
    
    Set afoRegistro = Nothing
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Function


Private Sub imgConsultar_Click()
    
'    Dim sRuta As String
'
'    Me.MousePointer = DEFAULT
'    If boxContador.Text = "" Then Exit Sub
'    If boxImagen.Text = "" Then
'        MsgBox "No hay asignado ningún ticket al gasto.", vbApplicationModal + vbInformation, "Información"
'        Exit Sub
'    End If
'    If LCase(Mid(boxImagen.Text, Len(boxImagen.Text) - 3)) = ".pdf" Then
'        sRuta = "visortickets.cristianlay.com/" & boxImagen.Text
'    Else
'        sRuta = "visortickets.cristianlay.com/VisorTicket.asp?I=" & boxImagen.Text
'    End If
'    frmExplorer.Create sRuta
'    frmExplorer.Show MODAL

    Dim sRuta As String
    
    Me.MousePointer = DEFAULT
    If boxContador.Text = "" Then Exit Sub
    
    If Trim(boxImagen.Text) = "" Then
        MsgBox "No hay asignado ningún ticket al gasto.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    
    sRuta = LCase(Trim(boxImagen.Text))
    sRuta = Replace(sRuta, "\", "/")
    sRuta = Replace(sRuta, "c:/ficherosgv/", "")
    sRuta = "../Marte/" & sRuta
    
    If LCase(Mid(Trim(boxImagen.Text), Len(Trim(boxImagen.Text)) - 3)) = ".pdf" Then
        sRuta = "personal.cristianlayusers.com\Gastos\VisorTickets.asp?P=" & sRuta
    Else
        sRuta = "personal.cristianlayusers.com\Gastos\VisorTickets.asp?I=" & sRuta
    End If
    frmExplorer.Create sRuta
    frmExplorer.Show MODAL
    
End Sub

Private Sub imgConsultar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgPortapapeles_Click()
    Me.MousePointer = DEFAULT
    PriGridGastos.aClipBoard 1
    MsgBox "Información copiada al Portapapeles.", vbApplicationModal + vbInformation, "Información"
End Sub

Private Sub imgPortapapeles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub optEuro_Click()

    If boxImporte <> "" And boxImporte.Enabled Then
    
        boxImporte = ""
        boxImporte.SetFocus
        
    Else
    
        If boxImporte.Enabled Then
            boxTipoGasto.SetFocus
        End If
        
    End If

End Sub

Private Sub optLocal_Click()

    If boxImporte <> "" And boxImporte.Enabled Then
    
        boxImporte = ""
        boxImporte.SetFocus
    
    Else
    
        If boxImporte.Enabled Then
            boxTipoGasto.SetFocus
        End If
        
    End If
    
End Sub

Private Function dTotalCampanya() As Double

    Dim i As Integer
    
    dTotalCampanya = 0

    For i = 1 To gridGastos.Rows - 1
    
        gridGastos.Row = i
        
        gridGastos.col = 4
        If InStr(1, LCase(gridGastos.Text), "subtotal") <> 0 Then
            gridGastos.col = 5
            dTotalCampanya = dTotalCampanya + CDbl(gridGastos.Text)
        End If
        
    Next
    
End Function

Private Sub boxFamiliaGasto_GotFocus()
etQueFamiliaGasto.Visible = True
If boxFamiliaGasto = "" Then boxFamiliaGasto = "Ordinario"
TextGotFocus boxFamiliaGasto
Select Case boxFamiliaGasto
    Case "Ordinario"
        boxFamiliaGasto = "O"
    Case "Formación"
        boxFamiliaGasto = "F"
    Case "Presentación"
        boxFamiliaGasto = "P"
    Case "Incentivo"
        boxFamiliaGasto = "I"
    Case Else
        boxFamiliaGasto = ""
End Select
boxFamiliaGasto.SelStart = 0
boxFamiliaGasto.SelLength = Len(boxFamiliaGasto.Text)
End Sub

Private Sub boxFamiliaGasto_KeyPress(KeyAscii As Integer)
EntraTexto KeyAscii, boxFamiliaGasto, 1, "oOfFpPiI", True
If KeyAscii = 27 Then
    Limpia_Boxes boxContador, boxTipoGasto, boxTipoPago, boxFecha, boxDonde, boxImporte, boxNumero, boxHora, boxFamiliaGasto, boxProveedor, boxPersonas, boxImagen, boxDeducible
    PriGridGastos.NoResaltarLineas
    boxTipoGasto.SetFocus
End If
End Sub

Private Sub boxFamiliaGasto_LostFocus()
etQueFamiliaGasto.Visible = False
TextLostFocus boxFamiliaGasto
Select Case UCase(boxFamiliaGasto)
    Case "O"
        boxFamiliaGasto = "Ordinario"
    Case "F"
        boxFamiliaGasto = "Formación"
    Case "P"
        boxFamiliaGasto = "Presentación"
    Case "I"
        boxFamiliaGasto = "Incentivo"
    Case Else
        boxFamiliaGasto = ""
End Select

End Sub
