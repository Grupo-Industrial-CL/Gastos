VERSION 5.00
Begin VB.Form frmModifDirec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Director"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4965
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMatricula 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1005
      MaxLength       =   15
      TabIndex        =   5
      Top             =   2760
      Width           =   1125
   End
   Begin VB.TextBox boxCanal 
      Height          =   285
      Left            =   1005
      TabIndex        =   4
      Top             =   2145
      Width           =   3165
   End
   Begin VB.TextBox boxCC 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1710
      Width           =   2835
   End
   Begin VB.TextBox boxEmpresa 
      Height          =   285
      Left            =   990
      TabIndex        =   2
      Top             =   1230
      Width           =   3165
   End
   Begin VB.CommandButton btCerrar 
      Caption         =   "C&errar"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   2760
      Picture         =   "frmModifDirec.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton btAceptar 
      Caption         =   "Aceptar"
      Height          =   735
      Left            =   1200
      Picture         =   "frmModifDirec.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox boxCodigo 
      Height          =   285
      Left            =   885
      TabIndex        =   0
      Top             =   150
      Width           =   735
   End
   Begin VB.TextBox boxJerarquia 
      Height          =   285
      Left            =   2895
      TabIndex        =   1
      Top             =   150
      Width           =   375
   End
   Begin VB.Label lblMatricula 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Matricula"
      Enabled         =   0   'False
      Height          =   255
      Left            =   150
      TabIndex        =   16
      Top             =   2760
      Width           =   765
   End
   Begin VB.Label etAyuda 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1:DN  2:DR  3:DP  4:DA  5:SU"
      Height          =   465
      Left            =   3450
      TabIndex        =   15
      Top             =   75
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label etCanal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Canal"
      Height          =   255
      Left            =   150
      TabIndex        =   14
      Top             =   2160
      Width           =   525
   End
   Begin VB.Label etCC 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cta. Corriente"
      Height          =   255
      Left            =   135
      TabIndex        =   13
      Top             =   1695
      Width           =   1050
   End
   Begin VB.Label etEmpresa 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empresa"
      Height          =   255
      Left            =   150
      TabIndex        =   12
      Top             =   1230
      Width           =   720
   End
   Begin VB.Label etCodigo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Código"
      Height          =   255
      Left            =   165
      TabIndex        =   11
      Top             =   150
      Width           =   615
   End
   Begin VB.Label etNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   990
      TabIndex        =   10
      Top             =   645
      Width           =   3135
   End
   Begin VB.Label etJerarquia 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jerarquía"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   255
      Left            =   165
      TabIndex        =   8
      Top             =   660
      Width           =   660
   End
End
Attribute VB_Name = "frmModifDirec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub boxCodigo_GotFocus()
    etNombre = ""
    boxEmpresa = ""
    boxCanal = ""
    boxCC = ""
    txtMatricula = ""
    lblMatricula.Enabled = False
    txtMatricula.Enabled = False
    TextGotFocus boxCodigo
End Sub

Private Sub boxCodigo_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, boxCodigo, 6
End Sub

Private Sub boxCodigo_LostFocus()
 Dim rsNombre As ADODB.Recordset
 Dim Jerarquia As String
 Dim rsDatos As ADODB.Recordset
 
 On Error GoTo ErrorSub
 
 TextLostFocus boxCodigo
 boxJerarquia.Tag = ""
 
 If boxCodigo = "" Or boxJerarquia = "" Then Exit Sub
 
Select Case boxJerarquia
    Case "1"
        Jerarquia = "'DN'"
    Case "2"
        Jerarquia = "'DR'"
    Case "3"
        Jerarquia = "'DP','DS'"
    Case "4"
        Jerarquia = "'DA'"
    Case "5"
        Jerarquia = "'SU'"
    Case Else
        Exit Sub
 End Select

 If boxJerarquia <> 5 Then
    sSql = "SELECT nombre,canal,empresa,cc,jerarquia,NVL(Matricula,'') Matricula FROM directores WHERE " & _
        "codpais=" & CodPaisTrabajo & " AND " & _
        "jerarquia IN (" & Jerarquia & ") AND " & _
        "codigo=" & boxCodigo & " AND " & _
        "fechabaja IS NULL"
 Else
    sSql = "SELECT nombre FROM supervisoras WHERE " & _
        "codpais=" & CodPaisTrabajo & " AND " & _
        "codsuperv=" & boxCodigo & " AND " & _
        "fechabaja IS NULL"
 End If
 
 If gBdGALILEO.bRecordset(sSql, rsNombre) Then
    etNombre = rsNombre("nombre")
    
    If boxJerarquia <> "5" And CodPaisTrabajo = giPAIS_PORTUGAL Then
        lblMatricula.Enabled = True
        txtMatricula.Enabled = True
    End If
    
    If etNombre = "" Then
        
        MsgBox "No existe el código introducido o está dado de baja.", vbApplicationModal + vbInformation, "Información"
        txtMatricula = ""
        lblMatricula.Enabled = False
        txtMatricula.Enabled = False
        
        Exit Sub
        
    End If
    
    If boxJerarquia <> 5 Then
        If Not IsNull(rsNombre("empresa")) Then _
            boxEmpresa = rsNombre("empresa")
        If Not IsNull(rsNombre("canal")) Then _
            boxCanal = rsNombre("canal")
        If Not IsNull(rsNombre("cc")) Then _
            boxCC = rsNombre("cc")
        boxJerarquia.Tag = Trim(rsNombre("Jerarquia"))
        If CodPaisTrabajo = giPAIS_PORTUGAL Then txtMatricula = rsNombre("Matricula")
    Else
        sSql = "SELECT empresa,canal,cc FROM contratossup WHERE " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "codsuperv=" & boxCodigo
            
        If gBdGALILEO.bRecordset(sSql, rsDatos) Then
            
            If Not IsNull(rsDatos("empresa")) Then _
                boxEmpresa = rsDatos("empresa")
            If Not IsNull(rsDatos("canal")) Then _
                boxCanal = rsDatos("canal")
            If Not IsNull(rsDatos("cc")) Then _
                boxCC = rsDatos("cc")
        End If
    End If
 End If
 
Set rsNombre = Nothing
Set rsDatos = Nothing
 
 
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


Private Sub boxEmpresa_GotFocus()
TextGotFocus boxEmpresa
End Sub

Private Sub boxEmpresa_KeyPress(KeyAscii As Integer)
EntraTexto KeyAscii, boxEmpresa, 30, LetrasTodo
End Sub

Private Sub boxEmpresa_LostFocus()
TextLostFocus boxEmpresa
End Sub

Private Sub boxCanal_GotFocus()
TextGotFocus boxCanal
End Sub

Private Sub boxCanal_KeyPress(KeyAscii As Integer)
EntraTexto KeyAscii, boxCanal, 30, LetrasTodo
End Sub

Private Sub boxCanal_LostFocus()
TextLostFocus boxCanal
End Sub
Private Sub boxCC_GotFocus()
TextGotFocus boxCC
End Sub

Private Sub boxCC_KeyPress(KeyAscii As Integer)
EntraTexto KeyAscii, boxCC, 30, LetrasTodo
End Sub

Private Sub boxCC_LostFocus()
TextLostFocus boxCC
End Sub
Private Sub boxJerarquia_GotFocus()
    etNombre = ""
    boxEmpresa = ""
    boxCanal = ""
    boxCC = ""
    txtMatricula = ""
    lblMatricula.Enabled = False
    txtMatricula.Enabled = False
    TextGotFocus boxJerarquia
    etAyuda.Visible = True
End Sub

Private Sub boxJerarquia_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, boxJerarquia, 1
End Sub

Private Sub boxJerarquia_LostFocus()
 Dim rsNombre As ADODB.Recordset
 Dim Jerarquia As String
 Dim rsDatos As ADODB.Recordset
 
 On Error GoTo ErrorSub
 
 etAyuda.Visible = False
 TextLostFocus boxJerarquia
 boxJerarquia.Tag = ""
 
 If boxCodigo = "" Or boxJerarquia = "" Then Exit Sub
 
Select Case boxJerarquia
    Case "1"
        Jerarquia = "'DN'"
    Case "2"
        Jerarquia = "'DR'"
    Case "3"
        Jerarquia = "'DP','DS'"
    Case "4"
        Jerarquia = "'DA'"
    Case "5"
        Jerarquia = "'SU'"
    Case Else
        Exit Sub
 End Select
 
 If boxJerarquia <> 5 Then
    sSql = "SELECT nombre,empresa,canal,cc,jerarquia,NVL(Matricula,'') Matricula FROM directores WHERE " & _
        "codpais=" & CodPaisTrabajo & " AND " & _
        "jerarquia IN (" & Jerarquia & ") AND " & _
        "codigo=" & boxCodigo & " AND " & _
        "fechabaja IS NULL"
 Else
    sSql = "SELECT nombre FROM supervisoras WHERE " & _
        "codpais=" & CodPaisTrabajo & " AND " & _
        "codsuperv=" & boxCodigo & " AND " & _
        "fechabaja IS NULL"
 End If
 
 If gBdGALILEO.bRecordset(sSql, rsNombre) Then
    
    etNombre = rsNombre("nombre")
    
    If boxJerarquia <> "5" And CodPaisTrabajo = giPAIS_PORTUGAL Then
        lblMatricula.Enabled = True
        txtMatricula.Enabled = True
    End If
    
    If etNombre = "" Then
        
        MsgBox "No existe el código introducido o está dado de baja.", vbApplicationModal + vbInformation, "Información"
        txtMatricula = ""
        lblMatricula.Enabled = False
        txtMatricula.Enabled = False
        
        Exit Sub
        
    End If

    
    If boxJerarquia <> 5 Then
        If Not IsNull(rsNombre("empresa")) Then _
            boxEmpresa = rsNombre("empresa")
        If Not IsNull(rsNombre("canal")) Then _
            boxCanal = rsNombre("canal")
        If Not IsNull(rsNombre("cc")) Then _
            boxCC = rsNombre("cc")
        boxJerarquia.Tag = Trim(rsNombre("Jerarquia"))
        If CodPaisTrabajo = giPAIS_PORTUGAL Then txtMatricula = rsNombre("Matricula")
    Else
    
        sSql = "SELECT empresa,canal,cc FROM contratossup WHERE " & _
            "codpais=" & CodPaisTrabajo & " AND " & _
            "codsuperv=" & boxCodigo
            
        If gBdGALILEO.bRecordset(sSql, rsDatos) Then
            
            If Not IsNull(rsDatos("empresa")) Then _
                boxEmpresa = rsDatos("empresa")
            If Not IsNull(rsDatos("canal")) Then _
                boxCanal = rsDatos("canal")
            If Not IsNull(rsDatos("cc")) Then _
                boxCC = rsDatos("cc")
        End If
    End If
 End If
 
Set rsNombre = Nothing
Set rsDatos = Nothing

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
Dim Box

On Error GoTo ErrorSub

Set Box = Es_Vacio(boxCodigo, etNombre, boxJerarquia)
If Not (Box Is Nothing) Then
    Box.SetFocus
    Exit Sub
End If

Set Box = Nothing

Select Case boxJerarquia
    Case "1"
        Jerarquia = "DN"
    Case "2"
        Jerarquia = "DR"
    Case "3"
        Jerarquia = "DP"
        Jerarquia = boxJerarquia.Tag
    Case "4"
        Jerarquia = "DA"
    Case "5"
        Jerarquia = "SU"
    Case Else
        Exit Sub
 End Select
 
If boxJerarquia <> 5 Then
    sSql = "UPDATE directores SET " & _
        "empresa = '" & boxEmpresa & "'," & _
        "canal = '" & boxCanal & "'," & _
        "cc = '" & boxCC & "'"
    If CodPaisTrabajo = giPAIS_PORTUGAL Then sSql = sSql & ",Matricula = '" & Trim(txtMatricula) & "'"
    sSql = sSql & _
        " WHERE codpais = " & CodPaisTrabajo & " AND " & _
        "codigo = " & boxCodigo & " AND " & _
        "jerarquia ='" & Jerarquia & "'"
Else
    sSql = "UPDATE contratossup SET " & _
       "empresa = '" & boxEmpresa & "'," & _
        "canal = '" & boxCanal & "'," & _
        "cc = '" & boxCC & "' " & _
        "WHERE codpais = " & CodPaisTrabajo & " AND " & _
        "codsuperv = " & boxCodigo
End If
   

If gBdGALILEO.bExecute(sSql) Then
    MsgBox "Datos modificados correctamente", vbApplicationModal
    boxCodigo.SetFocus
Else
    MsgBox "Error al modificar los datos", vbApplicationModal + vbCritical
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

Private Sub btCerrar_Click()
Unload Me
End Sub

Private Sub txtMatricula_GotFocus()
    TextGotFocus txtMatricula
End Sub

Private Sub txtMatricula_KeyPress(KeyAscii As Integer)
    EntraTexto KeyAscii, txtMatricula, 15, LetrasTodo
End Sub

Private Sub txtMatricula_LostFocus()
    TextLostFocus txtMatricula
End Sub
