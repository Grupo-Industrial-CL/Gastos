VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmValidarGastos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validar gastos"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmValidarGastos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid mfgParrilla 
      Height          =   2655
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4683
      _Version        =   393216
   End
   Begin VB.Frame fraDatos 
      Height          =   2175
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   9015
      Begin VB.CommandButton cmdValidar 
         Caption         =   "Validar"
         Height          =   615
         Left            =   6120
         Picture         =   "frmValidarGastos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   615
         Left            =   3240
         Picture         =   "frmValidarGastos.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdIniciar 
         Caption         =   "Iniciar"
         Height          =   615
         Left            =   4680
         Picture         =   "frmValidarGastos.frx":171E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdVolver 
         Caption         =   "Volver"
         Height          =   615
         Left            =   7560
         Picture         =   "frmValidarGastos.frx":1FE8
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox cmbJerarquia 
         Height          =   315
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   320
         Width           =   1815
      End
      Begin VB.TextBox txtCampanya 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1515
         Width           =   375
      End
      Begin VB.TextBox txtAnyo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   3
         Top             =   920
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   1
         Top             =   320
         Width           =   975
      End
      Begin VB.ComboBox cmbPaises 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   320
         Width           =   2175
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   16
         Top             =   915
         Width           =   4935
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Campaña"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   330
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Jerarquía"
         Height          =   195
         Index           =   2
         Left            =   6000
         TabIndex        =   13
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   12
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "País"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmValidarGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moParrilla As New CFlexGrid

Dim mlCodigo As Long
Dim msJerarquia As String
Dim miAnyo As Integer
Dim myCampanya As Byte
Dim miCodPais As Integer

Private Sub cmbJerarquia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = ENTER Then Tabular
End Sub

Private Sub cmbPaises_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = ENTER Then Tabular
End Sub

Private Sub cmdConsultar_Click()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset

    On Error GoTo Error_Rutina

    If cmbPaises = "" Then
        cmbPaises.SetFocus
        Exit Sub
    End If
    txtCodigo = Trim(txtCodigo)
    If txtCodigo = "" Then
        txtCodigo.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtCodigo) Then
        txtCodigo = ""
        txtCodigo.SetFocus
        Exit Sub
    End If
    If cmbJerarquia = "" Then
        cmbJerarquia.SetFocus
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
    If txtAnyo > Year(Now) Then
        txtAnyo = ""
        txtAnyo.SetFocus
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
    
    mlCodigo = txtCodigo
    miCodPais = CInt(Trim(Mid(cmbPaises.Text, 1, InStr(1, cmbPaises.Text, "-") - 1)))
    sSql = "SELECT * FROM "
    Select Case cmbJerarquia.ListIndex
    Case 0, 1, 2, 3
        sSql = sSql & "Directores WHERE CodPais = " & miCodPais & _
               " AND Codigo = " & mlCodigo & " AND Jerarquia IN ('"
        Select Case cmbJerarquia.ListIndex
        Case 0
            msJerarquia = "DN"
        Case 1
            msJerarquia = "DR"
        Case 2
            msJerarquia = "DA"
        Case 3
            msJerarquia = "DP"
            If bEsCodigoDS(miCodPais, mlCodigo) = True Then msJerarquia = msJerarquia & "','DS"
            'If (miCodPais = 34 And mlCodigo = 979) Or (miCodPais = 51 And mlCodigo = 459) Or (miCodPais = 7 And mlCodigo = 19) Then msJerarquia = "DS"
        Case Else
            msJerarquia = "EE"
        End Select
        sSql = sSql & msJerarquia & "') "
    Case 4
        msJerarquia = "SU"
        sSql = sSql & "Supervisoras WHERE CodPais = " & miCodPais & _
               " AND CodSuperv = " & mlCodigo
    End Select
    
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        MsgBox "Error al consultar el código.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    lblNombre = Trim(afoRegistro("Nombre"))
    Set afoRegistro = Nothing

    miAnyo = txtAnyo
    myCampanya = txtCampanya
    sSql = "SELECT g.*,t.NombreGasto FROM Gastos g,TiposGasto t " & _
           "WHERE CodPais = " & miCodPais & _
           " AND Codigo = " & mlCodigo & _
           " AND Jerarquia IN ('" & msJerarquia & _
           "') AND Anyo = " & miAnyo & _
           " AND Campanya = " & myCampanya & _
           " AND Validado = 'N' AND g.TipoGasto = t.TipoGasto AND t.Tipo = 'C' AND Identificador <> 0"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        MsgBox "No existen gastos que validar.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    While Not afoRegistro.EOF
        moParrilla.TextoFila moParrilla.LineaLibre, "", Format(afoRegistro("Fecha"), "dd/MM/yy"), _
            Trim(afoRegistro("NombreGasto")), afoRegistro("Importe"), afoRegistro("Numero"), _
            Trim(afoRegistro("Donde")), Trim(afoRegistro("TipoPago"))
        afoRegistro.MoveNext
    Wend
    Set afoRegistro = Nothing
    moParrilla.SituarEn 0, 1
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdIniciar_Click()
    
    moParrilla.Vaciar
    moParrilla.SituarEn 0, 1
    cmbPaises.ListIndex = -1
    txtCodigo = ""
    cmbJerarquia.ListIndex = -1
    txtAnyo = ""
    txtCampanya = ""
    lblNombre = ""
    mlCodigo = -1
    msJerarquia = ""
    miAnyo = -1
    myCampanya = 0
    miCodPais = -1
    cmbPaises.SetFocus
    
End Sub

Private Sub cmdValidar_Click()

    Dim sSql As String

    On Error GoTo Error_Rutina

    If lblNombre = "" Then
        MsgBox "Rellene todos los datos y pulse el botón 'Consultar' para ver los gastos pendientes de validar.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    moParrilla.SituarEn 0, 1
    If moParrilla.bLineaActualVacia Then
        MsgBox "No existen datos que validar con los datos introducidos.", vbApplicationModal + vbInformation, "Información"
        cmdIniciar_Click
        Exit Sub
    End If
    
    sSql = "UPDATE Gastos SET Validado = 'S',Validacion = 'Desde Aplic. Gastos' " & _
           "WHERE CodPais = " & miCodPais & _
           " AND Codigo = " & mlCodigo & _
           " AND Jerarquia IN ('" & msJerarquia & _
           "') AND Anyo = " & miAnyo & _
           " AND Campanya = " & myCampanya & _
           " AND Validado = 'N' AND Identificador <> 0"
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "No se han podido validar los gastos.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    
    MsgBox "Datos validados correctamente.", vbApplicationModal + vbInformation, "Información"
    cmdIniciar_Click

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub Form_Load()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset

    On Error GoTo Error_Rutina
    
    moParrilla.Create mfgParrilla, 7, 2
    moParrilla.Formatear 0, 0, ""
    moParrilla.Formatear 1, 1250, "Fecha", 4, 4
    moParrilla.Formatear 2, 1500, "Gasto", 4
    moParrilla.Formatear 3, 1250, "Importe", 4, 7
    moParrilla.Formatear 4, 1250, "Número", 4, 7
    moParrilla.Formatear 5, 3000, "Concepto", 4
    moParrilla.Formatear 6, 1000, "Pago", 4, 4

    sSql = "SELECT CodPais,Nombre FROM Paises " & _
           "WHERE NOT CodPais IN (55,94,95,96,97,98,99,34)"
    cmbPaises.AddItem "34 - España"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        While Not afoRegistro.EOF
            cmbPaises.AddItem afoRegistro("CodPais") & " - " & Trim(afoRegistro("Nombre"))
            afoRegistro.MoveNext
        Wend
    End If
    Set afoRegistro = Nothing
    
    cmbJerarquia.AddItem "Director Nacional"
    cmbJerarquia.AddItem "Director Regional"
    cmbJerarquia.AddItem "Director Área"
    cmbJerarquia.AddItem "Director Apoyo"
    cmbJerarquia.AddItem "Supervisora"

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub mfgParrilla_Click()
    moParrilla.SituarEn 0, 1
End Sub

Private Sub mfgParrilla_SelChange()
    moParrilla.SituarEn 0, 1
End Sub

Private Sub txtAnyo_GotFocus()
    TextGotFocus txtAnyo
End Sub

Private Sub txtAnyo_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    EntraNumero KeyAscii, txtAnyo, 4
End Sub

Private Sub txtAnyo_LostFocus()
    TextLostFocus txtAnyo
End Sub

Private Sub txtCampanya_GotFocus()
    TextGotFocus txtCampanya
End Sub

Private Sub txtCampanya_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    EntraNumero KeyAscii, txtCampanya, 2
End Sub

Private Sub txtCampanya_LostFocus()
    TextLostFocus txtCampanya
End Sub

Private Sub txtCodigo_GotFocus()
    TextGotFocus txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    EntraNumero KeyAscii, txtCodigo, 6
End Sub

Private Sub txtCodigo_LostFocus()
    TextLostFocus txtCodigo
End Sub
