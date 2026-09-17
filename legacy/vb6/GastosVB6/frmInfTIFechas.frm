VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfTIFechas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencias Internas - Informe por fechas"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfTIFechas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRechazado 
      Caption         =   "Rechazadas"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Tag             =   "D"
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkParcialmenteRechazado 
      Caption         =   "Parcialmente rechazadas"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Tag             =   "E"
      Top             =   2400
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkParcialmenteAprobado 
      Caption         =   "Parcialmente aprobadas"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Tag             =   "B"
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkAprobado 
      Caption         =   "Totalmente aprobadas"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Tag             =   "C"
      Top             =   2040
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkPendientes 
      Caption         =   "Pendientes"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Tag             =   "A"
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   615
      Left            =   6720
      Picture         =   "frmInfTIFechas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFechaDesde 
      Height          =   285
      Left            =   4560
      TabIndex        =   0
      Top             =   435
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   19726339
      CurrentDate     =   38755
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   615
      Left            =   6720
      Picture         =   "frmInfTIFechas.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFechaHasta 
      Height          =   285
      Left            =   6720
      TabIndex        =   1
      Top             =   435
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   19726339
      CurrentDate     =   38755
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Estado de las transferencias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   2790
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "hasta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   10
      Top             =   480
      Width           =   540
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Informe de transferencias internas desde"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "frmInfTIFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAprobado_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = ENTER Then
        Select Case chkAprobado.Value
        Case 0
            chkAprobado.Value = 1
        Case Else
            chkAprobado.Value = 0
        End Select
        Tabular
    End If
End Sub

Private Sub chkParcialmenteAprobado_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = ENTER Then
        Select Case chkParcialmenteAprobado.Value
        Case 0
            chkParcialmenteAprobado.Value = 1
        Case Else
            chkParcialmenteAprobado.Value = 0
        End Select
        Tabular
    End If
End Sub

Private Sub chkParcialmenteRechazado_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = ENTER Then
        Select Case chkParcialmenteRechazado.Value
        Case 0
            chkParcialmenteRechazado.Value = 1
        Case Else
            chkParcialmenteRechazado.Value = 0
        End Select
        Tabular
    End If
End Sub

Private Sub chkPendientes_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = ENTER Then
        Select Case chkPendientes.Value
        Case 0
            chkPendientes.Value = 1
        Case Else
            chkPendientes.Value = 0
        End Select
        Tabular
    End If
End Sub

Private Sub chkRechazado_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = ENTER Then
        Select Case chkRechazado.Value
        Case 0
            chkRechazado.Value = 1
        Case Else
            chkRechazado.Value = 0
        End Select
        Tabular
    End If
End Sub

Private Sub cmdConsultar_Click()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    Dim afoRegAux1 As ADODB.Recordset
    
    Dim sFichero As String
    Dim iFichero As Integer
    Dim sLinea As String
    
    Dim dImporte As Double
    Dim dTotal As Double
    Dim sEstados As String
    
    On Error GoTo Error_Rutina
    
    sEstados = ""
    If chkAprobado.Value = 1 Then sEstados = "'" & chkAprobado.Tag & "',"
    If chkPendientes.Value = 1 Then sEstados = sEstados & "'" & chkPendientes.Tag & "',"
    If chkParcialmenteAprobado.Value = 1 Then sEstados = sEstados & "'" & chkParcialmenteAprobado.Tag & "',"
    If chkParcialmenteRechazado.Value = 1 Then sEstados = sEstados & "'" & chkParcialmenteRechazado.Tag & "',"
    If chkRechazado.Value = 1 Then sEstados = sEstados & "'" & chkRechazado.Tag & "',"
    If sEstados = "" Then
        MsgBox "Seleccione algún estado de la transferencia para hacer la consulta.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    sEstados = "(" & Mid(sEstados, 1, Len(sEstados) - 1) & ")"
    
    Screen.MousePointer = HOURGLASS

    sSql = "SELECT * FROM c_Transf_Int " & _
           "WHERE Fecha BETWEEN '" & Format(dtpFechaDesde, "yyyy-MM-dd") & _
           "' AND '" & Format(dtpFechaHasta, "yyyy-MM-dd") & "' AND Status IN " & sEstados
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Screen.MousePointer = DEFAULT
        MsgBox "No existen transferencias entre las fechas indicadas.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    
    sFichero = Dir("C:\Listados_TI", vbDirectory)
    If sFichero = "" Then MkDir "C:\Listados_TI"
    iFichero = FreeFile
    sFichero = "C:\Listados_TI\Listado_TI_" & Format(dtpFechaDesde, "ddMMyy") & "_" & Format(dtpFechaHasta, "ddMMyy") & ".csv"
    Open sFichero For Output As #iFichero
    
    While Not afoRegistro.EOF
        
        sLinea = "Dep. Proveedor: " & sNombreDepartamento(afoRegistro("Proveedor"), afoRegistro("CodPais_P")) & "(" & afoRegistro("Proveedor") & ")    Dep. Cliente: " & sNombreDepartamento(afoRegistro("Cliente"), afoRegistro("CodPais_C")) & "(" & afoRegistro("Cliente") & ")     Estado: " & sEstado(afoRegistro("Status"), "C") & "     Usuario entrada: " & sNombreUsuario(afoRegistro("Usu_Entrada"), afoRegistro("Proveedor"), afoRegistro("CodPais_P")) & "     Notificado el: " & NoNull(afoRegistro("Fec_Env_Not"), "T")
        Print #iFichero, sLinea
        sLinea = ";"
        Print #iFichero, sLinea
        sLinea = ";Concepto;Cantidad;Unidad;Precio;Importe;Estado"
        Print #iFichero, sLinea
        
        dTotal = 0
        Set afoRegAux1 = Nothing
        sSql = "SELECT * FROM l_Transf_Int " & _
               "WHERE Id_Transferencia = " & afoRegistro("Id_Transferencia") & _
               " ORDER BY Linea"
        If gBdGALILEO.bRecordset(sSql, afoRegAux1) Then
            While Not afoRegAux1.EOF
                dImporte = afoRegAux1("Precio") * afoRegAux1("Cantidad")
                dTotal = dTotal + dImporte
                sLinea = ";" & Trim(afoRegAux1("Concepto")) & ";" & _
                    afoRegAux1("Cantidad") & ";" & Trim(afoRegAux1("Unidad")) & ";" & _
                    afoRegAux1("Precio") & ";" & dImporte & ";" & _
                    sEstado(afoRegAux1("Status"), "L")
                Print #iFichero, sLinea
                afoRegAux1.MoveNext
            Wend
            sLinea = ";;;;Total;" & dTotal
            Print #iFichero, sLinea
        End If
        Set afoRegAux1 = Nothing
        
        sLinea = ";"
        Print #iFichero, sLinea
        Print #iFichero, sLinea
        
        afoRegistro.MoveNext
        
    Wend
    Set afoRegistro = Nothing
    Close #iFichero
    
    Screen.MousePointer = DEFAULT
    MsgBox "Listado generado en '" & sFichero & "'.", vbApplicationModal + vbInformation, "Información"

Exit_Rutina:
    Exit Sub
Error_Rutina:
    Screen.MousePointer = DEFAULT
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Function sNombreUsuario(lCodigo As Long, iDepartamento As Integer, _
    iCodPais As Integer) As String

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset

    On Error GoTo Error_Rutina

    sNombreUsuario = "Desconocido"
    
    sSql = "SELECT Nombre FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & _
           " AND Codigo = " & lCodigo & _
           " AND Departamento = " & iDepartamento
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then sNombreUsuario = Trim(afoRegistro("Nombre"))
    Set afoRegistro = Nothing
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Function
Private Function sEstado(sStatus As String, sTipoEstado As String) As String

    On Error GoTo Error_Rutina

    sEstado = "-"
    Select Case sTipoEstado
    Case "C"
    
        Select Case sStatus
        Case "R"
            sEstado = "Rellenándose"
        Case "A"
            sEstado = "Pendiente"
        Case "B"
            sEstado = "Parcialmente aprobado"
        Case "C"
            sEstado = "Totalmente aprobado"
        Case "D"
            sEstado = "Rechazado"
        Case "E"
            sEstado = "Parcialmente rechazado"
        End Select
    
    Case "L"
    
        Select Case sStatus
        Case "A"
            sEstado = "Pendiente"
        Case "C"
            sEstado = "Aprobado"
        Case "D"
            sEstado = "Rechazado"
        End Select
    
    End Select

Exit_Rutina:
    Exit Function
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Function

Private Function sNombreDepartamento(lDepartamento As Long, iCodPais As Integer) As String
    
    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    On Error GoTo Error_Rutina

    sNombreDepartamento = "Desconocido"
    sSql = "SELECT Descripcion FROM DepartamentosCL " & _
           "WHERE CodPais = " & iCodPais & _
           " AND Codigo = " & lDepartamento
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then sNombreDepartamento = Trim(afoRegistro("Descripcion"))
    Set afoRegistro = Nothing
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Function

Private Sub cmdVolver_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    dtpFechaDesde = Now
    dtpFechaHasta = Now
    dtpFechaDesde.MaxDate = Now
    dtpFechaHasta.MaxDate = Now
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

