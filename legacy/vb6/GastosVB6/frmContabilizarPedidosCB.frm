VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmContabilizarPedidosCB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilizar Pedidos Caja Badajoz"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContabilizarPedidosCB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContabilizar 
      Caption         =   "Contabilizar"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   300
      Left            =   360
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   59899907
      CurrentDate     =   38230
   End
   Begin MSFlexGridLib.MSFlexGrid mfgParrilla 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3836
      _Version        =   393216
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   6360
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de contabilización"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   2070
   End
   Begin VB.Label lblPedido 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Pedido"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   570
   End
End
Attribute VB_Name = "frmContabilizarPedidosCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moParrilla As New CFlexGrid

Private Sub cmdContabilizar_Click()
    
    Dim sSql As String
    
    On Error GoTo Error_Rutina

    If lblPedido = "" Then Exit Sub
    If lblPedido.Tag = "" Then Exit Sub
    
    sSql = "UPDATE PedCB SET Status = 'F'," & _
           "FecContabilizado = '" & Format(dtpFecha.Value, "ddMMyyyy") & _
           "' WHERE Id_Pedido = " & lblPedido & _
           " AND Codigo = " & lblPedido.Tag
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "No se ha podido guardar la información.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    lblPedido = ""
    lblPedido.Tag = ""
    moParrilla.ElimLinea
    moParrilla.SituarEn 0, 1

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

    On Error GoTo Error_Rutina
    
    dtpFecha.Value = Now
    dtpFecha.MaxDate = Now

    moParrilla.Create mfgParrilla, 6, 2
    moParrilla.Formatear 0, 0, ""
    moParrilla.Formatear 1, 1000, "Pedido", 4, 7, True
    moParrilla.Formatear 2, 1000, "Oficina", 4, 7, True
    moParrilla.Formatear 3, 2750, "Nombre Oficina", 4
    moParrilla.Formatear 4, 1250, "Importe", 4, 7
    moParrilla.Formatear 5, 1500, "Fecha Envio", 4, 4
    
    RellenarParrilla

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub RellenarParrilla()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    On Error GoTo Error_Rutina

    moParrilla.Vaciar
    
    sSql = "SELECT p.Id_Pedido,p.Codigo,d.Nombre,p.Importe,p.FecEnviado " & _
           "FROM PedCB p,Distribuidoras d " & _
           "WHERE p.Codigo = d.CodDistrib AND d.CodPais = 34 " & _
           "AND p.Codigo BETWEEN 300000 AND 400000 " & _
           "AND Status IN ('D','E') " & _
           "ORDER BY p.Codigo"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Sub
    End If
    
    While Not afoRegistro.EOF
    
        moParrilla.TextoFila moParrilla.LineaLibre, afoRegistro("Id_Pedido"), _
            afoRegistro("Id_Pedido"), afoRegistro("Codigo"), _
            Trim(afoRegistro("Nombre")), afoRegistro("Importe"), _
            Format(afoRegistro("FecEnviado"), "dd/MM/yyyy")
        
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

Private Sub mfgParrilla_Click()
    
    If moParrilla.Grid.MouseRow = 0 Then moParrilla.OrdenarSiCorresponde
    If moParrilla.Grid.Row = 0 Then
        lblPedido = ""
        lblPedido.Tag = ""
        moParrilla.SituarEn 0, 1
        Exit Sub
    End If
    
    lblPedido = moParrilla.TextoLinCol(moParrilla.Grid.Row, 0)
    lblPedido.Tag = moParrilla.TextoLinCol(moParrilla.Grid.Row, 2)
    moParrilla.ResaltarLinea
    
End Sub
