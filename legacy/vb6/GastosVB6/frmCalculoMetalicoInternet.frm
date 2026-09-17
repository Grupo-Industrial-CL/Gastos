VERSION 5.00
Begin VB.Form frmCalculoMetalicoInternet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCalculoMetalicoInternet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIdentificador 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   320
      Width           =   2175
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   885
      Width           =   2415
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Total metálico"
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   990
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Identificador"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   915
   End
End
Attribute VB_Name = "frmCalculoMetalicoInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVolver_Click()
    Unload Me
End Sub

Private Sub txtIdentificador_GotFocus()
    TextGotFocus txtIdentificador
End Sub

Private Sub txtIdentificador_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        CalcularTotal
    End If
    
    EntraNumero KeyAscii, txtIdentificador, 15
    
End Sub

Private Sub txtIdentificador_LostFocus()
    TextLostFocus txtIdentificador
End Sub

Private Sub CalcularTotal()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    Dim Total As Double
    
    On Error GoTo Error_Rutina
    
    lblTotal = 0
    Me.MousePointer = 11
    
    sSql = "SELECT * FROM Gastos " & _
           "WHERE Identificador = '" & Trim(txtIdentificador) & "'"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        MsgBox "No existen gastos con ese identificador.", vbApplicationModal
        Set afoRegistro = Nothing
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Total = 0
    Set afoRegistro = Nothing
    
    sSql = "SELECT SUM(Importe) AS Total FROM Gastos " & _
           "WHERE Identificador = '" & Trim(txtIdentificador) & "' AND " & _
           "TipoPago = 'M' AND TipoGasto <> 'K'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Total = afoRegistro("Total")
    End If
    
    Set afoRegistro = Nothing
    sSql = "SELECT SUM(Importe * Numero) AS Total FROM Gastos " & _
           "WHERE Identificador = '" & Trim(txtIdentificador) & "' AND " & _
           "TipoPago = 'M' AND TipoGasto = 'K'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Total = Total + afoRegistro("Total")
    End If
    
    lblTotal = dConvierteMoneda(CStr(Total), 0, CByte(CodPaisTrabajo))
    Me.MousePointer = vbDefault
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub
