VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCambioPais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Pais"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   2130
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "C&errar"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   600
      Picture         =   "frmCambioPais.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2520
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid GridPaises 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4048
      _Version        =   393216
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmCambioPais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private priGridPaises As New CFlexGrid

Private Sub btCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim ColeccionPaises
Dim NombrePais As String
Dim rsPaises As New ADODB.Recordset

On Error GoTo ErrorSub

priGridPaises.Create GridPaises, 2, 2
priGridPaises.Formatear 0, 800, "Pais"
priGridPaises.Formatear 1, 1000, "Nombre"
priGridPaises.ResaltarLinea

For Each ColeccionPaises In goUsuario.PaisesAmbito
    NombrePais = ""
    sSql = "SELECT Nombre FROM Paises WHERE " & _
        "CodPais = " & ColeccionPaises
    If gBdGALILEO.bRecordset(sSql, rsPaises) Then
        NombrePais = rsPaises("Nombre")
    End If
    Set rsPaises = Nothing
    priGridPaises.TextoFila priGridPaises.LineaLibre, ColeccionPaises, NombrePais
Next

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

Private Sub GridPaises_Click()

Dim LineaEntera

On Error GoTo ErrorSub

If GridPaises.Row = 0 Then Exit Sub

priGridPaises.ResaltarLinea

LineaEntera = priGridPaises.ObtenerTextoFila(priGridPaises.LineaSeleccionada)

If Val(LineaEntera(0)) <> CodPaisTrabajo Then
    If MsgBox("¿Desea cambiar su pais de trabajo actual(" & CodPaisTrabajo & ") por (" & Trim(LineaEntera(0)) & ").?", vbApplicationModal + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
    CodPaisTrabajo = LineaEntera(0)
End If
Unload frmCambioPais

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

Private Sub GridPaises_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub GridPaises_KeyPress(KeyAscii As Integer)
priGridPaises.ResaltarLinea
If KeyAscii = 13 Then GridPaises_Click
End Sub
