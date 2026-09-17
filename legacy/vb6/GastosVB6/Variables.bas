Attribute VB_Name = "Variables"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Global gBdGALILEO As New cBaseDatos
Global gBdPERSEIDA As cBaseDatos
Global gBDGASTOS As cBaseDatos
Global goUsuario As CUsuario
Global MDIPadre As MDIForm

Global SeparadorDecimal As String
Global SeparadorMiles As String

Global sSql As String
Global RutaCopia As String

Global CodPaisTrabajo As Integer
Global Const gDecimales As Byte = 2
Global Empresa_CodPaisTrabajo As Integer

' MousePointer
Global Const DEFAULT = 0        ' 0 - Default
Global Const HOURGLASS = 11     ' 11 - Hourglass

'Global Const vbExclamation = vbExclamation
Global Const MB_ICONSTOP = vbCritical
Global Const MODAL = vbModal
Global Const KEY_RIGHT = vbKeyRight
Global Const ENTER = vbKeyReturn

Public oBapiControl As Object
Public oConnection As Object

Global gHoy As Date
Global gAnchoPantallaPrincipal As Long
Global gAnchoPantallaVirtual As Long

Global gsRutaEquipo As String

Sub Inicializar_Variables()
RutaCopia = ":\GastosViaje"
End Sub

Function DarRutaCorrecta() As String

Dim Fso As New FileSystemObject
Dim Drv As Drive
Dim RutaCopiaTemp As String

On Error GoTo ErrorSub

Set Fso = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set Drv = Fso.GetDrive(Fso.GetDriveName("V:"))
If Err.Number = 68 Then
    On Error GoTo ErrorSub
    Err.Clear
    If Dir("C" & RutaCopia, vbDirectory) = "" Then MkDir "C" & RutaCopia
    RutaCopiaTemp = "C" & RutaCopia
Else
    On Error GoTo ErrorSub
    If Drv.DriveType = Remote And Drv.ShareName = "\\Client\C$" Then
        If Dir("V" & RutaCopia, vbDirectory) = "" Then MkDir "V" & RutaCopia
        RutaCopiaTemp = "V" & RutaCopia
    Else
        If Dir("C" & RutaCopia, vbDirectory) = "" Then MkDir "C" & RutaCopia
        RutaCopiaTemp = "C" & RutaCopia
    End If
End If
Set Drv = Nothing
Set Fso = Nothing
DarRutaCorrecta = RutaCopiaTemp & "\"

ExitSub:
    Exit Function
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
End Function
