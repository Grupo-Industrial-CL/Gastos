Attribute VB_Name = "Funciones"
Option Explicit
Private Const sModulo_ As String = "Funciones"

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'AVERIGUAR EL NOMBRE DEL PC
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Global Const LetrasAZ As String = "abcdefghijklmnÒopqrstuvwxyzABCDEFGHIJKLMN—OPQRSTUVWXYZ"
Global Const LetrasAZSin— As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
Global Const LetrasEspeciales = "·ÈÌÛ˙¡…Õ”⁄‡ËÏÚ˘¿»Ã“Ÿ‰ÎÔˆ¸ƒÀœ÷‹‚ÍÓÙ˚¬ Œ‘€'Á«"
Global Const SignosEspeciales = "ø?!°()[]{},:;/"
Global Const Letras09 As String = "0123456789"
Global Const LetrasSigno As String = "-_@. "
Global Const LetrasTodo As String = LetrasAZ & Letras09 & LetrasSigno

Global Const dCambioMoneda_Pesetas As Double = 166.386
Global Const dCambioMoneda_Liras As Double = 1936.27
Global Const dCambioMoneda_Escudos As Double = 200.482

Public Enum Operaciones
    Multiplicar
    Sumar
    Restar
    Dividir
End Enum
Public Enum DireccionCambio
    Local_A_Euros
    Euros_A_Local
End Enum

Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" ( _
    ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Global Const SYNCHRONIZE = &H100000
Global Const INFINITE = &HFFFFFFFF

Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Const SM_CXVIRTUALSCREEN = 78
Private Const SM_CYVIRTUALSCREEN = 79
Private Const SM_CMONITORS = 80
Private Const SM_SAMEDISPLAYFORMAT = 81

Private Declare Function GetSystemMetrics Lib "user32" ( _
   ByVal nIndex As Long) As Long

Public Property Get VirtualScreenWidth() As Long
   VirtualScreenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
End Property
Public Property Get VirtualScreenHeight() As Long
   VirtualScreenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
End Property
Public Property Get DisplayMonitorCount() As Long
   DisplayMonitorCount = GetSystemMetrics(SM_CMONITORS)
End Property
Public Property Get AllMonitorsSame() As Long
   AllMonitorsSame = GetSystemMetrics(SM_SAMEDISPLAYFORMAT)
End Property

Sub EsperarProceso(ByVal PID As Long)
If PID <> 0 Then
    WaitForTerm PID
End If
End Sub

Private Sub WaitForTerm(ByVal PID As Long)
On Error GoTo Gestion_Error

'Variables locales
Dim phnd As Long

phnd = OpenProcess(SYNCHRONIZE, 0, PID)
If phnd <> 0 Then
Call WaitForSingleObject(phnd, INFINITE)
Call CloseHandle(phnd)
End If
Exit Sub
Gestion_Error:
Call MsgBox(Err.Number & ": " & Err.Description)
End Sub

Function sRecordsetToString(ByVal rsRecordSet As ADODB.Recordset, ByVal sCampo As String, Optional ByVal sSeparador As String = ",", Optional ByVal sEntreCorchetes As String = vbNullString, Optional ByVal sValorSiVacio As String = vbNullString, Optional ByVal bEvitarDuplicados As Boolean = False) As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "sRecordsetToString"

Dim CadenaControlCodigos As String

sRecordsetToString = vbNullString
CadenaControlCodigos = ""
If rsRecordSet.State = adStateOpen Then
    If Not rsRecordSet.EOF Then rsRecordSet.MoveFirst
    While Not rsRecordSet.EOF
        If bEvitarDuplicados Then
            If InStr(1, CadenaControlCodigos, "[" & rsRecordSet(sCampo) & "]") > 0 Then
            Else
                sRecordsetToString = sRecordsetToString & sSeparador & sEntreCorchetes & rsRecordSet(sCampo) & sEntreCorchetes
                CadenaControlCodigos = CadenaControlCodigos & "[" & rsRecordSet(sCampo) & "]"
            End If
        Else
            sRecordsetToString = sRecordsetToString & sSeparador & sEntreCorchetes & rsRecordSet(sCampo) & sEntreCorchetes
        End If
        
        rsRecordSet.MoveNext
    Wend
End If
If Len(sRecordsetToString) <> 0 Then
    sRecordsetToString = Mid(sRecordsetToString, 2)
Else
    If Len(sValorSiVacio) <> 0 Then sRecordsetToString = sValorSiVacio
End If

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Public Sub TextGotFocus(Control As Control)

On Error GoTo ControlErrores: Const sNombreSub_ As String = "TextGotFocus"

With Control
    .BackColor = &H80FFFF
    If TypeOf Control Is TextBox Then
        .SelStart = 0
        .SelLength = Len(.Text)

        If Len(.Text) <> 0 And Not IsNull(.Text) Then
            If IsNumeric(.Text) Then
                If Val(.Text) < 0 Then
                    .ForeColor = &HFF
                Else
                    .ForeColor = &H80000008
                End If
            End If
        End If
    End If
End With

ExitSub:
    Exit Sub
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Sub
    End Select
End Sub

Public Sub TextLostFocus(Control As Control)

On Error GoTo ControlErrores: Const sNombreSub_ As String = "TextLostFocus"

With Control
    .BackColor = vbWhite
    If TypeOf Control Is TextBox Then
        If Len(.Text) <> "" And Not IsNull(.Text) Then
            If IsNumeric(.Text) Then
                If Val(.Text) < 0 Then
                    .ForeColor = &HFF
                Else
                    .ForeColor = &H80000008
                End If
            End If
        End If
    End If
End With

ExitSub:
    Exit Sub
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Sub
    End Select
End Sub

Public Sub FechaGotFocus(Control As Control)

Dim FechaTemp As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "FechaGotFocus"

With Control
    If Len(.Text) = 10 Then
        FechaTemp = Mid$(.Text, 1, 2) & Mid$(.Text, 4, 2) & Mid$(.Text, 7, 4)
        .Text = FechaTemp
    End If
    .BackColor = &H80FFFF
    .SelStart = 0
    .SelLength = Len(.Text)
End With

ExitSub:
    Exit Sub
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Sub
    End Select
End Sub

Public Sub FechaLostFocus(Control As Control)

Dim DiaTemp, MesTemp, AnyoTemp, AnyoActual As String
Dim FechaTempValida As Boolean

On Error GoTo ControlErrores: Const sNombreSub_ As String = "FechaLostFocus"

With Control
    .BackColor = vbWhite
    FechaTempValida = False
    If Len(.Text) = 8 Then
        DiaTemp = Mid$(.Text, 1, 2)
        MesTemp = Mid$(.Text, 3, 2)
        AnyoTemp = Mid$(.Text, 5, 4)
        AnyoActual = Val(Format$(Now, "yyyy"))
      If Val(AnyoTemp) > 1900 And Val(AnyoTemp) < AnyoActual + 10 Then
        If IsNumeric(MesTemp) Then
            Select Case Val(MesTemp)
                Case 1, 3, 5, 7, 8, 10, 12
                    If IsNumeric(DiaTemp) Then If Val(DiaTemp) <= 31 And Val(DiaTemp) > 0 Then FechaTempValida = True
                Case 2
                    If IsNumeric(DiaTemp) Then
                        If Val(DiaTemp) <= 28 And Val(DiaTemp) > 0 Then
                            FechaTempValida = True
                        Else
                            If Val(DiaTemp) = 29 Then
                                If IsNumeric(AnyoTemp) Then
                                    If Val(AnyoTemp) Mod 4 = 0 Then
                                        FechaTempValida = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case 4, 6, 9, 11
                    If IsNumeric(DiaTemp) Then If Val(DiaTemp) <= 30 And Val(DiaTemp) > 0 Then FechaTempValida = True
            End Select
        End If
      End If
    End If
    If FechaTempValida Then
        .Text = DiaTemp & "/" & MesTemp & "/" & AnyoTemp
    Else
        .Text = vbNullString
    End If
End With

ExitSub:
    Exit Sub
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Sub
    End Select
End Sub

Public Function sNombrePais(ByVal iCodPais As Long, Optional ByVal bEnIngles As Boolean = False) As String

On Error GoTo ErrorFun

Select Case iCodPais
    Case giPAIS_ESPA—A
        sNombrePais = gsPAIS_ESPA—A
        If bEnIngles Then sNombrePais = gsPAIS_ESPA—A_Ingles
    Case giPAIS_ESPA—A2
        sNombrePais = gsPAIS_ESPA—A2
        If bEnIngles Then sNombrePais = gsPAIS_ESPA—A2_Ingles
    Case giPAIS_PORTUGAL
        sNombrePais = gsPAIS_PORTUGAL
        If bEnIngles Then sNombrePais = gsPAIS_PORTUGAL_Ingles
    Case giPAIS_ITALIA
        sNombrePais = gsPAIS_ITALIA
        If bEnIngles Then sNombrePais = gsPAIS_ITALIA_Ingles
    Case giPAIS_CHILE
        sNombrePais = gsPAIS_CHILE
        If bEnIngles Then sNombrePais = gsPAIS_CHILE_Ingles
    Case giPAIS_MEXICO
        sNombrePais = gsPAIS_MEXICO
        If bEnIngles Then sNombrePais = gsPAIS_MEXICO_Ingles
    Case giPAIS_PRICO
        sNombrePais = gsPAIS_PRICO
        If bEnIngles Then sNombrePais = gsPAIS_PRICO_Ingles
    Case giPAIS_BRASIL
        sNombrePais = gsPAIS_BRASIL
        If bEnIngles Then sNombrePais = gsPAIS_BRASIL_Ingles
    Case giPAIS_USA
        sNombrePais = gsPAIS_USA
        If bEnIngles Then sNombrePais = gsPAIS_USA_Ingles
    Case giPAIS_POLONIA
        sNombrePais = gsPAIS_POLONIA
        If bEnIngles Then sNombrePais = gsPAIS_POLONIA_Ingles
    Case giPAIS_REINOUNIDO
        sNombrePais = gsPAIS_REINOUNIDO
        If bEnIngles Then sNombrePais = gsPAIS_REINOUNIDO_Ingles
    Case giPAIS_CROACIA
        sNombrePais = gsPAIS_CROACIA
        If bEnIngles Then sNombrePais = gsPAIS_CROACIA_Ingles
    Case giPAIS_AUSTRIA
        sNombrePais = gsPAIS_AUSTRIA
        If bEnIngles Then sNombrePais = gsPAIS_AUSTRIA_Ingles
    Case giPAIS_ALEMANIA
        sNombrePais = gsPAIS_ALEMANIA
        If bEnIngles Then sNombrePais = gsPAIS_ALEMANIA_Ingles
    Case giPAIS_PANAMA
        sNombrePais = gsPAIS_PANAMA
        If bEnIngles Then sNombrePais = gsPAIS_PANAMA_Ingles
    Case giPAIS_CRICA
        sNombrePais = gsPAIS_CRICA
        If bEnIngles Then sNombrePais = gsPAIS_CRICA_Ingles
    Case giPAIS_MARRUECOS
        sNombrePais = gsPAIS_MARRUECOS
        If bEnIngles Then sNombrePais = gsPAIS_MARRUECOS_Ingles
    Case giPAIS_RCHECA
        sNombrePais = gsPAIS_RCHECA
        If bEnIngles Then sNombrePais = gsPAIS_RCHECA_Ingles
    Case giPAIS_VENEZUELA
        sNombrePais = gsPAIS_VENEZUELA
        If bEnIngles Then sNombrePais = gsPAIS_VENEZUELA_Ingles
    Case giPAIS_TUNEZ
        sNombrePais = gsPAIS_TUNEZ
        If bEnIngles Then sNombrePais = gsPAIS_TUNEZ_Ingles
    Case giPAIS_ARGENTINA
        sNombrePais = gsPAIS_ARGENTINA
        If bEnIngles Then sNombrePais = gsPAIS_ARGENTINA_Ingles
    Case giPAIS_COLOMBIA
        sNombrePais = gsPAIS_COLOMBIA
        If bEnIngles Then sNombrePais = gsPAIS_COLOMBIA_Ingles
    Case giPAIS_ANGOLA
        sNombrePais = gsPAIS_ANGOLA
        If bEnIngles Then sNombrePais = gsPAIS_ANGOLA_Ingles
    Case giPAIS_ARGELIA
        sNombrePais = gsPAIS_ARGELIA
        If bEnIngles Then sNombrePais = gsPAIS_ARGELIA_Ingles
    Case giPAIS_NIGERIA
        sNombrePais = gsPAIS_NIGERIA
        If bEnIngles Then sNombrePais = gsPAIS_NIGERIA_Ingles
    Case giPAIS_POLONIA
        sNombrePais = gsPAIS_POLONIA
        If bEnIngles Then sNombrePais = gsPAIS_POLONIA_Ingles
    Case giPAIS_HONGKONG
        sNombrePais = gsPAIS_HONGKONG
        If bEnIngles Then sNombrePais = gsPAIS_HONGKONG_Ingles
    Case giPAIS_RUMANIA
        sNombrePais = gsPAIS_RUMANIA
        If bEnIngles Then sNombrePais = gsPAIS_RUMANIA_Ingles
    Case giPAIS_TURQUIA
        sNombrePais = gsPAIS_TURQUIA
        If bEnIngles Then sNombrePais = gsPAIS_TURQUIA_Ingles
    Case giPAIS_REPDOM
        sNombrePais = gsPAIS_REPDOM
        If bEnIngles Then sNombrePais = gsPAIS_REPDOM_Ingles
    Case giPAIS_ECUADOR
        sNombrePais = gsPAIS_ECUADOR
        If bEnIngles Then sNombrePais = gsPAIS_ECUADOR_Ingles
    Case giPAIS_LIBIA
        sNombrePais = gsPAIS_LIBIA
        If bEnIngles Then sNombrePais = gsPAIS_LIBIA_Ingles
    Case giPAIS_PERU
        sNombrePais = gsPAIS_PERU
        If bEnIngles Then sNombrePais = gsPAIS_PERU_Ingles
    Case giPAIS_GUATEMALA
        sNombrePais = gsPAIS_GUATEMALA
        If bEnIngles Then sNombrePais = gsPAIS_GUATEMALA_Ingles
    Case giPAIS_COSTAMARFIL
        sNombrePais = gsPAIS_COSTAMARFIL
        If bEnIngles Then sNombrePais = gsPAIS_COSTAMARFIL_Ingles
    Case giPAIS_PERSEIDA
        sNombrePais = gsPAIS_PERSEIDA
        If bEnIngles Then sNombrePais = gsPAIS_PERSEIDA
    Case giPAIS_GASEXTREMADURA
        sNombrePais = gsPAIS_GASEXTREMADURA
        If bEnIngles Then sNombrePais = gsPAIS_GASEXTREMADURA_Ingles
    Case giPAIS_ONDUPACK
        sNombrePais = gsPAIS_ONDUPACK
         If bEnIngles Then sNombrePais = gsPAIS_ONDUPACK_Ingles
    Case giPAIS_ALTERENERSUN
        sNombrePais = gsPAIS_ALTERENERSUN
        If bEnIngles Then sNombrePais = gsPAIS_ALTERENERSUN_Ingles
    Case giPAIS_PLASTIVERD
        sNombrePais = gsPAIS_PLASTIVERD
        If bEnIngles Then sNombrePais = gsPAIS_PLASTIVERD_Ingles
    Case giPAIS_IQOXE
        sNombrePais = gsPAIS_IQOXE
        If bEnIngles Then sNombrePais = gsPAIS_IQOXE_Ingles
    Case giPAIS_GES
        sNombrePais = gsPAIS_GES
        If bEnIngles Then sNombrePais = gsPAIS_GES_Ingles
    Case giPAIS_LILOTECHNOLOGY
        sNombrePais = gsPAIS_LILOTECHNOLOGY
        If bEnIngles Then sNombrePais = gsPAIS_LILOTECHNOLOGY_Ingles
    Case giPAIS_ONDUPET
        sNombrePais = gsPAIS_ONDUPET
        If bEnIngles Then sNombrePais = gsPAIS_ONDUPET_Ingles
    Case giPAIS_MURCIACARTON
        sNombrePais = gsPAIS_MURCIACARTON
        If bEnIngles Then sNombrePais = gsPAIS_MURCIACARTON_Ingles
    Case giPAIS_MATIASGOMA
        sNombrePais = gsPAIS_MATIASGOMA
        If bEnIngles Then sNombrePais = gsPAIS_MATIASGOMA_Ingles
    Case giPAIS_AGSIDERURGICABALBOA
        sNombrePais = gsPAIS_AGSIDERURGICABALBOA
        If bEnIngles Then sNombrePais = gsPAIS_AGSIDERURGICABALBOA_Ingles
    Case giPAIS_GRUPOGALLARDOBALBOA
        sNombrePais = gsPAIS_GRUPOGALLARDOBALBOA
        If bEnIngles Then sNombrePais = gsPAIS_GRUPOGALLARDOBALBOA_Ingles
    Case giPAIS_ALFONSOGALLARDO
        sNombrePais = gsPAIS_ALFONSOGALLARDO
        If bEnIngles Then sNombrePais = gsPAIS_ALFONSOGALLARDO_Ingles
    Case giPAIS_CORRUGADOSLASAO
        sNombrePais = gsPAIS_CORRUGADOSLASAO
        If bEnIngles Then sNombrePais = gsPAIS_CORRUGADOSLASAO_Ingles
    Case giPAIS_MARCELIANOMARTIN
        sNombrePais = gsPAIS_MARCELIANOMARTIN
        If bEnIngles Then sNombrePais = gsPAIS_MARCELIANOMARTIN_Ingles
    Case giPAIS_FERROMALLAS
        sNombrePais = gsPAIS_FERROMALLAS
        If bEnIngles Then sNombrePais = gsPAIS_FERROMALLAS_Ingles
    Case giPAIS_CORRUGADOSGETAFE
        sNombrePais = gsPAIS_CORRUGADOSGETAFE
        If bEnIngles Then sNombrePais = gsPAIS_CORRUGADOSGETAFE_Ingles
    Case giPAIS_IQLIT
        sNombrePais = gsPAIS_IQLIT
        If bEnIngles Then sNombrePais = gsPAIS_IQLIT_Ingles
    Case giPAIS_GALVACOLOR
        sNombrePais = gsPAIS_GALVACOLOR
        If bEnIngles Then sNombrePais = gsPAIS_GALVACOLOR_Ingles
    Case giPAIS_CAMALECOSYSTEMS
        sNombrePais = gsPAIS_CAMALECOSYSTEMS
        If bEnIngles Then sNombrePais = gsPAIS_CAMALECOSYSTEMS_Ingles
    Case giPAIS_PAPELERASARLANZON
        sNombrePais = gsPAIS_PAPELERASARLANZON
        If bEnIngles Then sNombrePais = gsPAIS_PAPELERASARLANZON_Ingles
    Case giPAIS_FUNDACIONRL
        sNombrePais = gsPAIS_FUNDACIONRL
        If bEnIngles Then sNombrePais = gsPAIS_FUNDACIONRL_Ingles
    Case giPAIS_GASILUZ
        sNombrePais = gsPAIS_GASILUZ
        If bEnIngles Then sNombrePais = gsPAIS_GASILUZ_Ingles
    Case giPAIS_CRISTIANLAYDSS
        sNombrePais = gsPAIS_CRISTIANLAYDSS
        If bEnIngles Then sNombrePais = gsPAIS_CRISTIANLAYDSS_Ingles
    Case giPAIS_INDUSTRIASCL
        sNombrePais = gsPAIS_INDUSTRIASCL
        If bEnIngles Then sNombrePais = gsPAIS_INDUSTRIASCL_Ingles
    Case giPAIS_SOLARSTEEL
        sNombrePais = gsPAIS_SOLARSTEEL
        If bEnIngles Then sNombrePais = gsPAIS_SOLARSTEEL_Ingles
    Case giPAIS_CARTONAJESEXTREMADURA
        sNombrePais = gsPAIS_CARTONAJESEXTREMADURA
        If bEnIngles Then sNombrePais = gsPAIS_CARTONAJESEXTREMADURA_Ingles

    Case 999
        sNombrePais = "Nuevos Paises"
        If bEnIngles Then sNombrePais = "New Country's"
    Case 998
        sNombrePais = "Tienda"
        If bEnIngles Then sNombrePais = "Shop"
    Case Else
        sNombrePais = "Desconocido " & iCodPais
        If bEnIngles Then sNombrePais = "Unknown " & iCodPais
End Select
  
ExitFun:
    Exit Function
ErrorFun:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr$(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
        End If
    End With
    GoTo ExitFun
    Resume Next
End Function

Sub Limpia_Boxes(ParamArray Boxes())

Dim OBJETO

On Error GoTo ControlErrores: Const sNombreSub_ As String = "Limpia_Boxes"

For Each OBJETO In Boxes
    If TypeOf OBJETO Is Label Then
        OBJETO.Caption = vbNullString
    Else
        If TypeOf OBJETO Is TextBox Then
            OBJETO.Text = vbNullString
        Else
            OBJETO = vbNullString
        End If
    End If
Next

ExitSub:
    Exit Sub
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Sub
    End Select
End Sub

Public Function DC_BoletinesPR(ByVal Boletin As String) As Long

Dim iResto As Long
Dim iTotal As Long
Dim iMultiplicador As Long

On Error GoTo ErrorFun

If Not IsNumeric(Boletin) Then
    MsgBox "El boletÌn debe ser numÈrico para calcular su dÌgito de control.", vbApplicationModal
    Exit Function
End If

iTotal = 0
iMultiplicador = 2
While Len(Boletin) > 1
    iResto = Right$(Boletin, 1)
    iTotal = iTotal + iResto * iMultiplicador
    If iMultiplicador = 2 Then iMultiplicador = 1 Else iMultiplicador = 2
    Boletin = Left$(Boletin, Len(Boletin) - 1)
Wend

iResto = Right$(iTotal, 1)

If iResto = 0 Then iResto = 10
DC_BoletinesPR = 10 - iResto

ExitFun:
    Exit Function
ErrorFun:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr$(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
        End If
    End With
    GoTo ExitFun
    Resume Next
End Function

Public Function DC_Distri(ByVal Distri As Long) As Long

Dim Codigo As String
Dim Codi, Ceros As String * 6
Dim Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Aux1, DIGITO As Long
    
On Error GoTo ErrorFun

Ceros = "000000"
Codigo = Str$(Distri)
Codi = Format$(Codigo, Ceros)
    
Valor1 = Mid$(Codi, 1, 1) * 7
Valor2 = Mid$(Codi, 2, 1) * 2
Valor3 = Mid$(Codi, 3, 1) * 3
Valor4 = Mid$(Codi, 4, 1) * 4
Valor5 = Mid$(Codi, 5, 1) * 5
Valor6 = Mid$(Codi, 6, 1) * 6
Aux1 = Valor1 + Valor2 + Valor3 + Valor4 + Valor5 + Valor6
    
Codi = Mid$(Ceros, 1, (3 - (Len(Aux1)))) & Aux1
         
Aux1 = Val(Mid$(Codi, 1, 1)) + Val(Mid$(Codi, 2, 1)) + Val(Mid$(Codi, 3, 1))
DIGITO = Mid$(Aux1, Len(Aux1), 1)
DC_Distri = DIGITO
  
ExitFun:
    Exit Function
ErrorFun:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr$(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
        End If
    End With
    GoTo ExitFun
    Resume Next
End Function

Sub EntraNumero(KeyAscii As Integer, Control As Control, Enteros As Long, Optional Decimales As Long = 0, Optional PuntuaciÛn As String = vbNullString, Optional Signo As String = vbNullString)
    Dim Punto As Long
    
    If TypeOf Control Is TextBox Or _
       TypeOf Control Is ComboBox _
    Then
    Else
        Exit Sub
    End If
    
If KeyAscii = 8 Then Exit Sub
If Control.Locked Then Exit Sub
If KeyAscii = 27 Then
    Control.Text = vbNullString
    Exit Sub
End If
If Not (KeyAscii > 47 And KeyAscii < 58) And _
   Not KeyAscii = 45 And _
   Not KeyAscii = 13 And _
   Not KeyAscii = 46 Then
    
    KeyAscii = 17
    Exit Sub
End If

If KeyAscii = 13 Or KeyAscii = 45 Then
    If KeyAscii = 45 Then
        If Signo = "S" Then
            If Mid$(Control.Text, 1, 1) = "-" Then
                If Len(Control.Text) > 1 Then
                    Control.Text = Mid$(Control.Text, 2)
                Else
                    Control.Text = vbNullString
                End If
            Else
                Control.Text = "-" + Control.Text
            End If
        Else
            KeyAscii = 17
            Exit Sub
        End If
    End If
    Control.Tag = Control.Text
    If Len(PuntuaciÛn) <> 0 And Control.Text <> "-" And Len(Control.Text) <> 0 Then
               
        Control.Text = Format$(Control.Tag, PuntuaciÛn)
        
    End If
    If Enteros > 0 Then
        If KeyAscii = 13 Then
            SendKeys "{TAB}"
            DoEvents
        End If
        KeyAscii = 0
        Exit Sub
    End If
    KeyAscii = 0
Else
    If KeyAscii = 46 Then
        If Decimales = 0 Or _
           Len(Control.Text) = 0 Or _
           InStr(1, Control.Text, ".") Then
            KeyAscii = 17
            Exit Sub
        End If
    End If
    Punto = InStr(1, Control.Text, ".")
    If Control.SelLength = 0 Then
        If Punto = 0 Then
            If Len(Control.Text) >= Abs(Enteros) Then
                KeyAscii = 17
            End If
        Else
            If Punto - 1 > Abs(Enteros) Then
                KeyAscii = 17
            End If
            If Len(Control.Text) - Punto >= Decimales Then
                KeyAscii = 17
            End If
        End If
    End If
End If

End Sub

Sub EntraTexto(KeyAscii As Integer, Control As Control, Longitud As Long, CaracteresValidos As String, Optional Mayusculas As Boolean = False)
    
Dim Caracter As String * 1
Dim CaracterEncontrado As Boolean
Dim ContadorCaracteres As Long
Dim ContadorCaracteresHasta As Long
    
    If TypeOf Control Is TextBox _
    Then
    Else
        Exit Sub
    End If
    
On Error GoTo ControlErrores: Const sNombreSub_ As String = "EntraTexto"

If KeyAscii = 8 Then Exit Sub
If KeyAscii = 27 Then
    Control.Text = vbNullString
    Exit Sub
End If

If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
    DoEvents
Else
    If Control.SelLength = 0 Then
        If Len(Control.Text) >= Longitud Then
            KeyAscii = 17
        End If
    End If
End If

If Len(CaracteresValidos) <> 0 Then
    CaracterEncontrado = False
    ContadorCaracteresHasta = Len(CaracteresValidos)
    For ContadorCaracteres = 1 To ContadorCaracteresHasta
        Caracter = Mid$(CaracteresValidos, ContadorCaracteres, 1)
        If KeyAscii = Asc(Caracter) Then
            CaracterEncontrado = True
            Exit For
        End If
    Next
    If Not CaracterEncontrado Then
        KeyAscii = 17
        Exit Sub
    End If
End If

If Mayusculas Then
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End If

ExitSub:
    Exit Sub
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Sub
    End Select
End Sub

Function Es_Vacio(ParamArray Boxes()) As Variant

Dim Encontrado As Boolean
Dim OBJETO, ObjetoEncontrado, ObjetoAnterior

On Error GoTo ErrorFun

Encontrado = False

For Each OBJETO In Boxes
    If TypeOf OBJETO Is TextBox Then
        If IsNull(OBJETO.Text) Then
            Encontrado = True
            Set ObjetoEncontrado = OBJETO
            Exit For
        Else
            If Len(OBJETO.Text) = 0 Then
                Encontrado = True
                Set ObjetoEncontrado = OBJETO
                Exit For
            End If
        End If
    Else
        If TypeOf OBJETO Is Label Then
            If Len(OBJETO.Caption) = 0 Then
                Encontrado = True
                Set ObjetoEncontrado = OBJETO
                Exit For
            End If
        Else
            If IsNull(OBJETO) Then
                Encontrado = True
                Set ObjetoEncontrado = OBJETO
                Exit For
            End If
        End If
    End If
Next

If Encontrado Then
    If ObjetoEncontrado.Enabled And Not (TypeOf OBJETO Is Label) Then
        Set Es_Vacio = ObjetoEncontrado
    Else
        Set ObjetoAnterior = OBJETO
        Set OBJETO = Nothing
        For Each OBJETO In Boxes
            If OBJETO <> ObjetoEncontrado Then
                If OBJETO.Enabled And Not (TypeOf OBJETO Is Label) Then Set ObjetoAnterior = OBJETO
            Else
                Set Es_Vacio = ObjetoAnterior
                Exit For
            End If
        Next
    End If
Else
    Set Es_Vacio = Nothing
End If

ExitFun:
    Exit Function
ErrorFun:
    With Err
        If MsgBox(.Number & " - " & .Description & Format$(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
        End If
    End With
    GoTo ExitFun
    Resume Next
End Function

Sub Baja_Linea(Linea As Variant, ParamArray Boxes() As Variant)

Dim campo
Dim Contador As Long

On Error GoTo ControlErrores: Const sNombreSub_ As String = "Baja_Linea"

Contador = 0
For Each campo In Linea
    If TypeOf Boxes(Contador) Is Label Then
        Boxes(Contador).Caption = campo
    Else
        If TypeOf Boxes(Contador) Is TextBox Then
            Boxes(Contador).Text = campo
        Else
            Boxes(Contador) = campo
        End If
    End If
    Contador = Contador + 1
Next

ExitSub:
    Exit Sub
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Sub
    End Select
End Sub

Sub Concatenar(ByVal Origen1 As String, ByVal Origen2 As String, ByVal Destino As String)

Dim buffer() As Byte
Dim numOrigen1, numOrigen2, numDestino As Long

On Error GoTo ControlErrores: Const sNombreSub_ As String = "Concatenar"

numOrigen1 = FreeFile
Open Origen1 For Binary Access Read As #numOrigen1
numOrigen2 = FreeFile
Open Origen2 For Binary Access Read As #numOrigen2
numDestino = FreeFile
Open Destino For Binary Access Write As #numDestino

ReDim buffer(1 To LOF(1))
Get #numOrigen1, , buffer
Put #numDestino, , buffer
ReDim buffer(1 To LOF(2))
Get #numOrigen2, , buffer
Put #numDestino, , buffer
Close #numOrigen1, #numOrigen2, #numDestino

ExitSub:
    Exit Sub
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Sub
    End Select
End Sub

Public Function Semana_Del_Anyo(ByVal QueDia As Date) As String

Dim Dia1Anyo As Date
Dim FechaQueDia As Date
Dim stFechaQueDia As String
Dim stDia1Anyo As String
Dim DiaDelAnyo As Long

On Error GoTo ControlErrores: Const sNombreSub_ As String = "Semana_Del_Anyo"

Semana_Del_Anyo = vbNullString
If Len(QueDia) <> 0 And Not IsNull(QueDia) Then
    If IsDate(QueDia) Then
        stFechaQueDia = Format$(QueDia, "dd/mm/yyyy")
        FechaQueDia = stFechaQueDia
        stDia1Anyo = Format$("01/01/" & Year(QueDia), "dd/mm/yyyy")
        Dia1Anyo = stDia1Anyo
        DiaDelAnyo = FechaQueDia - Dia1Anyo
        Select Case Weekday(Dia1Anyo, vbSunday)
            Case 1 'Domingo
                DiaDelAnyo = DiaDelAnyo + 7
            Case 2 'Lunes
                DiaDelAnyo = DiaDelAnyo + 1
            Case 3 'Martes
                DiaDelAnyo = DiaDelAnyo + 2
            Case 4 'Miercoles
                DiaDelAnyo = DiaDelAnyo + 3
            Case 5 'Jueves
                DiaDelAnyo = DiaDelAnyo + 4
            Case 6 'Viernes
                DiaDelAnyo = DiaDelAnyo + 5
            Case 7 'Sabado
                DiaDelAnyo = DiaDelAnyo + 6
        End Select
        If DiaDelAnyo <= 7 Then
            Semana_Del_Anyo = "1"
        Else
            Semana_Del_Anyo = Str$(DiaDelAnyo \ 7)
        End If
    End If
End If

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Function DiasEnMes(ByVal QueFecha As Date) As Long

Dim Mes As Long
Dim Anyo As Long

On Error GoTo ControlErrores: Const sNombreSub_ As String = "DiasEnMes"

Mes = Month(QueFecha)
Anyo = Year(QueFecha)
If Mes = 2 Then
    If (Anyo Mod 4) Then
        DiasEnMes = 28
    Else
        DiasEnMes = 29
    End If
Else
    If Mes = 4 Or Mes = 6 Or Mes = 9 Or Mes = 11 Then
        DiasEnMes = 30
    Else
        DiasEnMes = 31
    End If
End If

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Public Function ObtenerSeparadorMiles() As String

Dim FmtData As String * 1
Dim ret As Long

ret = GetLocaleInfo(&H400, &HF, FmtData, 25)
ObtenerSeparadorMiles = FmtData

End Function

Function ObtenerSeparadorDecimal() As String

Dim FmtData As String * 1
Dim ret As Long

ret = GetLocaleInfo(&H400, &HE, FmtData, 25)
ObtenerSeparadorDecimal = FmtData

End Function

Function FNumero(ByVal Numero, Optional ByVal Decimales = 0) As String

Dim FormatoDecimales As String
Dim ContadorDecimales As Long

On Error GoTo Err_FNumero

If IsNull(Numero) Then FNumero = vbNullString: Exit Function
If Len(Trim$(Numero)) = 0 Then FNumero = vbNullString: Exit Function

FormatoDecimales = vbNullString
For ContadorDecimales = 1 To Decimales
    FormatoDecimales = FormatoDecimales & "0"
Next

If Len(FormatoDecimales) = 0 Then
    FormatoDecimales = "#" & ObtenerSeparadorMiles & "##0"
Else
    FormatoDecimales = "#" & ObtenerSeparadorMiles & "##0" & ObtenerSeparadorDecimal & FormatoDecimales
End If
FNumero = Format$(Val(Numero), FormatoDecimales)

Exit_FNumero:
    Exit Function
Err_FNumero:
    MsgBox Err.Description
    Resume Exit_FNumero
    Resume Next
End Function

'Function QuitarPonerCampanyas_Id(ByVal Cam_Id_Inicio As Long, ByVal CampanyasASumar As Long) As Long
'
'Dim AnyoInicio As Long, CampanyaInicio As Long, AnyoResultante As Long, CampanyaResultante As Long
'Dim iContadorCampanyas As Long, iContadorCampanyasHasta As Long
'
'On Error GoTo ControlErrores: Const sNombreSub_ As String = "QuitarPonerCampanyas_Id"
'
'AnyoInicio = Left$(Cam_Id_Inicio, 4)
'CampanyaInicio = Right$(Cam_Id_Inicio, 2)
'
'AnyoResultante = AnyoInicio
'CampanyaResultante = CampanyaInicio
'iContadorCampanyasHasta = Abs(CampanyasASumar)
'For iContadorCampanyas = 1 To iContadorCampanyasHasta
'    If CampanyasASumar > 0 Then
'        CampanyaResultante = CampanyaResultante + 1
'        If CampanyaResultante > MaxCampanyaAnterior Then
'            CampanyaResultante = 1
'            AnyoResultante = AnyoResultante + 1
'        End If
'    ElseIf CampanyasASumar < 0 Then
'        CampanyaResultante = CampanyaResultante - 1
'        If CampanyaResultante < 1 Then
'            CampanyaResultante = MaxCampanyaAnterior
'            AnyoResultante = AnyoResultante - 1
'        End If
'    End If
'Next
'
'QuitarPonerCampanyas_Id = AnyoResultante * 100 + CampanyaResultante
'
'ExitSub:
'    Exit Function
'ControlErrores:
'    MsgBox Err.Number & "-" & Err.Description & "-" & Err.Source
'    If IsDebug Then Stop
'    Err.Clear
'    Resume Next
'    Resume
'    Exit Function
'End Function

'Sub QuitarPonerCampanyas(ByVal AnyoInicio As Long, ByVal CampanyaInicio As Long, ByVal CampanyasASumar As Long, ByRef AnyoResultante As Long, ByRef CampanyaResultante As Long)
'
'Dim iContadorCampanyas As Long, iContadorCampanyasHasta As Long
'
'On Error GoTo ControlErrores: Const sNombreSub_ As String = "QuitarPonerCampanyas"
'
'AnyoResultante = AnyoInicio
'CampanyaResultante = CampanyaInicio
'iContadorCampanyasHasta = Abs(CampanyasASumar)
'For iContadorCampanyas = 1 To iContadorCampanyasHasta
'    If CampanyasASumar > 0 Then
'        CampanyaResultante = CampanyaResultante + 1
'        If CampanyaResultante > MaxCampanyaAnterior Then
'            CampanyaResultante = 1
'            AnyoResultante = AnyoResultante + 1
'        End If
'    ElseIf CampanyasASumar < 0 Then
'        CampanyaResultante = CampanyaResultante - 1
'        If CampanyaResultante < 1 Then
'            CampanyaResultante = MaxCampanyaAnterior
'            AnyoResultante = AnyoResultante - 1
'        End If
'    End If
'Next
'
'ExitSub:
'    Exit Sub
'ControlErrores:
'    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
'        Case vbIgnore
'            Err.Clear
'            Resume Next
'        Case vbRetry
'            Err.Clear
'            Resume
'        Case vbAbort
'            Err.Clear
'            Exit Sub
'    End Select
'End Sub

Function QuitarCaracteres(ByVal QueCadena As String, ByVal QueQuitar As String, Optional ByVal QuePoner As String = vbNullString) As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "QuitarCaracteres"

If InStr(1, QueCadena, QueQuitar) > 0 Then
    QuitarCaracteres = Replace(QueCadena, QueQuitar, QuePoner)
Else
    QuitarCaracteres = QueCadena
End If

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

'   Retorna valor no nulo dependiendo del tipo de dato
Function NoNull(ByVal ArVar, Optional ByVal arTipo As String = "N") As Variant

    If Len(ArVar) > 0 Then
        Select Case arTipo
          Case "A"
            NoNull = ArVar
          Case "N"
            NoNull = Val(ArVar)
          Case "F"
            If ArVar = "00/00/0000" Then
                NoNull = vbNullString
            Else
                If InStr(ArVar, "/") > 0 Then
                    NoNull = ArVar
                Else
                    NoNull = Format$(ArVar, "00/00/0000")
                End If
            End If
          Case Else
            NoNull = ArVar
        End Select
    Else
        Select Case arTipo
          Case "A"
            NoNull = vbNullString
          Case "N"
            NoNull = 0
          Case "F"
            NoNull = CDate("00:00:0000")
          Case Else
            NoNull = vbNullString
        End Select
    End If
End Function

Function Operar(Valor1, QueOperacion As Operaciones, Valor2) As Variant
   
Const mnErrDivByZero = 11, mnErrOverFlow = 6, mnErrBadCall = 5, mnErrType = 13

On Error GoTo ErrorDivision

Select Case QueOperacion
    Case Sumar
        Operar = CDbl(Valor1) + CDbl(Valor2)
    Case Restar
        Operar = Valor1 - Valor2
    Case Multiplicar
        Operar = Valor1 * Valor2
    Case Dividir
        Operar = Valor1 / Valor2
    Case Else
        Stop
End Select

Exit Function

ErrorDivision:
   If Err.Number = mnErrDivByZero Or _
      Err.Number = mnErrOverFlow Or _
      Err.Number = mnErrType Or _
      Err.Number = mnErrBadCall Then
        Operar = Null     ' Si el error era DivisiÛn por cero,
                           ' Desbordamiento o Llamada no v·lida
                           ' a un procedimiento, devuelve Null.
   Else
        If MsgBox(Err.Number & " - " & Err.Description & Chr$(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & Err.Source) = vbYes Then
            Stop
            Resume Next
        End If
   End If
   Resume Next
End Function

Function F2(ByVal QueFormateo As String) As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "F2"

If IsNull(QueFormateo) Then QueFormateo = vbNullString

F2 = Format$(QueFormateo, "00")

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Function F3(ByVal QueFormateo As String) As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "F3"

If IsNull(QueFormateo) Then QueFormateo = vbNullString

F3 = Format$(QueFormateo, "000")

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Function F4(ByVal QueFormateo As String) As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "F4"

If IsNull(QueFormateo) Then QueFormateo = vbNullString

F4 = Format$(QueFormateo, "0000")

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Function F5(ByVal QueFormateo As String) As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "F5"

If IsNull(QueFormateo) Then QueFormateo = vbNullString

F5 = Format$(QueFormateo, "00000")

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Function F6(ByVal QueFormateo As String) As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "F6"

If IsNull(QueFormateo) Then QueFormateo = vbNullString

F6 = Format$(QueFormateo, "000000")

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Function F7(ByVal QueFormateo As String) As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "F7"

If IsNull(QueFormateo) Then QueFormateo = vbNullString

F7 = Format$(QueFormateo, "0000000")

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Function F8(ByVal QueFormateo As String) As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "F8"

If IsNull(QueFormateo) Then QueFormateo = vbNullString

F8 = Format$(QueFormateo, "00000000")

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Function FAlinearALaDerecha(ByVal QueFormateo As String, ByVal LongitudDeseada As Long) As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "FAlinearALaDerecha"

If IsNull(QueFormateo) Then QueFormateo = vbNullString
If Len(QueFormateo) <= LongitudDeseada Then
    FAlinearALaDerecha = Right$(Space$(LongitudDeseada) & QueFormateo, LongitudDeseada)
Else
    FAlinearALaDerecha = QueFormateo
End If

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Public Function dConvierteMoneda(ByVal sValor As String, ByVal yDireccion As DireccionCambio, _
     ByVal yCodPais As Long) As Double
    
On Error GoTo ControlErrores: Const sNombreSub_ As String = "dConvierteMoneda"

    Const dCambioMoneda_Espanya As Double = 166.386
    Const dCambioMoneda_Italia As Double = 1936.27
    Const dCambioMoneda_Portugal As Double = 200.482
    
    Dim dValor As Double
    Dim dCambio As Double

    'sÌ yDireccion es 0 pasamos de pesetas a euros, y sÌ es 1 de euros a pesetas
    
    Select Case yCodPais
    Case 34
        dCambio = dCambioMoneda_Espanya
    Case 39
        dCambio = dCambioMoneda_Italia
    Case giPAIS_PORTUGAL
        dCambio = dCambioMoneda_Portugal
    Case Else
        dCambio = 1
    End Select
    
    If Len(sValor) = 0 Or sValor = Null Then sValor = 0
    dValor = CDbl(sValor)
    
    If yDireccion = 0 Then
        dConvierteMoneda = Format$(Round(Val(Format$(dValor / dCambio, "#.000")), 2), "#.00")
    Else
        dConvierteMoneda = dRedondea(CStr(dValor * dCambio))
    End If

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Public Function dRedondea(ByVal sCadena As String) As Double

On Error GoTo ControlErrores: Const sNombreSub_ As String = "dRedondea"

    Dim Punto As Long
    Dim sCad
    
    sCad = sCadena
   
    sCad = Format$(sCad, "###########0.0")
    Punto = InStr(1, sCad, ".")
    If Punto > 0 Then
          sCad = Format$(sCad, "###########0")
    End If
    
    dRedondea = CDbl(sCad)

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Function DameColeccionCadena(ByVal QueCadena As String) As Collection

Dim CadenaTemp As String
Dim DondeComa As Long
Dim ColeccionTemporal As New Collection

On Error GoTo MirarError

Do
    DondeComa = InStr(1, QueCadena, ",")
    If DondeComa <> 0 Then
        CadenaTemp = Mid$(QueCadena, 1, DondeComa - 1)
        QueCadena = Mid$(QueCadena, DondeComa + 1)
    Else
        CadenaTemp = QueCadena
    End If
    ColeccionTemporal.Add CadenaTemp
Loop While DondeComa <> 0

Set DameColeccionCadena = ColeccionTemporal

ExitSub:
    Exit Function
MirarError:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr$(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
            Resume
        End If
    End With
    GoTo ExitSub
End Function

Function DameCadenaColeccion(QueObjeto, Optional ByVal ConStrings As Boolean = False, Optional EntreComillas As String = vbNullString) As String

Dim CadenaTemp As String
Dim OBJETO

On Error GoTo MirarError

    CadenaTemp = vbNullString
    
    If TypeOf QueObjeto Is VBA.Collection Then
        For Each OBJETO In QueObjeto
            CadenaTemp = CadenaTemp & "," & DameCadenaColeccion(OBJETO, ConStrings, EntreComillas)
        Next
    Else
        If IsNumeric(QueObjeto) Or ConStrings Then
            CadenaTemp = CadenaTemp & "," & EntreComillas & QueObjeto & EntreComillas
        Else
            CadenaTemp = CadenaTemp & "," & EntreComillas & "0" & EntreComillas
        End If
    End If
    If Len(CadenaTemp) <> 0 Then CadenaTemp = Mid$(CadenaTemp, 2)
    DameCadenaColeccion = CadenaTemp

ExitSub:
    Exit Function
MirarError:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr$(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
            Resume
        End If
    End With
    GoTo ExitSub
End Function

Public Function CapturarError(ErrNumber As Long, ErrDescription As String, ErrSource As String, ErrSub As String, ErrModulo As String, Optional ByVal ErrLinea As Long = 0, Optional ByVal TextoAFichero As String = "")

    Dim QueBoton
    Dim sMensaje As String

    sMensaje = "Se ha producido el siguiente error:" & vbCrLf & vbCrLf & _
               vbTab & "Origen:       " & vbTab & ErrSource & vbCrLf & _
               vbTab & "Modulo:       " & vbTab & ErrModulo & vbCrLf & _
               vbTab & "Sub/Function: " & vbTab & ErrSub & vbCrLf
    If ErrLinea <> 0 Then
    sMensaje = sMensaje & _
               vbTab & "Linea Error: " & vbTab & ErrLinea & vbCrLf
    End If
    sMensaje = sMensaje & _
               vbTab & "N∫ error:     " & vbTab & ErrNumber & vbCrLf & _
               vbTab & "DescripciÛn:  " & vbTab & ErrDescription & vbCrLf & vbCrLf & _
               "-----------------------------------------------------" & vbCrLf & _
               "Pulsar ... " & vbCrLf & _
               vbTab & "'Anular' para salir del mÛdulo con error (Exit Sub)" & vbCrLf & _
               vbTab & "'Reitentar' para repetir la lÌnea del error (Resume)" & vbCrLf & _
               vbTab & "'Omitir' para saltarse la lÌnea del error (Resume Next)"
    QueBoton = MsgBox(sMensaje, vbAbortRetryIgnore + vbCritical + vbDefaultButton2, "ERROR GRAVE")
    CapturarError = QueBoton
    
'    If TextoAFichero <> "" Then
'        Dim sHora As String
'        Dim FicheroLibre As Long
'
'        On Error Resume Next
'
'        sHora = Format$(Now, "hh:mm:ss")
'        FicheroLibre = FreeFile
'        Open Ruta & "LOGS\_Errores\" & Format$(Date, "yyyy-mm-dd") & "_" & ErrSub & ".txt" For Append As #FicheroLibre
'        Print #FicheroLibre, sHora & vbCrLf; sMensaje & vbCrLf & TextoAFichero & vbCrLf & vbCrLf
'        Close #FicheroLibre
'    End If

End Function

Public Sub msgErrorInsert(ByVal QueInsert As String)
MsgBox "ERROR AL INSERTAR" & vbCrLf & "---------------------------" & vbCrLf & QueInsert, vbApplicationModal + vbCritical, "ERROR AL INSERTAR"
End Sub

Public Sub msgErrorUpdate(ByVal QueUpdate As String)
MsgBox "ERROR AL ACTUALIZAR" & vbCrLf & "---------------------------" & vbCrLf & QueUpdate, vbApplicationModal + vbCritical, "ERROR AL ACTUALIZAR"
End Sub

Public Sub AQUITAR(Optional ByVal sMensaje As String = vbNullString)
If Len(sMensaje) <> 0 Then
    MsgBox sMensaje, vbApplicationModal + vbCritical, "PARADA PROVOCADA"
    Stop
End If
End Sub

Public Sub ATRADUCIR(Optional ByVal sMensaje As String = vbNullString)
If Len(sMensaje) <> 0 Then
    MsgBox sMensaje, vbApplicationModal + vbCritical, "PARADA PROVOCADA"
    Stop
End If
End Sub

Public Function sDameNombrePC()

Dim buffer As String
Dim estado As Long

On Error GoTo MirarError

buffer = String$(255, " ")
estado = GetComputerName(buffer, 255)
If estado <> 0 Then
    sDameNombrePC = Trim$(Left$(buffer, 255))
    sDameNombrePC = QuitarCaracteres(sDameNombrePC, Chr$(0), vbNullString)
    sDameNombrePC = UCase$(sDameNombrePC)
End If
If Len(Trim$(sDameNombrePC)) = 0 Then MsgBox "Error al recuperar el nombre del PC.", vbApplicationModal

ExitSub:
    Exit Function
MirarError:
    With Err
        If MsgBox(.Number & " - " & .Description & Chr$(13) & _
                  "Desea parar el programa?", _
                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
                  "ERROR EN " & .Source) = vbYes Then
            Stop
            Resume
        End If
    End With
    GoTo ExitSub
End Function

Sub VaciarCajasTexto(ByVal Formulario As Form, ByVal Contenedor As Control)

Dim OBJETO

On Error GoTo ControlErrores: Const sNombreSub_ As String = "VaciarCajas"

For Each OBJETO In Formulario.Controls
    If TypeOf OBJETO Is TextBox Then
        If OBJETO.Container Is Contenedor Then
            OBJETO.Text = vbNullString
        End If
    End If
Next

ExitSub:
    Exit Sub
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Sub
    End Select
End Sub

Function FormatearTexto(ByVal QueFormatear As String, ByVal CuantosCaracteres As Long, Optional ByVal Alinear As String = "D") As String

On Error GoTo ControlErrores: Const sNombreSub_ As String = "FormatearTexto"

If Alinear = "D" Then
    FormatearTexto = Right$(Space$(100) & Trim$(QueFormatear), CuantosCaracteres)
Else
    FormatearTexto = Left$(Trim$(QueFormatear) & Space$(100), CuantosCaracteres)
End If
FormatearTexto = QuitarCaracteres(FormatearTexto, " ", "†")

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Function Espacios(CuantosEspacios As Long)

Dim iContador As Long

On Error GoTo ControlErrores: Const sNombreSub_ As String = "Espacios"

Espacios = vbNullString
For iContador = 1 To CuantosEspacios
    Espacios = Espacios & "†"
Next

ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

'Sub QuitarPonerCampanyas(ByVal AnyoInicio As Long, ByVal CampanyaInicio As Long, ByVal CampanyasASumar As Long, ByRef AnyoResultante As Long, ByRef CampanyaResultante As Long)
'
'Dim iContadorCampanyas As Long, iContadorCampanyasHasta As Long
'
'On Error GoTo ControlErrores: Const sNombreSub_ As String = "QuitarPonerCampanyas"
'
'AnyoResultante = AnyoInicio
'CampanyaResultante = CampanyaInicio
'iContadorCampanyasHasta = Abs(CampanyasASumar)
'For iContadorCampanyas = 1 To iContadorCampanyasHasta
'    If CampanyasASumar > 0 Then
'        CampanyaResultante = CampanyaResultante + 1
'        If CampanyaResultante > MaxCampanyaAnterior Then
'            CampanyaResultante = 1
'            AnyoResultante = AnyoResultante + 1
'        End If
'    ElseIf CampanyasASumar < 0 Then
'        CampanyaResultante = CampanyaResultante - 1
'        If CampanyaResultante < 1 Then
'            CampanyaResultante = MaxCampanyaAnterior
'            AnyoResultante = AnyoResultante - 1
'        End If
'    End If
'Next
'
'ExitSub:
'    Exit Sub
'ControlErrores:
'    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
'        Case vbIgnore
'            Err.Clear
'            Resume Next
'        Case vbRetry
'            Err.Clear
'            Resume
'        Case vbAbort
'            Err.Clear
'            Exit Sub
'    End Select
'End Sub

Public Sub Tabular()
    SendKeys "{TAB}"
    DoEvents
End Sub

Sub TablaPrueba(ByVal rsTablaPrueba As ADODB.Recordset)

Dim rsCopia As ADODB.Recordset

If rsTablaPrueba Is Nothing Then Exit Sub
Set rsCopia = rsTablaPrueba.Clone
If Not rsCopia.EOF Then
    Dim objField
    Dim sCadena As String
    
    sCadena = vbNullString
    For Each objField In rsCopia.Fields
        sCadena = sCadena & objField.Name & vbTab
    Next
    Debug.Print sCadena
    
    While Not rsCopia.EOF
        sCadena = vbNullString
        For Each objField In rsTablaPrueba.Fields
            sCadena = sCadena & rsCopia(objField.Name) & vbTab
        Next
        Debug.Print sCadena
        
        rsCopia.MoveNext
    Wend
End If
End Sub

Function sCadenaAleatoria(ByVal CuantosCaracteres As Long, Optional ByVal sCaracteresValidos As String = "abcdefghijklmnopqrstuvwxyz") As String

Dim iContador As Long

On Error GoTo ControlErrores: Const sNombreSub_ As String = "sCadenaAleatoria"

Randomize
sCadenaAleatoria = vbNullString
For iContador = 1 To CuantosCaracteres
    sCadenaAleatoria = sCadenaAleatoria & Mid$(sCaracteresValidos, Int(Len(sCaracteresValidos) * Rnd + 1), 1)
Next
sCadenaAleatoria = LCase$(sCadenaAleatoria)
ExitSub:
    Exit Function
ControlErrores:
    Select Case CapturarError(Err.Number, Err.Description, Err.Source, sNombreSub_, sModulo_)
        Case vbIgnore
            Err.Clear
            Resume Next
        Case vbRetry
            Err.Clear
            Resume
        Case vbAbort
            Err.Clear
            Exit Function
    End Select
End Function

Function RoundUp(ByVal Numero As Double) As Long
RoundUp = Int(Numero) + 1
End Function

Function RoundDown(ByVal Numero As Double) As Long
RoundDown = Int(Numero)
End Function

Public Function IsDebug() As Boolean
  On Error Resume Next
  Debug.Assert 1 / 0
  IsDebug = (Err <> 0)
End Function

Public Sub AcumularImporte(ByRef sValorAcumulado As String, ByVal sValorASumar As String)
If IsNumeric(sValorASumar) Then
    If IsNumeric(sValorAcumulado) Then
        sValorAcumulado = CDbl(sValorAcumulado) + CDbl(sValorASumar)
    Else
        sValorAcumulado = sValorASumar
    End If
End If
End Sub

Public Sub CerrarPrograma()
On Error Resume Next
CloseForms
End
End Sub

Private Sub CloseForms()
Dim Formu
On Error Resume Next
'If have modal forms open
For Each Formu In Forms
    Unload Formu
    Set Formu = Nothing
Next
End Sub

Function DarRutaCorrecta_CMI(RutCopia As String, iCodPais As Integer) As String
'C:\CMI-Producto\

    Dim sDirectorio As String
    Dim bEnc As Boolean
    
    Dim sRutaEquipo As String
    
    On Error GoTo Error_Rutina
    
    sRutaEquipo = "C:\"
    'If goUsuario.PaisTrabajo <> giPAIS_ESPA—A Then sRutaEquipo = "V:\"
    If iCodPais <> giPAIS_ESPA—A Then sRutaEquipo = "V:\"
    
    bEnc = False
    If Mid(RutCopia, 1, 1) = ":" Then
        sDirectorio = Dir(Mid(sRutaEquipo, 1, 1) & RutCopia, vbDirectory)
    Else
        sDirectorio = Dir(sRutaEquipo & RutCopia, vbDirectory)
    End If
    Do While sDirectorio <> ""
        
        If Mid(RutCopia, 1, 1) = ":" Then
            If LCase(sDirectorio) = LCase(Mid(RutCopia, 3)) Then bEnc = True
        Else
            If LCase(sDirectorio) = LCase(RutCopia) Then bEnc = True
        End If
        sDirectorio = Dir
        
    Loop
    
    If Not bEnc Then MkDir sRutaEquipo & RutCopia
    
    If Mid(RutCopia, 1, 1) = ":" Then
        DarRutaCorrecta_CMI = Mid(sRutaEquipo, 1, 1) & RutCopia & "\"
    Else
        DarRutaCorrecta_CMI = sRutaEquipo & RutCopia
    End If
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Function

'Function DarRutaCorrecta(RutCopia As String) As String
'Dim Fso As New FileSystemObject
'Dim Drv As Drive
'Dim RutaCopia, RutaCopiaTemp As String
''Dim Fso, Drv
'
'On Error GoTo ErrorSub
'RutaCopia = RutCopia '":\DistribDesarrollo"
''definida como global arriba porque fallaba
'
'Set Fso = CreateObject("Scripting.FileSystemObject")
'On Error Resume Next
'Set Drv = Fso.GetDrive(Fso.GetDriveName("V:"))
'If Err.Number = 68 Then
'    On Error GoTo ErrorSub
'    Err.Clear
'    If Dir("C" & RutaCopia, vbDirectory) = "" Then MkDir "C" & RutaCopia
'    RutaCopiaTemp = "C" & RutaCopia
'Else
'    On Error GoTo ErrorSub
'    If Drv.DriveType = Remote And Drv.ShareName = "\\Client\C$" Then
'        If Dir("V" & RutaCopia, vbDirectory) = "" Then MkDir "V" & RutaCopia
'        RutaCopiaTemp = "V" & RutaCopia
'    Else
'        If Dir("C" & RutaCopia, vbDirectory) = "" Then MkDir "C" & RutaCopia
'        RutaCopiaTemp = "C" & RutaCopia
'    End If
'End If
'Set Drv = Nothing
'Set Fso = Nothing
'DarRutaCorrecta = RutaCopiaTemp & "\"
'
'ExitSub:
'    Exit Function
'ErrorSub:
'    With Err
'        If MsgBox(.Number & " - " & .Description & Chr(13) & _
'                  "Desea parar el programa?", _
'                  vbApplicationModal + vbYesNo + vbDefaultButton2, _
'                  "ERROR EN " & .Source) = vbYes Then
'            Stop
'            Resume Next
'        End If
'    End With
'    GoTo ExitSub
'End Function
    
Public Function GetDefaultPrinter() As String
    
    Dim strDefault    As String
    Dim lngbuf        As Long

    strDefault = String(255, Chr(0))
    lngbuf = GetProfileString("Windows", "Device", "", strDefault, Len(strDefault))
    GetDefaultPrinter = ""
    If lngbuf > 0 Then GetDefaultPrinter = Mid(strDefault, 1, InStr(1, strDefault, ",") - 1)

End Function

Public Function Redondeo(ByVal aNumberToRound As Double, _
  Optional ByVal aDecimalPlaces As Long = 0) As Double

On Error GoTo ErrHandler

Dim nFactor As Double
Dim nTemp As Double

    nFactor = 10 ^ aDecimalPlaces
    nTemp = (aNumberToRound * nFactor) + 0.5
    Redondeo = Int(CDec(nTemp)) / nFactor

ExitPoint:
Exit Function
ErrHandler:
    Select Case Err.Number
        Case Else
            Redondeo = 0
            Resume ExitPoint
    End Select
End Function

Public Function sCadena_a_HTML(ByVal sCadena As String) As String
sCambiarLetraHTML sCadena, "·", "&aacute;"
sCambiarLetraHTML sCadena, "¡", "&Aacute;"
sCambiarLetraHTML sCadena, "È", "&eacute;"
sCambiarLetraHTML sCadena, "…", "&Eacute;"
sCambiarLetraHTML sCadena, "Ì", "&iacute;"
sCambiarLetraHTML sCadena, "Õ", "&Iacute;"
sCambiarLetraHTML sCadena, "Û", "&oacute;"
sCambiarLetraHTML sCadena, "”", "&Oacute;"
sCambiarLetraHTML sCadena, "˙", "&uacute;"
sCambiarLetraHTML sCadena, "⁄", "&Uacute;"
sCambiarLetraHTML sCadena, "Ò", "&ntilde;"
sCambiarLetraHTML sCadena, "—", "&Ntilde;"

sCambiarLetraHTML sCadena, "ø", "&#191;"
sCambiarLetraHTML sCadena, "°", "&#161;"
sCambiarLetraHTML sCadena, "™", "&ordf;"
sCambiarLetraHTML sCadena, "∫", "&ordm;"
sCambiarLetraHTML sCadena, "¸", "&uuml;"
sCambiarLetraHTML sCadena, "‹", "&Uuml;"


sCadena_a_HTML = sCadena
End Function

Private Function sCambiarLetraHTML(ByRef sCadena As String, ByVal sLetraBuscar As String, ByVal sLetraPoner As String) As String
If InStr(1, sCadena, sLetraBuscar) > 0 Then sCadena = Replace$(sCadena, sLetraBuscar, sLetraPoner)
End Function
