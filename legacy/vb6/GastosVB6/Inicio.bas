Attribute VB_Name = "Inicio"
Option Explicit


Sub Main()

    Dim sConnect As String
    Dim frmLogin As New frmUserLogin
    
    On Error GoTo ErrorSub
    
    If App.PrevInstance Then
        MsgBox "¡La Aplicación ya se está ejecutando!", vbCritical, "Aviso"
        End
    End If
        
    'SeparadorDecimal = ObtenerSeparadorDecimal
    'SeparadorMiles = ObtenerSeparadorMiles
    '
    'If SeparadorDecimal <> "." Or SeparadorMiles <> "," Then
    '    MsgBox "Error en la configuración regional." & Chr(13) & _
    '        "Ponga como separador decimal (" & SeparadorDecimal & ") " & _
    '        "y como separador de miles (" & SeparadorMiles & ")"
    '    Exit Sub
    'End If
    
    'frmPresentacion.Show
    'DoEvents
    'Sleep 1000
    Inicializar_Variables
    gAnchoPantallaPrincipal = -1
    
    Screen.MousePointer = vbHourglass
    gBdGALILEO.Create DSN_CRISLAY, BD_CRISLAY, "informix", "1nformix"
    Screen.MousePointer = DEFAULT
    
    frmLogin.Create giAPL_GASTOS_VIAJE, gBdGALILEO.cn
    
    frmLogin.Show vbModal
    DoEvents
            
    Set goUsuario = frmLogin.Usuario
    Set MDIPadre = MDIPrincipal
            
    If Not goUsuario Is Nothing Then
        If goUsuario.Permiso(giOP_GTO_ACCESO) Then
        
            gsRutaEquipo = ""
            InicializarRuta
            If gsRutaEquipo = "" Then
                If goUsuario.PaisTrabajo = giPAIS_ESPAÑA Then
                    gsRutaEquipo = "C:\"
                Else
                    gsRutaEquipo = "V:\"
                End If
            End If
        
            CodPaisTrabajo = goUsuario.PaisTrabajo
            Empresa_CodPaisTrabajo = 1
            MDIPrincipal.Show
            DoEvents
        Else
            MsgBox "No tiene permiso para utilizar esta aplicación", vbExclamation, "Usuario No Autorizado"
        End If
    End If
    
    Unload frmPresentacion
    Screen.MousePointer = DEFAULT
    
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



