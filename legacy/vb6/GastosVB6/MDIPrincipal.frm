VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Gastos"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10755
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdCuadroDialogo 
      Left            =   240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu menuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu menuCambiarPais 
         Caption         =   "Cambiar Pais/Empresa"
         Shortcut        =   ^A
      End
      Begin VB.Menu menuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu menuGastos 
      Caption         =   "&Gastos Red"
      Visible         =   0   'False
      Begin VB.Menu menuIntroGastos 
         Caption         =   "&Introducir Gastos"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnAnexarDescuentos 
         Caption         =   "Anexar Descuentos"
      End
      Begin VB.Menu menuModifDirec 
         Caption         =   "&Modificar Director"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnCalcMetInt 
         Caption         =   "Calcular Metálico Internet"
      End
      Begin VB.Menu mnValidarGastos 
         Caption         =   "Validar Gastos"
      End
      Begin VB.Menu mnCargarGastosFichero 
         Caption         =   "Cargar gastos desde fichero"
      End
      Begin VB.Menu Separador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnConsultaEstado 
         Caption         =   "Consultar estado de los gastos"
      End
      Begin VB.Menu mnConsultaUsuariosZonaPrivada 
         Caption         =   "Usuarios Zona Privada"
      End
      Begin VB.Menu Separador12 
         Caption         =   "-"
      End
      Begin VB.Menu mnCargaFicSolred 
         Caption         =   "Carga fichero Solred"
      End
   End
   Begin VB.Menu mnPagos 
      Caption         =   "Pagos Red"
      Visible         =   0   'False
   End
   Begin VB.Menu menuInformes 
      Caption         =   "&Informes Red"
      Visible         =   0   'False
      Begin VB.Menu menuPorPantalla 
         Caption         =   "&Pantalla"
         Shortcut        =   ^I
      End
      Begin VB.Menu menuPorFichero 
         Caption         =   "&Fichero"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnGastosInternos 
      Caption         =   "Gastos Internos"
      Begin VB.Menu mnIntroducirGastosInternos 
         Caption         =   "Introducir gastos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnIntroducirGastos2 
         Caption         =   "Introducir gastos"
      End
      Begin VB.Menu mnInfPersonal 
         Caption         =   "Informes gastos"
      End
      Begin VB.Menu mnSep01_GastosInt 
         Caption         =   "-"
      End
      Begin VB.Menu mnDepartamentos 
         Caption         =   "Departamentos"
      End
      Begin VB.Menu mnJefesDepartamentos 
         Caption         =   "Jefes departamentos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnUsuariosDepartamentos 
         Caption         =   "Usuarios departamentos"
      End
      Begin VB.Menu mnSep04_GastosInt 
         Caption         =   "-"
      End
      Begin VB.Menu mnGestionOfiPreValidacion 
         Caption         =   "Gestión Oficinas PreValidación"
      End
      Begin VB.Menu mnGestionOfiPreValidacion2 
         Caption         =   "Gestión Oficinas PreValidación (Estruc. Org.)"
      End
      Begin VB.Menu mnInformacionCentrosCoste 
         Caption         =   "Información centros de coste"
      End
      Begin VB.Menu mnSep02_GastosInt 
         Caption         =   "-"
      End
      Begin VB.Menu mnCargaFicSolred2 
         Caption         =   "Carga fichero Solred"
      End
      Begin VB.Menu mnCrearTarjeta 
         Caption         =   "Crear tarjeta"
      End
      Begin VB.Menu mnSep03_GastosInt 
         Caption         =   "-"
      End
      Begin VB.Menu mnCargarFicheroLayTours 
         Caption         =   "Cargar fichero LayTours"
      End
   End
   Begin VB.Menu mnInformesDireccion 
      Caption         =   "Informe"
      Visible         =   0   'False
   End
   Begin VB.Menu mnTransferenciasInternas 
      Caption         =   "Transferencias Internas"
      Begin VB.Menu mnInformeTIFechas 
         Caption         =   "Informe entre fechas"
      End
   End
   Begin VB.Menu mnCB 
      Caption         =   "Caja Badajoz"
      Visible         =   0   'False
      Begin VB.Menu mnContabilizarPedidosCB 
         Caption         =   "Contabilizar pedidos"
      End
   End
   Begin VB.Menu mnPerseida 
      Caption         =   "Perseida"
      Visible         =   0   'False
      Begin VB.Menu mnContPedPerseida 
         Caption         =   "Contabilizar pedidos"
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Activate()
    
    Me.Caption = "Gestión de Gastos de Viaje - Pais/Empresa: " & sNombrePais(CodPaisTrabajo) & " - Usuario: " & goUsuario.Nombre & " - PC: " & Environ("COMPUTERNAME") & " - Versión: " & CStr(VB.App.Major) & "." & CStr(VB.App.Minor) & "." & CStr(VB.App.Revision)
    If gAnchoPantallaPrincipal = -1 Then
        gAnchoPantallaPrincipal = Me.Width
        gAnchoPantallaVirtual = VirtualScreenWidth * 15
    End If
    
End Sub

Private Sub MDIForm_Load()
    
    On Error GoTo Error_Rutina
    
    mnGestionOfiPreValidacion.Enabled = False
    Select Case goUsuario.IdUsuario
    Case "lanchi", "mmarsanchez", "avegas", "agonzalez", "ivazquez", "luciagil", "inmamoya", "ialcal", "lvazquezmo"
        mnGestionOfiPreValidacion.Enabled = True
    End Select
    
    Select Case goUsuario.PaisTrabajo
    Case giPAIS_GASEXTREMADURA, giPAIS_ONDUPACK, giPAIS_ONDUPET, giPAIS_ALTERENERSUN, giPAIS_PLASTIVERD, giPAIS_IQOXE, giPAIS_GES, giPAIS_LILOTECHNOLOGY, giPAIS_MATIASGOMA, _
         giPAIS_AGSIDERURGICABALBOA, giPAIS_GRUPOGALLARDOBALBOA, giPAIS_ALFONSOGALLARDO, giPAIS_CORRUGADOSLASAO, giPAIS_MARCELIANOMARTIN, giPAIS_FERROMALLAS, _
         giPAIS_CORRUGADOSGETAFE, giPAIS_GALVACOLOR, giPAIS_SOLARSTEEL, giPAIS_CARTONAJESEXTREMADURA
        menuGastos.Visible = False
        menuInformes.Visible = False
        mnPagos.Visible = False
        mnInformesDireccion.Visible = False
        mnCB.Visible = False
        mnPerseida.Visible = False
        mnTransferenciasInternas.Visible = False
        
        mnGastosInternos.Visible = True
        
        Exit Sub
    End Select
    
    Select Case goUsuario.IdUsuario
    Case "ldelgado"
        menuGastos.Visible = False
        menuInformes.Visible = False
        mnPagos.Visible = False
        mnInformesDireccion.Visible = False
        mnCB.Visible = False
        mnPerseida.Visible = False
        mnGastosInternos.Visible = True
        mnTransferenciasInternas.Visible = False
    
    Case "rleal", "jcarballo", "jfsantos"
        menuGastos.Visible = False
        mnPagos.Visible = False
        mnCB.Visible = False
        mnPerseida.Visible = False
        If goUsuario.IdUsuario <> "jfsantos" Then
            menuInformes.Visible = False
        Else
            mnInformesDireccion.Visible = False
        End If
        mnInfPersonal.Visible = False
        mnTransferenciasInternas.Visible = False
    Case "msantana"
        mnGastosInternos.Visible = True
    Case Else
        If goUsuario.IdUsuario = "lanchi" Or goUsuario.IdUsuario = "jmbenitez" Or goUsuario.IdUsuario = "fchamorro" _
        Or goUsuario.IdUsuario = "ajcarrasco" Or goUsuario.IdUsuario = "abenavides" _
        Or goUsuario.IdUsuario = "agonzalez" Or goUsuario.IdUsuario = "ivazquez" Or goUsuario.IdUsuario = "luciagil" _
        Or goUsuario.IdUsuario = "inmamoya" Or goUsuario.IdUsuario = "lvazquezmo" Or goUsuario.IdUsuario = "ialcal" Or goUsuario.IdUsuario = "villafaina" _
        Or goUsuario.IdUsuario = "fgonzalez" Or goUsuario.IdUsuario = "janaharro" _
        Or goUsuario.IdUsuario = "aortiz" Or goUsuario.IdUsuario = "vlunar" Or goUsuario.IdUsuario = "macaceres" Or goUsuario.IdUsuario = "fsanchezc" _
        Or goUsuario.IdUsuario = "nvazquez" Or goUsuario.IdUsuario = "jromero" Or goUsuario.IdUsuario = "dsanchez" _
        Or goUsuario.IdUsuario = "rcarrascal" _
        Or goUsuario.IdUsuario = "623" Or goUsuario.IdUsuario = "702" Then Exit Sub
        mnInformesDireccion.Visible = False
        mnCB.Visible = False
        mnPerseida.Visible = False
        mnGastosInternos.Visible = False
        mnTransferenciasInternas.Visible = False
        Select Case goUsuario.IdUsuario
        Case "ssaouri", "jpmurillo", "roriana", "bhantout"
            mnGastosInternos.Visible = True
        End Select
    End Select
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub menuCambiarPais_Click()
'frmCambioPais.Show vbModal
    frmSeleccionarPais.Show MODAL
End Sub

Private Sub menuIntroGastos_Click()
frmIntroGastos.Show vbModal
End Sub

Private Sub menuModifDirec_Click()
frmModifDirec.Show
End Sub

Private Sub menuPorFichero_Click()
frmFicheros.Show vbModal
End Sub

Private Sub menuPorPantalla_Click()
frmInformes.Show vbModal
End Sub

Private Sub menuSalir_Click()
    EmpleadoDesBloquear_UsuGasto
    gBdGALILEO.cn.Close
    Set gBdGALILEO = Nothing
    Set goUsuario = Nothing
    End
End Sub

Private Sub mnAnexarDescuentos_Click()
    frmDescuentos.Show vbModal
End Sub

Private Sub mnCalcMetInt_Click()
    frmCalculoMetalicoInternet.Show vbModal
End Sub

Private Sub mnCargaFicSolred_Click()

    If goUsuario.IdUsuario <> "lanchi" And goUsuario.IdUsuario <> "agonzalez" And goUsuario.IdUsuario <> "ivazquez" And goUsuario.IdUsuario <> "inmamoya" And _
       goUsuario.IdUsuario <> "jmbenitez" And goUsuario.IdUsuario <> "702" And goUsuario.IdUsuario <> "ialcal" And goUsuario.IdUsuario <> "lvazquezmo" And _
       goUsuario.IdUsuario <> "623" And goUsuario.IdUsuario <> "janaharro" And goUsuario.IdUsuario <> "fsanchezc" And _
       goUsuario.IdUsuario <> "nvazquez" And goUsuario.IdUsuario <> "jromero" And goUsuario.IdUsuario <> "dsanchez" And _
       goUsuario.IdUsuario <> "aortiz" And goUsuario.IdUsuario <> "vlunar" And goUsuario.IdUsuario <> "macaceres" Then Exit Sub
    
    frmCargaFicheroSolred2.Show MODAL

End Sub

Private Sub mnCargaFicSolred2_Click()

    If goUsuario.IdUsuario <> "lanchi" And goUsuario.IdUsuario <> "agonzalez" And goUsuario.IdUsuario <> "ivazquez" And goUsuario.IdUsuario <> "inmamoya" And _
       goUsuario.IdUsuario <> "jmbenitez" And goUsuario.IdUsuario <> "702" And goUsuario.IdUsuario <> "ialcal" And goUsuario.IdUsuario <> "lvazquezmo" And _
       goUsuario.IdUsuario <> "623" And goUsuario.IdUsuario <> "janaharro" And goUsuario.IdUsuario <> "fsanchezc" And _
       goUsuario.IdUsuario <> "nvazquez" And goUsuario.IdUsuario <> "jromero" And goUsuario.IdUsuario <> "dsanchez" And _
       goUsuario.IdUsuario <> "aortiz" And goUsuario.IdUsuario <> "vlunar" And goUsuario.IdUsuario <> "macaceres" Then Exit Sub
    
    frmCargaFicheroSolred2.Show MODAL

End Sub

Private Sub mnCargarFicheroLayTours_Click()
    
    If goUsuario.IdUsuario <> "lanchi" And goUsuario.IdUsuario <> "agonzalez" And goUsuario.IdUsuario <> "inmamoya" And _
       goUsuario.IdUsuario <> "ialcal" And goUsuario.IdUsuario <> "lvazquezmo" And goUsuario.IdUsuario <> "fsanchezc" And _
       goUsuario.IdUsuario <> "nvazquez" And goUsuario.IdUsuario <> "jromero" And goUsuario.IdUsuario <> "dsanchez" And _
       goUsuario.IdUsuario <> "aortiz" And goUsuario.IdUsuario <> "vlunar" And goUsuario.IdUsuario <> "macaceres" Then Exit Sub
    
    frmCargarFicheroORBIS.Show MODAL

End Sub

Private Sub mnCargarGastosFichero_Click()

    If goUsuario.IdUsuario <> "agonzalez" And goUsuario.IdUsuario <> "ivazquez" And goUsuario.IdUsuario <> "inmamoya" And _
       goUsuario.IdUsuario <> "janaharro" And goUsuario.IdUsuario <> "fsanchezc" And _
       goUsuario.IdUsuario <> "nvazquez" And goUsuario.IdUsuario <> "jromero" And goUsuario.IdUsuario <> "dsanchez" And _
       goUsuario.IdUsuario <> "aortiz" And goUsuario.IdUsuario <> "vlunar" And goUsuario.IdUsuario <> "macaceres" And _
       goUsuario.IdUsuario <> "lanchi" And goUsuario.IdUsuario <> "623" And _
       goUsuario.IdUsuario <> "702" And goUsuario.IdUsuario <> "ialcal" And goUsuario.IdUsuario <> "lvazquezmo" Then Exit Sub
    
    frmCargarGastosFichero.Show MODAL

End Sub

Private Sub mnConsultaEstado_Click()
    frmEstadoGastos.Show vbModal
End Sub

Private Sub mnConsultaUsuariosZonaPrivada_Click()

    '08-11-21 Por indicaciones de Filo, quitamos la opción de que puedan ver la clave de los usuarios
    Exit Sub
    '08-11-21 Fin

    If goUsuario.IdUsuario <> "agonzalez" And goUsuario.IdUsuario <> "ivazquez" And goUsuario.IdUsuario <> "lanchi" And _
       goUsuario.IdUsuario <> "jmbenitez" And goUsuario.IdUsuario <> "623" And _
       goUsuario.IdUsuario <> "702" And goUsuario.IdUsuario <> "janaharro" And goUsuario.IdUsuario <> "fsanchezc" And _
       goUsuario.IdUsuario <> "nvazquez" And goUsuario.IdUsuario <> "jromero" And goUsuario.IdUsuario <> "dsanchez" Then Exit Sub
    
    frmUsuarioZonaPrivada.Show vbModal

End Sub

Private Sub mnContabilizarPedidosCB_Click()
    frmContabilizarPedidosCB.Show vbModal
End Sub

Private Sub mnContPedPerseida_Click()

    Select Case goUsuario.IdUsuario
    Case "lanchi", "rcarrascal", "atorvisco", "JOSELA"
        frmContabilizarPedidosPerseida.Show vbModal
    Case Else
        Exit Sub
    End Select

End Sub

Private Sub mnCrearTarjeta_Click()
    
    Select Case goUsuario.IdUsuario
    Case "lanchi", "agonzalez", "ivazquez", "inmamoya", "ialcal", "lvazquezmo", "aortiz", "vlunar", "macaceres", "fsanchezc", "nvazquez", "jromero", "dsanchez"
        frmCrearTarjetaSOLRED.Show MODAL
    Case Else
        Exit Sub
    End Select
    
End Sub

Private Sub mnDepartamentos_Click()
    Select Case goUsuario.IdUsuario
    Case "lanchi", "agonzalez", "ivazquez", "inmamoya", "ialcal", "fsanchezc", "nvazquez", "jromero", "dsanchez", "lvazquezmo", "aortiz", "vlunar", "macaceres"
        frmDepartamentosCL.Show MODAL
    Case Else
        Exit Sub
    End Select
End Sub

Private Sub mnGestionOfiPreValidacion_Click()
    Select Case goUsuario.IdUsuario
    Case "mmarsanchez", "avegas", "lanchi", "agonzalez", "ivazquez", "luciagil", "inmamoya", "ialcal", "lvazquezmo"
        frmGestionOficinasPreValidacion.Show MODAL
    Case Else
        Exit Sub
    End Select
End Sub

Private Sub mnGestionOfiPreValidacion2_Click()
    Select Case goUsuario.IdUsuario
    Case "mmarsanchez", "lanchi", "agonzalez", "ivazquez", "luciagil", "inmamoya", "ialcal", "lvazquezmo"
        frmGestionOficinasPreValidacionV2.Show MODAL
    Case Else
        Exit Sub
    End Select
End Sub

Private Sub mnInformacionCentrosCoste_Click()
    'frmInformacionCentrosCoste
End Sub

Private Sub mnInformesDireccion_Click()
    frmInfDireccion.Show vbModal
End Sub

Private Sub mnInformeTIFechas_Click()
    Select Case goUsuario.IdUsuario
    Case "abenavides", "ajcarrasco", "jmbenitez", "lanchi", "fchamorro"
        frmInfTIFechas.Show MODAL
    Case Else
    
    End Select
End Sub

Private Sub mnInfPersonal_Click()
    
    Select Case CodPaisTrabajo
    Case giPAIS_GES
        Select Case goUsuario.IdUsuario
        Case "lanchi", "mmarsanchez", "avegas", "agonzalez", "ivazquez", "luciagil", "inmamoya", "ialcal", "lvazquezmo"
            frmInfGastosPersonalGES.Show MODAL
        Case Else
            MsgBox "No tiene acceso a esta opción.", vbApplicationModal + vbExclamation, "Información"
            Exit Sub
        End Select
    Case Else
        Select Case goUsuario.IdUsuario
        Case "abenavides", "ajcarrasco", "jmbenitez", "lanchi", "agonzalez", "ivazquez", "luciagil", "inmamoya", "lvazquezmo", "ldelgado", "moperez", "623", "702", "janaharro", "fsanchezc", "nvazquez", "jromero", "dsanchez", "ialcal", "aortiz", "vlunar", "macaceres"
            frmInfGastosPersonal.Show MODAL
        End Select
    End Select
    
End Sub

Private Sub mnIntroducirGastos2_Click()
    frmIntroducirGastosInternos_v2.Show 'MODAL
End Sub

Private Sub mnIntroducirGastosInternos_Click()
    frmIntroducirGastosInternos.Show MODAL
End Sub

Private Sub mnJefesDepartamentos_Click()
'    Select Case goUsuario.IdUsuario
'    Case "lanchi", "agonzalez", "ivazquez"
'        frmJefesDepartamento.Show MODAL
'    Case Else
'        Exit Sub
'    End Select
End Sub

Private Sub mnPagos_Click()
    frmPagos.Show vbModal
End Sub

Private Sub mnUsuariosDepartamentos_Click()

    Select Case goUsuario.IdUsuario
    Case "lanchi", "agonzalez", "ivazquez", "inmamoya", "ialcal", "fsanchezc", "nvazquez", "jromero", "dsanchez", "lvazquezmo", "aortiz", "vlunar", "macaceres"
        
        Set gBDGASTOS = Nothing
        Set gBDGASTOS = New cBaseDatos
        gBDGASTOS.CreatePorCadenaDeConexion "Data Source='" & DSN_SQL_GASTOS & "'"
        If Not gBDGASTOS.CreadoConExito Then
            MsgBox "No se puede abrir la opción porque no tiene conexión a la BD, contacte con IT.", vbApplicationModal + vbCritical, "E R R O R"
            Exit Sub
        End If
        frmUsuariosPersonal.Show MODAL
        gBDGASTOS.cn.Close
        Set gBDGASTOS = Nothing
        
    Case Else
        Exit Sub
    End Select
    
End Sub

Private Sub mnValidarGastos_Click()
    
    Select Case goUsuario.IdUsuario
    Case "lanchi", "agonzalez", "ivazquez", "623", "702", "janaharro", "fsanchezc", "nvazquez", "jromero", "dsanchez", "inmamoya", "ialcal", "lvazquezmo", "aortiz", "vlunar", "macaceres"
        frmValidarGastos.Show MODAL
    End Select
    
End Sub
