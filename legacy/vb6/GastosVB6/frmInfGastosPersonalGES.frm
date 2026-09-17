VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfGastosPersonalGES 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18720
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   18720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPersonal 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   17415
      Begin VB.OptionButton optAnticipos 
         Caption         =   "anticipos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11160
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optGastos 
         Caption         =   "gastos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "CONSULTAR"
         Height          =   1215
         Left            =   15120
         Picture         =   "frmInfGastosPersonalGES.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   200
         Width           =   2055
      End
      Begin VB.CheckBox chkBajas 
         Caption         =   "Listar usuarios baja"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbDepartamentos 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   200
         Width           =   5775
      End
      Begin VB.ComboBox cmbEmpleados 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   345
         Left            =   10680
         TabIndex        =   3
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yy"
         Format          =   43581443
         CurrentDate     =   42384
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   345
         Left            =   12720
         TabIndex        =   4
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yy"
         Format          =   43581443
         CurrentDate     =   42384
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Buscar "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   8880
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Buscar desde el "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   8880
         TabIndex        =   14
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "al"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   12360
         TabIndex        =   13
         Top             =   240
         Width           =   165
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Empleado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   765
         Width           =   930
      End
      Begin VB.Image imgAceptarEmpleado 
         Height          =   480
         Left            =   6720
         Picture         =   "frmInfGastosPersonalGES.frx":0ECA
         Top             =   1080
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfgParrilla 
      Height          =   6135
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   10821
      _Version        =   393216
   End
   Begin VB.Image imgExportar 
      Height          =   480
      Left            =   18000
      Picture         =   "frmInfGastosPersonalGES.frx":1794
      ToolTipText     =   "Copiar al Portapapeles"
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   18000
      Picture         =   "frmInfGastosPersonalGES.frx":205E
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Informe gastos/anticipos personal interno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   8
      Left            =   720
      TabIndex        =   9
      Top             =   75
      Width           =   7605
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   14400
      Picture         =   "frmInfGastosPersonalGES.frx":256A
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmInfGastosPersonalGES.frx":2874
      Top             =   60
      Width           =   480
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   150
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E2C07A&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   18720
   End
End
Attribute VB_Name = "frmInfGastosPersonalGES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mafoRegistro As ADODB.Recordset
Dim mafoRegEmpleados As ADODB.Recordset

Dim moParrilla As New CFlexGrid

Dim mbEmpleados As Boolean
Dim miCodDepartamento As Integer
Dim miCodEmpleado As Long
Dim msJerEmpleado As String

Dim cMouseW As New cMouseWheel

Private Sub cmbDepartamentos_Click()

    Dim sAux As String

    If cmbDepartamentos.Tag = "C" Then Exit Sub
    VaciarBox cmbEmpleados
    If cmbDepartamentos.Text = "" Then Exit Sub
    
    sAux = "N"
    If chkBajas.Value = vbChecked Then sAux = "S"
    Set mafoRegEmpleados = Nothing
    mbEmpleados = False
    If bDameEmpleadosDepartamento(CodPaisTrabajo, cmbDepartamentos.ItemData(cmbDepartamentos.ListIndex), mafoRegEmpleados, sAux, 0) Then
        mbEmpleados = True
        While Not mafoRegEmpleados.EOF
            sAux = Trim(mafoRegEmpleados("Nombre"))
            If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegEmpleados("Codigo") & ") " & sAux
            cmbEmpleados.AddItem sAux
            cmbEmpleados.ItemData(cmbEmpleados.NewIndex) = mafoRegEmpleados("Codigo")
            mafoRegEmpleados.MoveNext
        Wend
        mafoRegEmpleados.MoveFirst
    End If
    
End Sub

Private Sub cmbDepartamentos_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub cmbEmpleados_Change()

    Dim sAux As String

    If mbEmpleados = False Then Exit Sub

    'VaciarBox cmbEmpleados
    Dim iCont As Integer
    
    For iCont = 0 To cmbEmpleados.ListCount - 1
        cmbEmpleados.RemoveItem 0
    Next iCont
    
    If cmbEmpleados.Text = "" Then
    
        While Not mafoRegEmpleados.EOF
        
            sAux = Trim(mafoRegEmpleados("Nombre"))
            If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegEmpleados("Codigo") & ") " & sAux
            cmbEmpleados.AddItem sAux
            cmbEmpleados.ItemData(cmbEmpleados.NewIndex) = mafoRegEmpleados("Codigo")
            mafoRegEmpleados.MoveNext
        
        Wend
    
    Else
    
        While Not mafoRegEmpleados.EOF
        
            If InStr(1, UCase(Trim(mafoRegEmpleados("Nombre"))), UCase(cmbEmpleados.Text)) Then
                sAux = Trim(mafoRegEmpleados("Nombre"))
                If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegEmpleados("Codigo") & ") " & sAux
                cmbEmpleados.AddItem sAux
                cmbEmpleados.ItemData(cmbEmpleados.NewIndex) = mafoRegEmpleados("Codigo")
            End If
            
            mafoRegEmpleados.MoveNext
        
        Wend
    
    End If
    mafoRegEmpleados.MoveFirst
    
End Sub

Private Sub cmbEmpleados_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub cmdConsultar_Click()
    imgAceptarEmpleado_Click
End Sub

Private Sub Form_Load()

    If DisplayMonitorCount = 1 Then
        Me.Left = (Screen.Width / 2) - (Me.Width / 2)
        Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Else
        'Me.Move Screen.Width + (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
        'Me.Move gAnchoPantallaPrincipal + (((Screen.Width - gAnchoPantallaPrincipal) - Me.Width) \ 2), (Screen.Height - Me.Height) \ 2
        If MDIPrincipal.Left + 120 < Screen.Width Then
        'El MDI está en la pantalla 1, lo centro en la pantalla 1
            Me.Left = (Screen.Width / 2) - (Me.Width / 2)
            Me.Top = (Screen.Height / 2) - (Me.Height / 2)
        Else
            Me.Left = gAnchoPantallaPrincipal + (((gAnchoPantallaVirtual - gAnchoPantallaPrincipal) - Me.Width) \ 2)
            Me.Top = (Screen.Height - Me.Height) \ 2
            'Me.Move gAnchoPantallaPrincipal + (((gAnchoPantallaVirtual - gAnchoPantallaPrincipal) - Me.Width) \ 2), (Screen.Height - Me.Height) \ 2
        End If
    End If

    shRecuadroForm.Width = Me.Width
    shRecuadroForm.Height = Me.Height
    shRecuadroForm.Left = 0
    shRecuadroForm.Top = 0

    dtpFechaDesde.Value = Now - 30
    dtpFechaHasta.Value = Now
    
    moParrilla.Create mfgParrilla, 18, 2
    moParrilla.Formatear 0, 0, "Contador"
    moParrilla.Formatear 1, 1300, "Nº personal", 4, 7, True
    moParrilla.Formatear 2, 3500, "Nombre del empleado o candidato", 4, , True
    moParrilla.Formatear 3, 1400, "Nº sociedad", 4, 7, True
    moParrilla.Formatear 4, 1600, "Centro Coste", 4, , True
    moParrilla.Formatear 5, 2300, "Clase gasto viaje", 4, , True
    moParrilla.Formatear 6, 1800, "Impte. mon. local", 4, 7, True
    moParrilla.Formatear 7, 1000, "Moneda", 4, 4, True
    moParrilla.Formatear 8, 1800, "Importe €", 4, 7, True
    moParrilla.Formatear 9, 1300, "Cambio", 4, 7
    moParrilla.Formatear 10, 1600, "Fe. comprobante", 4, 4, True
    moParrilla.Formatear 11, 4000, "Descripción", 4
    moParrilla.Formatear 12, 1250, "Prevalidado", 4, 4, True
    moParrilla.Formatear 13, 1750, "Ofi. Prevalidación", 4, , True
    moParrilla.Formatear 14, 1000, "Validado", 4, 4, True
    moParrilla.Formatear 15, 3500, "Validador", 4, , True
    moParrilla.Formatear 16, 1250, "Conta. SAP", 4, 4, True
    moParrilla.Formatear 17, 1000, "Pagado", 4, 4, True
    moParrilla.SituarEn 0, 1
    
    cmbDepartamentos.Tag = "C"
    If bDameDepartamentosCL(CodPaisTrabajo, mafoRegistro, 0) Then
        While Not mafoRegistro.EOF
            cmbDepartamentos.AddItem Trim(mafoRegistro("Departamento"))
            cmbDepartamentos.ItemData(cmbDepartamentos.NewIndex) = mafoRegistro("CodDepartamento")
            mafoRegistro.MoveNext
        Wend
    End If
    cmbDepartamentos.Tag = ""
    Set mafoRegistro = Nothing
    
    cMouseW.InitMouseWheel mfgParrilla.hwnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub fraPersonal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub imgAceptarEmpleado_Click()
    
    Me.MousePointer = DEFAULT

    miCodDepartamento = -1
    miCodEmpleado = -1
    msJerEmpleado = ""

    If cmbDepartamentos.Text <> "" Then miCodDepartamento = cmbDepartamentos.ItemData(cmbDepartamentos.ListIndex)
    If cmbEmpleados.Text <> "" Then
        If cmbEmpleados.ListIndex > -1 Then
            miCodEmpleado = cmbEmpleados.ItemData(cmbEmpleados.ListIndex)
            msJerEmpleado = sDameJerarquiaEmpleado(CodPaisTrabajo, miCodDepartamento, miCodEmpleado)
        Else
            cmbEmpleados.SetFocus
            Exit Sub
        End If
    End If
    
    If dtpFechaHasta < dtpFechaDesde Then
        MsgBox "La fecha hasta no puede ser menor que la fecha desde.", vbApplicationModal + vbInformation, "Información"
        dtpFechaDesde.SetFocus
        Exit Sub
    End If
    
    RellenarDatosGastos
    
End Sub

Private Sub imgAceptarEmpleado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgExportar_Click()
    Me.MousePointer = DEFAULT
    moParrilla.SituarEn 0, 1
    If moParrilla.bLineaActualVacia Then Exit Sub
    moParrilla.aClipBoard 1
    MsgBox "Información copiada al Portapapeles.", vbApplicationModal + vbInformation, "Información"
End Sub

Private Sub imgExportar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgSalir_Click()
    Me.MousePointer = DEFAULT
    Set mafoRegistro = Nothing
    Set mafoRegEmpleados = Nothing
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub RellenarDatosGastos()

    Dim sTipoBusqueda As String
    
    moParrilla.Vaciar
    moParrilla.SituarEn 0, 1
    
    If optGastos.Value = True Then sTipoBusqueda = "gastos"
    If optAnticipos.Value = True Then sTipoBusqueda = "anticipos"
    
    If Not bDameGastosInternos_Informe(CodPaisTrabajo, Format(dtpFechaDesde.Value, "ddMMyyyy"), Format(dtpFechaHasta.Value, "ddMMyyyy"), miCodDepartamento, miCodEmpleado, msJerEmpleado, mafoRegistro, sTipoBusqueda) Then
        Set mafoRegistro = Nothing
        Exit Sub
    End If
    While Not mafoRegistro.EOF
    
        moParrilla.TextoFila moParrilla.LineaLibre, mafoRegistro("Contador"), mafoRegistro("Codigo"), Trim(mafoRegistro("Nombre")), Format(mafoRegistro("Departamento"), "0000"), _
            Trim(mafoRegistro("PEP_CECO")), Trim(mafoRegistro("TipoGasto")), mafoRegistro("Importe_Ori"), Trim(mafoRegistro("Moneda")), mafoRegistro("Importe"), mafoRegistro("Cambio"), _
            Format(mafoRegistro("Fecha"), "dd/MM/yy"), Trim(mafoRegistro("Donde")), Trim(mafoRegistro("PreValidado")), Trim(mafoRegistro("OficinaPrevalidacion")), _
            Trim(mafoRegistro("Validado")), Trim(mafoRegistro("NombreValidador")), Trim(mafoRegistro("Conta_SAP")), Trim(mafoRegistro("Pagado"))
            
        mafoRegistro.MoveNext
        
    Wend
    Set mafoRegistro = Nothing
    moParrilla.SituarEn 0, 1

End Sub

Private Sub mfgParrilla_Click()

    If moParrilla.Grid.MouseRow = 0 Then
        moParrilla.OrdenarSiCorresponde
        moParrilla.SituarEn 0, 1
        Exit Sub
    End If

    If moParrilla.bLineaActualVacia Then
        moParrilla.SituarEn 0, 1
        Exit Sub
    End If
    moParrilla.ResaltarLinea
    
End Sub

Private Sub mfgParrilla_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub
