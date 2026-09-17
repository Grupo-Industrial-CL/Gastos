VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGestionOficinasPreValidacionV2 
   BorderStyle     =   0  'None
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   18945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOficinas 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   480
      TabIndex        =   12
      Top             =   960
      Width           =   4935
      Begin VB.TextBox txtNombreOficina 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         MaxLength       =   45
         TabIndex        =   13
         Top             =   3360
         Width           =   4215
      End
      Begin MSFlexGridLib.MSFlexGrid mfgParrillaOficinas 
         Height          =   2535
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4471
         _Version        =   393216
      End
      Begin VB.Label lblid_Oficina 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   2760
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imgCrearOficina 
         Height          =   480
         Left            =   4440
         Picture         =   "frmGestionOficinasPreValidacionV2.frx":0000
         ToolTipText     =   "Añadir oficina de prevalidación"
         Top             =   3285
         Width           =   480
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
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
         Left            =   0
         TabIndex        =   16
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Oficinas prevalidación"
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
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   2145
      End
   End
   Begin VB.Frame fraDatosOficina 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3855
      Left            =   6600
      TabIndex        =   1
      Top             =   960
      Width           =   11775
      Begin VB.ComboBox cmbPreValidadores 
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
         Left            =   6360
         TabIndex        =   3
         Top             =   3360
         Width           =   4215
      End
      Begin VB.ComboBox cmbEstructurasOrganizativas 
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
         Left            =   0
         TabIndex        =   2
         Top             =   3360
         Width           =   4215
      End
      Begin MSFlexGridLib.MSFlexGrid mfgParrillaEstructurasOrganizativas 
         Height          =   2535
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4471
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid mfgParrillaPrevalidadores 
         Height          =   2535
         Left            =   6360
         TabIndex        =   5
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4471
         _Version        =   393216
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "*Doble click para eliminar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   7
         Left            =   9240
         TabIndex        =   11
         Top             =   2760
         Width           =   1965
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "*Doble click para eliminar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   6
         Left            =   2880
         TabIndex        =   10
         Top             =   2760
         Width           =   1965
      End
      Begin VB.Image imgAgregarPreValidador 
         Height          =   480
         Left            =   10800
         Picture         =   "frmGestionOficinasPreValidacionV2.frx":08CA
         ToolTipText     =   "Añadir prevalidador"
         Top             =   3285
         Width           =   480
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Prevalidadores"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   9
         Top             =   3120
         Width           =   1425
      End
      Begin VB.Image imgAgregarEstructuraOrganizativa 
         Height          =   480
         Left            =   4440
         Picture         =   "frmGestionOficinasPreValidacionV2.frx":1194
         ToolTipText     =   "Añadir estructura organizativa a prevalidar sus gastos"
         Top             =   3285
         Width           =   480
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Estructuras organizativas"
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
         Left            =   0
         TabIndex        =   8
         Top             =   3120
         Width           =   2520
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Prevalidadores asignados"
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
         Left            =   6360
         TabIndex        =   7
         Top             =   0
         Width           =   2475
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Estructuras que se prevalidarán"
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
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   3150
      End
   End
   Begin VB.Image imgGestionar 
      Height          =   480
      Left            =   5520
      Picture         =   "frmGestionOficinasPreValidacionV2.frx":1A5E
      ToolTipText     =   "Gestionar oficina"
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image imgNuevaOficina 
      Height          =   480
      Left            =   5520
      Picture         =   "frmGestionOficinasPreValidacionV2.frx":2328
      ToolTipText     =   "Cambiar oficina"
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmGestionOficinasPreValidacionV2.frx":2944
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   14400
      Picture         =   "frmGestionOficinasPreValidacionV2.frx":360E
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Gestión Oficinas Prevalidación"
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
      TabIndex        =   0
      Top             =   75
      Width           =   5415
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   18240
      Picture         =   "frmGestionOficinasPreValidacionV2.frx":3918
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
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
      Width           =   18945
   End
End
Attribute VB_Name = "frmGestionOficinasPreValidacionV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moParrillaOficinas As New CFlexGrid
Dim moParrillaEstructurasOrganizativas As New CFlexGrid
Dim moParrillaPreValidadores As New CFlexGrid

Dim mafoRegEstructurasOrganizativas As ADODB.Recordset
Dim mafoRegPreValidadores As ADODB.Recordset
Dim mbEstructurasOrganizativas As Boolean
Dim mbPreValidadores As Boolean

Private Sub cmbPreValidadores_Change()

    If mbPreValidadores = False Then Exit Sub

    Dim sAux As String
    Dim iCont As Integer
    
    For iCont = 0 To cmbPreValidadores.ListCount - 1
        cmbPreValidadores.RemoveItem 0
    Next iCont
    
    If cmbPreValidadores.Text = "" Then
    
        While Not mafoRegPreValidadores.EOF
        
            sAux = Trim(mafoRegPreValidadores("Nombre"))
            If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegPreValidadores("Codigo") & ") " & sAux
            cmbPreValidadores.AddItem sAux
            cmbPreValidadores.ItemData(cmbPreValidadores.NewIndex) = mafoRegPreValidadores("Codigo")
            mafoRegPreValidadores.MoveNext
        
        Wend
    
    Else
    
        While Not mafoRegPreValidadores.EOF
        
            If InStr(1, UCase(Trim(mafoRegPreValidadores("Nombre"))), UCase(cmbPreValidadores.Text)) Then
                sAux = Trim(mafoRegPreValidadores("Nombre"))
                If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegPreValidadores("Codigo") & ") " & sAux
                cmbPreValidadores.AddItem sAux
                cmbPreValidadores.ItemData(cmbPreValidadores.NewIndex) = mafoRegPreValidadores("Codigo")
            End If
            
            mafoRegPreValidadores.MoveNext
        
        Wend
    
    End If
    mafoRegPreValidadores.MoveFirst

End Sub

Private Sub cmbEstructurasOrganizativas_Change()

    If mbEstructurasOrganizativas = False Then Exit Sub

    Dim sAux As String
    Dim iCont As Integer
    
    For iCont = 0 To cmbEstructurasOrganizativas.ListCount - 1
        cmbEstructurasOrganizativas.RemoveItem 0
    Next iCont
    
    If cmbEstructurasOrganizativas.Text = "" Then
    
        While Not mafoRegEstructurasOrganizativas.EOF
        
            sAux = Trim(mafoRegEstructurasOrganizativas("Departamento")) & " - " & Trim(mafoRegEstructurasOrganizativas("Division_Personal")) & " - " & Trim(mafoRegEstructurasOrganizativas("Clave_Organizacion"))
            cmbEstructurasOrganizativas.AddItem sAux
            mafoRegEstructurasOrganizativas.MoveNext
        
        Wend
    
    Else
    
        While Not mafoRegEstructurasOrganizativas.EOF
        
            If InStr(1, UCase(Trim(mafoRegEstructurasOrganizativas("Departamento")) & " - " & Trim(mafoRegEstructurasOrganizativas("Division_Personal")) & " - " & Trim(mafoRegEstructurasOrganizativas("Clave_Organizacion"))), UCase(cmbEstructurasOrganizativas.Text)) Then
                sAux = Trim(mafoRegEstructurasOrganizativas("Departamento")) & " - " & Trim(mafoRegEstructurasOrganizativas("Division_Personal")) & " - " & Trim(mafoRegEstructurasOrganizativas("Clave_Organizacion"))
                cmbEstructurasOrganizativas.AddItem sAux
            End If
            
            mafoRegEstructurasOrganizativas.MoveNext
        
        Wend
    
    End If
    mafoRegEstructurasOrganizativas.MoveFirst

End Sub

Private Sub Form_Load()
    
    shRecuadroForm.Width = Me.Width
    shRecuadroForm.Height = Me.Height
    shRecuadroForm.Left = 0
    shRecuadroForm.Top = 0
    
    mbEstructurasOrganizativas = False
    mbPreValidadores = False
    
    moParrillaOficinas.Create mfgParrillaOficinas, 2, 2
    moParrillaOficinas.Formatear 0, 0, "id"
    moParrillaOficinas.Formatear 1, 4500, "Oficinas prevalidación", 4
    moParrillaOficinas.SituarEn 0, 1
    
    moParrillaEstructurasOrganizativas.Create mfgParrillaEstructurasOrganizativas, 4, 2
    moParrillaEstructurasOrganizativas.Formatear 0, 0, "Departamento"
    moParrillaEstructurasOrganizativas.Formatear 1, 0, "Division_Personal"
    moParrillaEstructurasOrganizativas.Formatear 2, 0, "Clave_Organizacion"
    moParrillaEstructurasOrganizativas.Formatear 3, 4500, "Estructuras a prevalidar", 4
    moParrillaEstructurasOrganizativas.SituarEn 0, 1
    
    moParrillaPreValidadores.Create mfgParrillaPrevalidadores, 2, 2
    moParrillaPreValidadores.Formatear 0, 0, "Usuario"
    moParrillaPreValidadores.Formatear 1, 4500, "Prevalidadores", 4
    moParrillaPreValidadores.SituarEn 0, 1
            
    RellenarOficinas
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub fraDatosOficina_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub fraOficinas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub imgAgregarPreValidador_Click()
    
    Dim id_Oficina As Integer
    
    Me.MousePointer = DEFAULT

    If cmbPreValidadores.Text = "" Or cmbPreValidadores.ListIndex = -1 Then
        cmbPreValidadores.SetFocus
        Exit Sub
    End If

    id_Oficina = moParrillaOficinas.TextoLinCol(moParrillaOficinas.Grid.Row, 0)
    If Not bPreValidacion_AsignarPreValidador(CodPaisTrabajo, cmbPreValidadores.ItemData(cmbPreValidadores.ListIndex), id_Oficina) Then
        MsgBox "Se ha producido un error y no se ha podido asignar el prevalidador a la oficina de prevalidación.", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    RellenarDatosOficina

End Sub

Private Sub imgAgregarPreValidador_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgAgregarEstructuraOrganizativa_Click()
    
    Dim id_Oficina As Integer
    
    Me.MousePointer = DEFAULT

    If cmbEstructurasOrganizativas.Text = "" Or cmbEstructurasOrganizativas.ListIndex = -1 Then
        cmbEstructurasOrganizativas.SetFocus
        Exit Sub
    End If

    id_Oficina = lblid_Oficina
    If Not bPreValidacion_AsignarEstructuraOrganizativa(CodPaisTrabajo, cmbEstructurasOrganizativas.Text, id_Oficina) Then
        MsgBox "Se ha producido un error y no se ha podido asignar la estructura organizativa a la oficina de prevalidación.", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    RellenarDatosOficina

End Sub

Private Sub imgAgregarEstructuraOrganizativa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgCrearOficina_Click()
    
    Me.MousePointer = DEFAULT
    
    txtNombreOficina = Trim(LimpiarCadena(txtNombreOficina))
    If txtNombreOficina = "" Then Exit Sub
        
    PreValidacion_CrearOficina CodPaisTrabajo, txtNombreOficina.Text
    RellenarOficinas
    
End Sub

Private Sub imgCrearOficina_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgGestionar_Click()

    Me.MousePointer = DEFAULT

    txtNombreOficina.Text = ""
    fraOficinas.Enabled = False
    fraDatosOficina.Enabled = True
    imgGestionar.Visible = False
    imgNuevaOficina.Visible = True

End Sub

Private Sub imgGestionar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgNuevaOficina_Click()

    Me.MousePointer = DEFAULT

    LimpiarDatos
    moParrillaOficinas.SituarEn 0, 1
    fraOficinas.Enabled = True
    fraDatosOficina.Enabled = False
    imgGestionar.Visible = True
    imgNuevaOficina.Visible = False
    
End Sub

Private Sub imgNuevaOficina_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgSalir_Click()
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub mfgParrillaOficinas_Click()
    
    LimpiarDatos
    If moParrillaOficinas.bLineaActualVacia Then
        moParrillaOficinas.SituarEn 0, 1
        Exit Sub
    End If
    moParrillaOficinas.ResaltarLinea
    
    RellenarDatosOficina
    
End Sub

Private Sub mfgParrillaPrevalidadores_Click()
    If moParrillaPreValidadores.bLineaActualVacia Then
        moParrillaPreValidadores.SituarEn 0, 1
        Exit Sub
    End If
    moParrillaPreValidadores.ResaltarLinea
End Sub

Private Sub mfgParrillaPrevalidadores_DblClick()

    Dim id_Oficina  As Integer
    
    If moParrillaPreValidadores.bLineaActualVacia Then
        moParrillaPreValidadores.SituarEn 0, 1
        Exit Sub
    End If
    moParrillaPreValidadores.ResaltarLinea
    
    If MsgBox("Va a eliminar el registro seleccionado, ¿está seguro?", vbAbortRetryIgnore + vbQuestion + vbYesNo, "Confirmación") = vbNo Then Exit Sub
    
    id_Oficina = moParrillaOficinas.TextoLinCol(moParrillaOficinas.Grid.Row, 0)
    If Not bPreValidacion_EliminarPrevalidador(CodPaisTrabajo, moParrillaPreValidadores.TextoLinCol(moParrillaPreValidadores.Grid.Row, 0)) Then
        MsgBox "Se ha producido un error y no se ha eliminado el registro seleccionado.", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    RellenarDatosOficina

End Sub

Private Sub mfgParrillaEstructurasOrganizativas_Click()
    If moParrillaEstructurasOrganizativas.bLineaActualVacia Then
        moParrillaPreValidadores.SituarEn 0, 1
        Exit Sub
    End If
    moParrillaEstructurasOrganizativas.ResaltarLinea
End Sub

Private Sub mfgParrillaEstructurasOrganizativas_DblClick()

    Dim id_Oficina  As Integer
    
    If moParrillaEstructurasOrganizativas.bLineaActualVacia Then
        moParrillaEstructurasOrganizativas.SituarEn 0, 1
        Exit Sub
    End If
    moParrillaEstructurasOrganizativas.ResaltarLinea
    
    If MsgBox("Va a eliminar el registro seleccionado, ¿está seguro?", vbApplicationModal + vbQuestion + vbYesNo, "Confirmación") = vbNo Then Exit Sub
    
    id_Oficina = lblid_Oficina
    If Not bPreValidacion_EliminarAsignacionEstructuraOrganizativaOficina(CodPaisTrabajo, id_Oficina, _
                moParrillaEstructurasOrganizativas.TextoLinCol(moParrillaEstructurasOrganizativas.Grid.Row, 0), _
                moParrillaEstructurasOrganizativas.TextoLinCol(moParrillaEstructurasOrganizativas.Grid.Row, 1), _
                moParrillaEstructurasOrganizativas.TextoLinCol(moParrillaEstructurasOrganizativas.Grid.Row, 2)) Then
        MsgBox "Se ha producido un error y no se ha eliminado el registro seleccionado.", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    RellenarDatosOficina

End Sub

Private Sub txtNombreOficina_GotFocus()
    TextGotFocus txtNombreOficina
End Sub

Private Sub txtNombreOficina_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        imgCrearOficina_Click
        Exit Sub
    End If
End Sub

Private Sub txtNombreOficina_LostFocus()
    TextLostFocus txtNombreOficina
End Sub

Private Sub RellenarOficinas()

    Dim afoRegistro As ADODB.Recordset

    moParrillaOficinas.Vaciar
    moParrillaOficinas.SituarEn 0, 1
    moParrillaEstructurasOrganizativas.Vaciar
    moParrillaEstructurasOrganizativas.SituarEn 0, 1
    moParrillaPreValidadores.Vaciar
    moParrillaPreValidadores.SituarEn 0, 1
    
    txtNombreOficina = ""
    VaciarBox cmbEstructurasOrganizativas
    cmbEstructurasOrganizativas.ListIndex = -1
    cmbEstructurasOrganizativas = ""
    VaciarBox cmbPreValidadores
    cmbPreValidadores.ListIndex = -1
    cmbPreValidadores = ""
    
    If Not bPreValidacion_DameOficinas(CodPaisTrabajo, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Sub
    End If
    While Not afoRegistro.EOF
        moParrillaOficinas.TextoFila moParrillaOficinas.LineaLibre, afoRegistro("id"), Trim(afoRegistro("Nombre"))
        afoRegistro.MoveNext
    Wend
    Set afoRegistro = Nothing
    moParrillaOficinas.SituarEn 0, 1

End Sub

Private Sub RellenarDatosOficina()

    Dim afoRegistro As ADODB.Recordset

    Dim id_Oficina As Integer
    Dim sEstructura As String
    
    moParrillaEstructurasOrganizativas.Vaciar
    moParrillaEstructurasOrganizativas.SituarEn 0, 1
    moParrillaPreValidadores.Vaciar
    moParrillaPreValidadores.SituarEn 0, 1
    
    id_Oficina = moParrillaOficinas.TextoLinCol(moParrillaOficinas.Grid.Row, 0)
    lblid_Oficina = id_Oficina
    
    'Rellenamos la parrilla con las estructuras organizativas asignadas a la oficina
    If bPreValidacion_DameEstructurasOrganizativasAsignadosOficina(CodPaisTrabajo, id_Oficina, afoRegistro) Then
        
        While Not afoRegistro.EOF
            moParrillaEstructurasOrganizativas.TextoFila moParrillaEstructurasOrganizativas.LineaLibre, afoRegistro("Departamento"), _
                afoRegistro("Division_Personal"), afoRegistro("Clave_Organizacion"), afoRegistro("Departamento") & " - " & afoRegistro("Division_Personal") & " - " & afoRegistro("Clave_Organizacion")
            afoRegistro.MoveNext
        Wend
    End If
    Set afoRegistro = Nothing
    moParrillaEstructurasOrganizativas.SituarEn 0, 1
    
    'Cargamos el combo con las estructuras organizativas que no estén asignadas a una oficina de prevalidación
    CargarEstructurasOrganizativas
    
    'Rellenamos los prevalidadores asignados a la oficina
    If bPreValidacion_DamePreValidadoresAsignadosOficina(CodPaisTrabajo, id_Oficina, afoRegistro) Then
        While Not afoRegistro.EOF
            moParrillaPreValidadores.TextoFila moParrillaPreValidadores.LineaLibre, afoRegistro("Usuario"), Trim(afoRegistro("Nombre"))
            afoRegistro.MoveNext
        Wend
    End If
    Set afoRegistro = Nothing
    moParrillaPreValidadores.SituarEn 0, 1
    
    'Cargamos el combo con los prevalidadores que no estén asignadas a una oficina de prevalidación
    CargarPreValidadores

End Sub

Private Sub CargarEstructurasOrganizativas()
'Buscamos las estructuras organizativas que no estén asignadas a una oficina de prevalidación

    Dim sAux As String

    Set mafoRegEstructurasOrganizativas = Nothing
    mbEstructurasOrganizativas = False
    cmbEstructurasOrganizativas.ListIndex = -1
    cmbEstructurasOrganizativas = ""
    VaciarBox cmbEstructurasOrganizativas
    
    If Not bPreValidacion_DameEstructurasOrganizativas(CodPaisTrabajo, mafoRegEstructurasOrganizativas) Then
        Set mafoRegEstructurasOrganizativas = Nothing
        Exit Sub
    End If
    
    mbEstructurasOrganizativas = True
    While Not mafoRegEstructurasOrganizativas.EOF
    
        sAux = Trim(mafoRegEstructurasOrganizativas("Departamento")) & " - " & Trim(mafoRegEstructurasOrganizativas("Division_Personal")) & " - " & Trim(mafoRegEstructurasOrganizativas("Clave_Organizacion"))
        cmbEstructurasOrganizativas.AddItem sAux
        
        mafoRegEstructurasOrganizativas.MoveNext
        
    Wend
    mafoRegEstructurasOrganizativas.MoveFirst

End Sub

Private Sub CargarPreValidadores()

    Dim sAux As String

    Set mafoRegPreValidadores = Nothing
    cmbPreValidadores.ListIndex = -1
    cmbPreValidadores = ""
    mbPreValidadores = False
    
    If Not bPreValidacion_DamePreValidadoresV2(CodPaisTrabajo, mafoRegPreValidadores) Then
        Set mafoRegPreValidadores = Nothing
        Exit Sub
    End If
    
    mbPreValidadores = True
    While Not mafoRegPreValidadores.EOF
    
        sAux = Trim(mafoRegPreValidadores("Nombre"))
        If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegPreValidadores("Codigo") & ") " & sAux
        cmbPreValidadores.AddItem sAux
        cmbPreValidadores.ItemData(cmbPreValidadores.NewIndex) = mafoRegPreValidadores("Codigo")
        
        mafoRegPreValidadores.MoveNext
        
    Wend
    mafoRegPreValidadores.MoveFirst

End Sub

Private Sub LimpiarDatos()
    Set mafoRegEstructurasOrganizativas = Nothing
    Set mafoRegPreValidadores = Nothing
    mbEstructurasOrganizativas = False
    mbPreValidadores = False
    VaciarBox cmbEstructurasOrganizativas
    cmbEstructurasOrganizativas.ListIndex = -1
    cmbEstructurasOrganizativas = ""
    VaciarBox cmbPreValidadores
    cmbPreValidadores.ListIndex = -1
    cmbPreValidadores = ""
    moParrillaEstructurasOrganizativas.Vaciar
    moParrillaEstructurasOrganizativas.SituarEn 0, 1
    moParrillaPreValidadores.Vaciar
    moParrillaPreValidadores.SituarEn 0, 1
    lblid_Oficina.Caption = ""
End Sub


