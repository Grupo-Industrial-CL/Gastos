VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGestionOficinasPreValidacion 
   BorderStyle     =   0  'None
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18930
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
   ScaleHeight     =   5250
   ScaleWidth      =   18930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   13320
      TabIndex        =   5
      Top             =   4440
      Width           =   4215
   End
   Begin VB.ComboBox cmbResponsables 
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
      Left            =   6960
      TabIndex        =   3
      Top             =   4440
      Width           =   4215
   End
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
      Left            =   600
      MaxLength       =   45
      TabIndex        =   1
      Top             =   4440
      Width           =   4215
   End
   Begin MSFlexGridLib.MSFlexGrid mfgParrillaOficinas 
      Height          =   2535
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid mfgParrillaResponsables 
      Height          =   2535
      Left            =   6960
      TabIndex        =   2
      Top             =   1320
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid mfgParrillaPrevalidadores 
      Height          =   2535
      Left            =   13320
      TabIndex        =   4
      Top             =   1320
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
      Left            =   16200
      TabIndex        =   14
      Top             =   3840
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
      Left            =   9840
      TabIndex        =   13
      Top             =   3840
      Width           =   1965
   End
   Begin VB.Image imgAgregarPreValidador 
      Height          =   480
      Left            =   17760
      Picture         =   "frmGestionOficinasPreValidacion.frx":0000
      ToolTipText     =   "Añadir prevalidador"
      Top             =   4365
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
      Left            =   13320
      TabIndex        =   12
      Top             =   4200
      Width           =   1425
   End
   Begin VB.Image imgAgregarResposable 
      Height          =   480
      Left            =   11400
      Picture         =   "frmGestionOficinasPreValidacion.frx":08CA
      ToolTipText     =   "Añadir responsable a prevalidar sus gastos"
      Top             =   4365
      Width           =   480
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Responsables"
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
      Left            =   6960
      TabIndex        =   11
      Top             =   4200
      Width           =   1320
   End
   Begin VB.Image imgCrearOficina 
      Height          =   480
      Left            =   5040
      Picture         =   "frmGestionOficinasPreValidacion.frx":1194
      ToolTipText     =   "Añadir oficina de prevalidación"
      Top             =   4365
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
      Left            =   600
      TabIndex        =   10
      Top             =   4200
      Width           =   735
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
      Left            =   13320
      TabIndex        =   9
      Top             =   1080
      Width           =   2475
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Responsables que se prevalidarán"
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
      Left            =   6960
      TabIndex        =   8
      Top             =   1080
      Width           =   3330
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
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   2145
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   18240
      Picture         =   "frmGestionOficinasPreValidacion.frx":1A5E
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
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
      TabIndex        =   6
      Top             =   75
      Width           =   5415
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   14400
      Picture         =   "frmGestionOficinasPreValidacion.frx":1F6A
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmGestionOficinasPreValidacion.frx":2274
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
      Width           =   18945
   End
End
Attribute VB_Name = "frmGestionOficinasPreValidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moParrillaOficinas As New CFlexGrid
Dim moParrillaResponsables As New CFlexGrid
Dim moParrillaPreValidadores As New CFlexGrid

Dim mafoRegResponsables As ADODB.Recordset
Dim mafoRegPreValidadores As ADODB.Recordset
Dim mbResponsables As Boolean
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

Private Sub cmbResponsables_Change()

    If mbResponsables = False Then Exit Sub

    Dim sAux As String
    Dim iCont As Integer
    
    For iCont = 0 To cmbResponsables.ListCount - 1
        cmbResponsables.RemoveItem 0
    Next iCont
    
    If cmbResponsables.Text = "" Then
    
        While Not mafoRegResponsables.EOF
        
            sAux = Trim(mafoRegResponsables("Nombre"))
            If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegResponsables("Codigo") & ") " & sAux
            cmbResponsables.AddItem sAux
            cmbResponsables.ItemData(cmbResponsables.NewIndex) = mafoRegResponsables("Codigo")
            mafoRegResponsables.MoveNext
        
        Wend
    
    Else
    
        While Not mafoRegResponsables.EOF
        
            If InStr(1, UCase(Trim(mafoRegResponsables("Nombre"))), UCase(cmbResponsables.Text)) Then
                sAux = Trim(mafoRegResponsables("Nombre"))
                If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegResponsables("Codigo") & ") " & sAux
                cmbResponsables.AddItem sAux
                cmbResponsables.ItemData(cmbResponsables.NewIndex) = mafoRegResponsables("Codigo")
            End If
            
            mafoRegResponsables.MoveNext
        
        Wend
    
    End If
    mafoRegResponsables.MoveFirst

End Sub

Private Sub Form_Load()
    
    shRecuadroForm.Width = Me.Width
    shRecuadroForm.Height = Me.Height
    shRecuadroForm.Left = 0
    shRecuadroForm.Top = 0
    
    mbResponsables = False
    mbPreValidadores = False
    
    moParrillaOficinas.Create mfgParrillaOficinas, 2, 2
    moParrillaOficinas.Formatear 0, 0, "id"
    moParrillaOficinas.Formatear 1, 4500, "Oficinas prevalidación", 4
    moParrillaOficinas.SituarEn 0, 1
    
    moParrillaResponsables.Create mfgParrillaResponsables, 2, 2
    moParrillaResponsables.Formatear 0, 0, "Usuario"
    moParrillaResponsables.Formatear 1, 4500, "Responsables a prevalidar", 4
    moParrillaResponsables.SituarEn 0, 1
    
    moParrillaPreValidadores.Create mfgParrillaPrevalidadores, 2, 2
    moParrillaPreValidadores.Formatear 0, 0, "Usuario"
    moParrillaPreValidadores.Formatear 1, 4500, "Prevalidadores", 4
    moParrillaPreValidadores.SituarEn 0, 1
            
    RellenarOficinas
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub imgAgregarResposable_Click()
    
    Dim id_Oficina As Integer
    
    Me.MousePointer = DEFAULT

    If cmbResponsables.Text = "" Or cmbResponsables.ListIndex = -1 Then
        cmbResponsables.SetFocus
        Exit Sub
    End If

    id_Oficina = moParrillaOficinas.TextoLinCol(moParrillaOficinas.Grid.Row, 0)
    If Not bPreValidacion_AsignarResponsable(CodPaisTrabajo, cmbResponsables.ItemData(cmbResponsables.ListIndex), id_Oficina) Then
        MsgBox "Se ha producido un error y no se ha podido asignar el responsable a la oficina de prevalidación.", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    RellenarDatosOficina

End Sub

Private Sub imgAgregarResposable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub imgSalir_Click()
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub mfgParrillaOficinas_Click()
    
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
    If Not bPreValidacion_EliminarPrevalidador(CodPaisTrabajo, moParrillaPreValidadores.TextoLinCol(moParrillaResponsables.Grid.Row, 0)) Then
        MsgBox "Se ha producido un error y no se ha eliminado el registro seleccionado.", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    RellenarDatosOficina

End Sub

Private Sub mfgParrillaResponsables_Click()
    If moParrillaResponsables.bLineaActualVacia Then
        moParrillaResponsables.SituarEn 0, 1
        Exit Sub
    End If
    moParrillaResponsables.ResaltarLinea
End Sub

Private Sub mfgParrillaResponsables_DblClick()

    Dim id_Oficina  As Integer
    
    If moParrillaResponsables.bLineaActualVacia Then
        moParrillaResponsables.SituarEn 0, 1
        Exit Sub
    End If
    moParrillaResponsables.ResaltarLinea
    
    If MsgBox("Va a eliminar el registro seleccionado, ¿está seguro?", vbAbortRetryIgnore + vbQuestion + vbYesNo, "Confirmación") = vbNo Then Exit Sub
    
    id_Oficina = moParrillaOficinas.TextoLinCol(moParrillaOficinas.Grid.Row, 0)
    If Not bPreValidacion_EliminarAsignacionResponsableOficina(CodPaisTrabajo, id_Oficina, moParrillaResponsables.TextoLinCol(moParrillaResponsables.Grid.Row, 0)) Then
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
    moParrillaResponsables.Vaciar
    moParrillaResponsables.SituarEn 0, 1
    moParrillaPreValidadores.Vaciar
    moParrillaPreValidadores.SituarEn 0, 1
    
    txtNombreOficina = ""
    cmbResponsables.ListIndex = -1
    cmbResponsables = ""
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
    Dim lDepartamento As Long
    
    Set mafoRegResponsables = Nothing
    Set mafoRegPreValidadores = Nothing
    mbResponsables = False
    mbPreValidadores = False
    cmbResponsables.ListIndex = -1
    cmbResponsables = ""
    cmbPreValidadores.ListIndex = -1
    cmbPreValidadores = ""
    moParrillaResponsables.Vaciar
    moParrillaResponsables.SituarEn 0, 1
    moParrillaPreValidadores.Vaciar
    moParrillaPreValidadores.SituarEn 0, 1
    
    id_Oficina = moParrillaOficinas.TextoLinCol(moParrillaOficinas.Grid.Row, 0)
    
    If Not bPreValidacion_DameResponsablesAsignadosOficina(CodPaisTrabajo, id_Oficina, afoRegistro) Then
        CargarResponsables id_Oficina
        Set afoRegistro = Nothing
        Exit Sub
    End If
    lDepartamento = 0
    While Not afoRegistro.EOF
        '12-11-19 No vamos a obligar que los validadores o los responsables sean del mismo departamento
        'lDepartamento = afoRegistro("Departamento")
        moParrillaResponsables.TextoFila moParrillaResponsables.LineaLibre, Trim(afoRegistro("Usu_Responsable")), Trim(afoRegistro("Nombre"))
        afoRegistro.MoveNext
    Wend
    Set afoRegistro = Nothing
    moParrillaResponsables.SituarEn 0, 1
    
    CargarResponsables id_Oficina, lDepartamento
    
    If Not bPreValidacion_DamePreValidadoresAsignadosOficina(CodPaisTrabajo, id_Oficina, afoRegistro) Then
        CargarPreValidadores lDepartamento
        Set afoRegistro = Nothing
        Exit Sub
    End If
    While Not afoRegistro.EOF
        moParrillaPreValidadores.TextoFila moParrillaPreValidadores.LineaLibre, afoRegistro("Usuario"), Trim(afoRegistro("Nombre"))
        afoRegistro.MoveNext
    Wend
    Set afoRegistro = Nothing
    moParrillaPreValidadores.SituarEn 0, 1

    CargarPreValidadores lDepartamento

End Sub

Private Sub CargarResponsables(id_Oficina As Integer, Optional lDepartamento As Long = 0)

    Dim sAux As String

    Set mafoRegResponsables = Nothing
    cmbResponsables.ListIndex = -1
    cmbResponsables = ""
    mbResponsables = False
    
    If Not bPreValidacion_DameResponsables(CodPaisTrabajo, id_Oficina, lDepartamento, mafoRegResponsables) Then
        Set mafoRegResponsables = Nothing
        Exit Sub
    End If
    
    mbResponsables = True
    While Not mafoRegResponsables.EOF
    
        sAux = Trim(mafoRegResponsables("Nombre"))
        If CodPaisTrabajo = giPAIS_GES Then sAux = "(" & mafoRegResponsables("Codigo") & ") " & sAux
        cmbResponsables.AddItem sAux
        cmbResponsables.ItemData(cmbResponsables.NewIndex) = mafoRegResponsables("Codigo")
        
        mafoRegResponsables.MoveNext
        
    Wend
    mafoRegResponsables.MoveFirst

End Sub

Private Sub CargarPreValidadores(lDepartamento As Long)

    Dim sAux As String

    Set mafoRegPreValidadores = Nothing
    cmbPreValidadores.ListIndex = -1
    cmbPreValidadores = ""
    mbPreValidadores = False
    
'    If lDepartamento = 0 Then Exit Sub
    
    If Not bPreValidacion_DamePreValidadores(CodPaisTrabajo, lDepartamento, mafoRegPreValidadores) Then
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


