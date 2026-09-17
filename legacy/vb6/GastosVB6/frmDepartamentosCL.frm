VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDepartamentosCL 
   BorderStyle     =   0  'None
   Caption         =   "Departamentos"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11220
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
   Moveable        =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSoloActivos 
      Caption         =   "Solo departamentos activos"
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdActivoDesactivo 
      Caption         =   "Activo/Desactivo"
      Height          =   735
      Left            =   7800
      Picture         =   "frmDepartamentosCL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5600
      Width           =   1575
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   735
      Left            =   9480
      Picture         =   "frmDepartamentosCL.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5600
      Width           =   1575
   End
   Begin VB.TextBox txtCECO 
      Height          =   285
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   2
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdGuardarDepartamento 
      Caption         =   "Guardar"
      Height          =   735
      Left            =   6120
      Picture         =   "frmDepartamentosCL.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5600
      Width           =   1575
   End
   Begin VB.TextBox txtDepartamento 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   5880
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid mfgParrilla 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7858
      _Version        =   393216
   End
   Begin VB.Label lblActivo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "CECO"
      Height          =   195
      Index           =   0
      Left            =   4440
      TabIndex        =   8
      Top             =   5640
      Width           =   510
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Departamento"
      Height          =   195
      Index           =   13
      Left            =   360
      TabIndex        =   7
      Top             =   5640
      Width           =   1230
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   10560
      Picture         =   "frmDepartamentosCL.frx":1A5E
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Departamentos"
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
      Width           =   2760
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   5280
      Picture         =   "frmDepartamentosCL.frx":1F6A
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmDepartamentosCL.frx":2274
      ToolTipText     =   "Pica y mueve el formulario"
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
      Width           =   11220
   End
End
Attribute VB_Name = "frmDepartamentosCL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mafoRegistro As ADODB.Recordset

Dim moParrilla As New cFlexGrid2

Dim miCodDepartamentoNuevo As Integer

Dim cMouseW As New cMouseWheel

Private Sub CargarDepartamentos()

    Dim lColor As Long
    Dim iFilaLibre As Long
    Dim bEnc As Boolean

    moParrilla.Vaciar
    moParrilla.SituarEn 0, 1

    miCodDepartamentoNuevo = 0
    If bDameDepartamentosCL(CodPaisTrabajo, mafoRegistro, -1) Then
        
        While Not mafoRegistro.EOF
        
            bEnc = True
            If chkSoloActivos.value Then If Trim(mafoRegistro("Activo")) = "N" Then bEnc = False
        
            If bEnc Then
                
                iFilaLibre = moParrilla.LineaLibre
                lColor = vbWhite
                If Trim(mafoRegistro("Activo")) = "N" Then lColor = &HFFC0FF
                
                moParrilla.TextoFila iFilaLibre, mafoRegistro("CodDepartamento"), Trim(mafoRegistro("Departamento")), mafoRegistro("CECO"), Trim(mafoRegistro("Divisa")), Trim(mafoRegistro("Activo"))
                moParrilla.ColorFila CLng(iFilaLibre), lColor
                
            End If
            
            If mafoRegistro("CodDepartamento") > miCodDepartamentoNuevo Then miCodDepartamentoNuevo = mafoRegistro("CodDepartamento")
            
            mafoRegistro.MoveNext
            
        Wend
        mafoRegistro.MoveFirst
        
    End If
    miCodDepartamentoNuevo = miCodDepartamentoNuevo + 1
        
End Sub

Private Sub chkSoloActivos_Click()

    LimpiarCampos
    CargarDepartamentos

End Sub

Private Sub cmdActivoDesactivo_Click()

    Dim sActivo As String
    Dim afoRegAux As ADODB.Recordset
    
    If lblCodigo.Caption <> "" Then
    
        If lblActivo.Caption = "N" Then
        'El departamento está desactivado, con lo que puedo activarlo sin ningún tipo de comprobación
            sActivo = "S"
        Else
        'El departamento está activo, antes de desactivarlo comprobamos que no tenga usuarios activos
            
            If bDameEmpleadosDepartamento(CodPaisTrabajo, CInt(lblCodigo.Caption), afoRegAux) Then
                Set afoRegAux = Nothing
                MsgBox "El departamento tiene usuarios asignados y no puede desactivarse.", vbApplicationModal + vbExclamation, "Información"
                Exit Sub
            End If
            Set afoRegAux = Nothing
            
            sActivo = "N"
        
        End If
        
        bActivarDesactivarDepartamento CodPaisTrabajo, CInt(lblCodigo.Caption), sActivo
        LimpiarCampos
        CargarDepartamentos
        
    End If
    
End Sub

Private Sub cmdGuardarDepartamento_Click()

    Dim sNombreDepartamento As String
    Dim sCECO As String
    
    sNombreDepartamento = Trim(LimpiarCadena(txtDepartamento))
    If sNombreDepartamento = "" Then
        txtDepartamento.SetFocus
        Exit Sub
    End If
    sCECO = Trim(LimpiarCadena(txtCECO))
    If sCECO = "" Then sCECO = "0"
    If IsNumeric(sCECO) = False Then
        txtCECO.SetFocus
        Exit Sub
    End If
    
    If lblCodigo.Caption <> "" Then
    
        While Not mafoRegistro.EOF
            If LCase(sNombreDepartamento) = LCase(Trim(mafoRegistro("Departamento"))) And CInt(lblCodigo.Caption) <> mafoRegistro("CodDepartamento") Then
                MsgBox "El nombre de departamento ya existe en el sistema con otro código.", vbApplicationModal + vbExclamation, "Información"
                txtDepartamento.SetFocus
                Exit Sub
            End If
            mafoRegistro.MoveNext
        Wend
        mafoRegistro.MoveFirst
    
        If bModificarDepartamento(CodPaisTrabajo, CInt(lblCodigo.Caption), sNombreDepartamento, CLng(sCECO)) Then
            LimpiarCampos
            CargarDepartamentos
            txtDepartamento.SetFocus
        Else
            MsgBox "No se ha podido crear el departamento.", vbApplicationModal + vbExclamation, "Error"
            txtDepartamento.SetFocus
        End If
    
    Else
    
        While Not mafoRegistro.EOF
            If LCase(sNombreDepartamento) = LCase(Trim(mafoRegistro("Departamento"))) Then
                MsgBox "El departamento ya existe en el sistema.", vbApplicationModal + vbExclamation, "Información"
                txtDepartamento.SetFocus
                Exit Sub
            End If
            mafoRegistro.MoveNext
        Wend
        mafoRegistro.MoveFirst
    
        If bInsertarNuevoDepartamento(CodPaisTrabajo, miCodDepartamentoNuevo, sNombreDepartamento, CLng(sCECO), "") Then
            LimpiarCampos
            CargarDepartamentos
            txtDepartamento.SetFocus
        Else
            MsgBox "No se ha podido crear el departamento.", vbApplicationModal + vbExclamation, "Error"
            txtDepartamento.SetFocus
        End If
    
    End If
    
End Sub

Private Sub cmdLimpiar_Click()

    LimpiarCampos
    txtDepartamento.SetFocus

End Sub

Private Sub LimpiarCampos()
    lblCodigo.Caption = ""
    txtDepartamento.Text = ""
    txtCECO.Text = ""
    lblActivo.Caption = ""
End Sub
Private Sub Form_Load()

    shRecuadroForm.Width = Me.Width
    shRecuadroForm.Height = Me.Height
    shRecuadroForm.Left = 0
    shRecuadroForm.Top = 0
    
    moParrilla.Create mfgParrilla, 5, 2
    moParrilla.Formatear 0, 0, "Codigo"
    moParrilla.Formatear 1, 5000, "Departamento", 4
    moParrilla.Formatear 2, 1500, "CECO", 4, 4
    moParrilla.Formatear 3, 1500, "Divisa", 4
    moParrilla.Formatear 4, 1000, "Activo", 4, 4
    moParrilla.SituarEn 0, 1
    
    CargarDepartamentos
    
    cMouseW.InitMouseWheel mfgParrilla.hwnd
    
    If CodPaisTrabajo = giPAIS_GES Then
        txtDepartamento.Enabled = False
        txtCECO.Enabled = False
        cmdGuardarDepartamento.Enabled = False
        cmdActivoDesactivo.Enabled = False
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub imgSalir_Click()
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub mfgParrilla_Click()
    If moParrilla.bLineaActualVacia Then
        moParrilla.SituarEn 0, 1
        Exit Sub
    End If
    moParrilla.ResaltarLinea
End Sub

Private Sub mfgParrilla_DblClick()

    Dim iFila As Long
    
    If moParrilla.bLineaActualVacia Then
        moParrilla.SituarEn 0, 1
        Exit Sub
    End If
    
    iFila = moParrilla.Grid.Row
    lblCodigo.Caption = moParrilla.TextoLinCol(iFila, 0)
    txtDepartamento.Text = moParrilla.TextoLinCol(iFila, 1)
    txtCECO.Text = moParrilla.TextoLinCol(iFila, 2)
    lblActivo.Caption = moParrilla.TextoLinCol(iFila, 4)
    
    moParrilla.ResaltarLinea
    
End Sub

Private Sub txtCECO_GotFocus()
    TextGotFocus txtCECO
End Sub

Private Sub txtCECO_KeyPress(KeyAscii As Integer)
    EntraNumero KeyAscii, txtCECO, 10, 0
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtCECO_LostFocus()
    TextLostFocus txtCECO
End Sub

Private Sub txtDepartamento_GotFocus()
    TextGotFocus txtDepartamento
End Sub

Private Sub txtDepartamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtDepartamento_LostFocus()
    TextLostFocus txtDepartamento
End Sub
