VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUsuariosPersonal 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22500
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   22500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRestablecerClave 
      Caption         =   "Restablecer clave"
      Height          =   735
      Left            =   11040
      Picture         =   "frmUsuariosPersonal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox chkSoloActivos 
      Caption         =   "Solo departamentos activos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtCadenaEncriptada 
      Height          =   285
      Left            =   15600
      MaxLength       =   25
      TabIndex        =   36
      Top             =   4560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox txtCadenaEncriptar 
      Height          =   285
      Left            =   15600
      MaxLength       =   25
      TabIndex        =   35
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdEncriptarCadena 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Encriptar cadena"
      Height          =   375
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtCECO 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      MaxLength       =   25
      TabIndex        =   9
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtid_Epsilon 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   8
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtid_SAP 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   360
      MaxLength       =   15
      TabIndex        =   7
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdListadoUsuarios 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ListadoUsuarios"
      Height          =   375
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdCambiarDepartamento 
      Caption         =   "Cambiar de dpto."
      Height          =   735
      Left            =   17640
      Picture         =   "frmUsuariosPersonal.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox txtCargo 
      Height          =   285
      Left            =   8640
      MaxLength       =   20
      TabIndex        =   3
      Top             =   4920
      Width           =   4455
   End
   Begin VB.CommandButton cmdEncriptarDNI 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Encriptar DNI"
      Height          =   375
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtCuentaBancaria 
      Height          =   285
      Left            =   8640
      MaxLength       =   50
      TabIndex        =   6
      Top             =   5640
      Width           =   4455
   End
   Begin VB.ComboBox cmbJerarquia 
      Height          =   315
      ItemData        =   "frmUsuariosPersonal.frx":1594
      Left            =   2040
      List            =   "frmUsuariosPersonal.frx":15A1
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   735
      Left            =   15600
      Picture         =   "frmUsuariosPersonal.frx":15B0
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdReactivar 
      Caption         =   "Marcar como activo"
      Height          =   735
      Left            =   19920
      Picture         =   "frmUsuariosPersonal.frx":1E7A
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdBaja 
      Caption         =   "Marcar como baja"
      Height          =   735
      Left            =   19920
      Picture         =   "frmUsuariosPersonal.frx":2744
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdGuardarUsuario 
      Caption         =   "Guardar"
      Height          =   735
      Left            =   13560
      Picture         =   "frmUsuariosPersonal.frx":2EAE
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   3480
      MaxLength       =   50
      TabIndex        =   2
      Top             =   4920
      Width           =   4935
   End
   Begin VB.TextBox txtDNI 
      Height          =   285
      Left            =   360
      MaxLength       =   25
      TabIndex        =   4
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   360
      MaxLength       =   15
      TabIndex        =   0
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txteMail 
      Height          =   285
      Left            =   2520
      MaxLength       =   150
      TabIndex        =   5
      Top             =   5640
      Width           =   5895
   End
   Begin MSFlexGridLib.MSFlexGrid mfgDepartamentos 
      Height          =   3135
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5530
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid mfgUsuariosDepartamento 
      Height          =   3135
      Left            =   5640
      TabIndex        =   12
      Top             =   1200
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   5530
      _Version        =   393216
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "CECO"
      Height          =   195
      Index           =   12
      Left            =   3720
      TabIndex        =   33
      Top             =   6120
      Width           =   510
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "id_Epsilon"
      Height          =   195
      Index           =   11
      Left            =   2040
      TabIndex        =   32
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "id_SAP"
      Height          =   195
      Index           =   10
      Left            =   360
      TabIndex        =   31
      Top             =   6120
      Width           =   600
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "* Doble clik para modificar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   9
      Left            =   20280
      TabIndex        =   28
      Top             =   960
      Width           =   2010
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cargo"
      Height          =   195
      Index           =   7
      Left            =   8640
      TabIndex        =   27
      Top             =   4680
      Width           =   525
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta bancaria"
      Height          =   195
      Index           =   6
      Left            =   8640
      TabIndex        =   25
      Top             =   5400
      Width           =   1410
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   21960
      Picture         =   "frmUsuariosPersonal.frx":3778
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Usuarios gastos"
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
      TabIndex        =   20
      Top             =   75
      Width           =   2820
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   5280
      Picture         =   "frmUsuariosPersonal.frx":3C84
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmUsuariosPersonal.frx":3F8E
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
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Departamentos"
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
      Index           =   13
      Left            =   240
      TabIndex        =   19
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Usuarios"
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
      Left            =   5640
      TabIndex        =   18
      Top             =   960
      Width           =   825
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      Height          =   195
      Index           =   3
      Left            =   3480
      TabIndex        =   17
      Top             =   4680
      Width           =   675
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "DNI"
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   16
      Top             =   5400
      Width           =   330
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   15
      Top             =   4680
      Width           =   645
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Jerarquía"
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   14
      Top             =   4680
      Width           =   795
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "eMail"
      Height          =   195
      Index           =   5
      Left            =   2520
      TabIndex        =   13
      Top             =   5400
      Width           =   435
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
      Width           =   22500
   End
End
Attribute VB_Name = "frmUsuariosPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mafoRegistro As ADODB.Recordset

Dim moDepartamentos As New cFlexGrid2
Dim moUsuariosDepartamento As New cFlexGrid2

Dim cMouseW As New cMouseWheel

Private Sub CargarDepartamentos()
    
    Dim bEnc As Boolean
    Dim lColor As Long
    Dim iFilaLibre As Long
    
    moDepartamentos.Vaciar
    moDepartamentos.SituarEn 0, 1
    
    If bDameDepartamentosCL(CodPaisTrabajo, mafoRegistro, -1) Then
    
        While Not mafoRegistro.EOF
            
            bEnc = True
            If chkSoloActivos.value Then If Trim(mafoRegistro("Activo")) = "N" Then bEnc = False
        
            If bEnc Then
                
                iFilaLibre = moDepartamentos.LineaLibre
                lColor = vbWhite
                If Trim(mafoRegistro("Activo")) = "N" Then lColor = &HFFC0FF
                
                moDepartamentos.TextoFila iFilaLibre, mafoRegistro("CodDepartamento"), Trim(mafoRegistro("Departamento"))
                moDepartamentos.ColorFila CLng(iFilaLibre), lColor
                
            End If
            
            mafoRegistro.MoveNext
            
        Wend
        
    End If
    
End Sub

Private Sub CargarUsuariosDepartamento()

    Dim afoRegAux As ADODB.Recordset

    Dim sDNI As String
    Dim sDNIAux As String
    Dim sFechaBaja As String
    Dim sCuenta_Bancaria As String
    Dim iFilaLibre As Long
    Dim lColor As Long

    moUsuariosDepartamento.Vaciar
    moUsuariosDepartamento.SituarEn 0, 1

    If moDepartamentos.bLineaActualVacia Then Exit Sub
    
    If bDameUsuariosDepartamento(CodPaisTrabajo, moDepartamentos.TextoLinCol(moDepartamentos.Grid.Row, 0), mafoRegistro) Then
        While Not mafoRegistro.EOF
        
            sDNIAux = Trim(NoNull(mafoRegistro("DNI"), "A"))
            If sDNIAux <> "" Then sDNIAux = Encrypt_v2(sDNIAux, , True)
            
            If bDameDatos_UsuarioPersonal_BDGastos(CodPaisTrabajo, moDepartamentos.TextoLinCol(moDepartamentos.Grid.Row, 0), mafoRegistro("Codigo"), afoRegAux) Then
                sDNI = Trim(afoRegAux("DNI"))
                sCuenta_Bancaria = Trim(NoNull(afoRegAux("Cuenta_Bancaria"), "A"))
            End If
            If sDNI <> "" Then sDNI = Encrypt_v2(sDNI, , True)
            If sCuenta_Bancaria <> "" Then sCuenta_Bancaria = Encrypt_v2(sCuenta_Bancaria, , True)
            Set afoRegAux = Nothing
            
            sFechaBaja = ""
            lColor = vbWhite
            If IsNull(mafoRegistro("FechaBaja")) = False Then
                sFechaBaja = Format(mafoRegistro("FechaBaja"), "dd-MM-yy")
                lColor = &HFFC0FF
            End If
            
            
            iFilaLibre = moUsuariosDepartamento.LineaLibre
            moUsuariosDepartamento.TextoFila iFilaLibre, mafoRegistro("Codigo"), Trim(mafoRegistro("Usuario")), Trim(mafoRegistro("Nombre")), _
                Trim(mafoRegistro("Jerarquia")), Trim(mafoRegistro("Cargo")), sDNIAux, sDNI, Trim(mafoRegistro("Dir_eMail")), sFechaBaja, sCuenta_Bancaria, _
                NoNull(mafoRegistro("id_SAP"), "N"), NoNull(mafoRegistro("id_Epsilon"), "A"), NoNull(mafoRegistro("CECO"), "A")
            moUsuariosDepartamento.ColorFila CLng(iFilaLibre), lColor
            
            mafoRegistro.MoveNext
            
        Wend
    End If
    moDepartamentos.ResaltarLinea
    
End Sub

Private Sub chkSoloActivos_Click()

    CargarDepartamentos

End Sub

Private Sub cmdBaja_Click()

    If moDepartamentos.LineaSeleccionada = -1 Then
        MsgBox "No ha seleccionado un departamento.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    If lblCodigo.Caption = "" Then Exit Sub
    
    If bAlta_Baja_UsuariosPersonal(CodPaisTrabajo, CInt(moDepartamentos.TextoLinCol(moDepartamentos.LineaSeleccionada, 0)), CInt(lblCodigo.Caption), 0) Then
        LimpiarCampos
        MsgBox "Acción realizada correctamente.", vbApplicationModal + vbInformation, "Información"
        CargarUsuariosDepartamento
    Else
        MsgBox "No se ha podido llevar a cabo la acción solicitada.", vbApplicationModal + vbExclamation, "E R R O R"
    End If
    
End Sub

Private Sub cmdCambiarDepartamento_Click()

    Dim iCodigo As Integer
    Dim iDepartamentoOrigen As Integer
    Dim iDepartamentoDestino As Integer
    Dim sMensaje As String
    
    If moDepartamentos.LineaSeleccionada = -1 Then
        MsgBox "No ha seleccionado un departamento.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    iDepartamentoOrigen = CInt(moDepartamentos.TextoLinCol(moDepartamentos.LineaSeleccionada, 0))
    
    iCodigo = 0
    If lblCodigo.Caption <> "" Then iCodigo = CInt(lblCodigo.Caption)
    If iCodigo = 0 Then
        MsgBox "No ha seleccionado un usuario.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    
    If MsgBox("Va realizar un cambio de departamento con el usuario seleccionado, ¿está seguro?", vbApplicationModal + vbQuestion, "Confirmación") = vbNo Then
        Exit Sub
    End If
    
    'Pedimos departamento destino
    frmSeleccionarDepartamento.Show MODAL
    iDepartamentoDestino = frmSeleccionarDepartamento.CodDepartamento
    If iDepartamentoDestino = 0 Then
        Exit Sub
    End If
    
    If bCambiar_DepartamentoUsuarioPersonal(CodPaisTrabajo, iCodigo, iDepartamentoOrigen, iDepartamentoDestino, sMensaje) Then

        CargarUsuariosDepartamento
        
        LimpiarCampos
        txtUsuario.SetFocus

    Else
        MsgBox "No se ha podido realizar la acción solicitada. Error: " & sMensaje, vbApplicationModal + vbExclamation, "Error"
    End If

End Sub

Private Sub cmdEncriptarCadena_Click()

    txtCadenaEncriptada.Text = Encrypt(txtCadenaEncriptar.Text)
    
End Sub

Private Sub cmdEncriptarDNI_Click()

    Dim sDNI As String
    Dim sCuentaBancaria As String
    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
'
'    txtDNI.Text = ""
'    txtCuentaBancaria.Text = ""
'
'    sDNI = "80054057G"
'    sDNI = Encrypt_v2(sDNI)
'    sCuentaBancaria = "ES34-123456789"
'    sCuentaBancaria = Encrypt_v2(sCuentaBancaria)
'
'    sSql = "INSERT INTO pedidosweb_er:tlanchi values (34,10,12,'" & sDNI & "','" & sCuentaBancaria & "')"
'    gBdGALILEO.bExecute sSql
'
'    sSql = "SELECT * FROM PEDIDOSWEB_ER:TLANCHI"
'    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
'        txtDNI.Text = Encrypt_v2(Trim(afoRegistro("DNI")), , True)
'        txtCuentaBancaria.Text = Encrypt_v2(Trim(afoRegistro("Cuenta_Bancaria")), , True)
'    End If
'    Set afoRegistro = Nothing
    
    
    '**********************************
    
'    sSql = "SELECT CodPais,Departamento,Codigo,DNI FROM USUARIOS_PERSONAL"
'    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
'        Set afoRegistro = Nothing
'        MsgBox "No se han encontrado datos", vbApplicationModal + vbExclamation, "Información"
'        Exit Sub
'    End If
'    While Not afoRegistro.EOF
'
'        sSql = "INSERT INTO Usuarios_Personal (CodPais,Departamento,Codigo,DNI) " & _
'               "VALUES (" & afoRegistro("CodPais") & "," & afoRegistro("Departamento") & "," & afoRegistro("Codigo") & ",'" & afoRegistro("DNI") & "')"
'        If Not gBDGASTOS.bExecute(sSql) Then
'            sSql = sSql
'        End If
'        afoRegistro.MoveNext
'
'    Wend
'    Set afoRegistro = Nothing
'    MsgBox "Hecho"

'**********************************

Exit Sub


'    Dim sDNI As String
    
    sDNI = "20234202K"
    sDNI = Encrypt_v2(sDNI)
    sDNI = sDNI
    sDNI = Encrypt_v2(sDNI, , True)
    sDNI = sDNI
    
    Exit Sub


 '   Dim sSql As String
  '  Dim afoRegistro As ADODB.Recordset
    
    sSql = "SELECT * FROM Usuarios_Personal WHERE codpais=70 and codigo=1 and departamento=6"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
    
        While Not afoRegistro.EOF
            
            If Trim(afoRegistro("DNI")) <> "" Then
            
                sSql = Encrypt_v2(Trim(afoRegistro("DNI")))
                If InStr(1, sSql, "'") > 0 Or InStr(1, sSql, """") > 0 Then
                    sSql = sSql
                End If
                
                sSql = "UPDATE Usuarios_Personal SET DNI = '" & sSql & "' " & _
                       "WHERE CodPais = " & afoRegistro("CodPais") & " AND Codigo = " & afoRegistro("Codigo") & " " & _
                       "AND Departamento = " & afoRegistro("Departamento") & " AND Jerarquia = '" & Trim(afoRegistro("Jerarquia")) & "'"
                If Not gBdGALILEO.bExecute(sSql) Then
                    sSql = sSql
                End If
                
            End If
            afoRegistro.MoveNext
        
        Wend
    
    End If
    Set afoRegistro = Nothing
    MsgBox "Fin", vbApplicationModal + vbInformation, "Información"

End Sub

Private Sub cmdGuardarUsuario_Click()
    
    Dim sUsuario As String
    Dim sJerarquia As String
    Dim sNombre As String
    Dim sCargo As String
    Dim sDNI As String
    Dim seMail As String
    Dim sCuentaBancaria As String
    Dim sid_SAP As String
    Dim sid_Epsilon As String
    Dim sCECO As String
    Dim iCodigo As Integer
    Dim sMensaje As String
    
    If moDepartamentos.LineaSeleccionada = -1 Then
        MsgBox "No ha seleccionado un departamento.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    
    sUsuario = LimpiarCadena(Trim(txtUsuario.Text))
    sJerarquia = Trim(cmbJerarquia.Text)
    sNombre = LimpiarCadena(Trim(txtNombre.Text))
    sCargo = LimpiarCadena(Trim(txtCargo.Text))
    sDNI = LimpiarCadena(Trim(txtDNI.Text))
    seMail = LimpiarCadena(Trim(txteMail.Text))
    sCuentaBancaria = LimpiarCadena(Trim(txtCuentaBancaria.Text))
    sid_SAP = Trim(txtid_SAP)
    sid_Epsilon = LimpiarCadena(Trim(txtid_Epsilon))
    sCECO = LimpiarCadena(Trim(txtCECO))
    
    If sUsuario = "" Then
        txtUsuario.SetFocus
        Exit Sub
    End If
    If sJerarquia = "" Then
        cmbJerarquia.SetFocus
        Exit Sub
    End If
    If sNombre = "" Then
        txtNombre.SetFocus
        Exit Sub
    End If
    If sCargo = "" Then sCargo = "-"
    If sDNI = "" Then
        txtDNI.SetFocus
        Exit Sub
    End If
    If seMail = "" Then
        txteMail.SetFocus
        Exit Sub
    Else
        If ValidarMail(seMail) = False Then
            MsgBox "El eMail introducido no es válido.", vbApplicationModal + vbExclamation, "E R R O R"
            txteMail.SetFocus
            Exit Sub
        End If
    End If
    If sid_SAP = "" Then sid_SAP = "0"
    If Not IsNumeric(sid_SAP) Then sid_SAP = "0"
    If sid_Epsilon = "" Then sid_Epsilon = "0"
    'If Not IsNumeric(sid_Epsilon) Then sid_Epsilon = "0"
    If sCECO = "" Then sCECO = "0"
    'If Not IsNumeric(sCECO) Then sCECO = "0"
     
    iCodigo = 0
    If lblCodigo.Caption <> "" Then iCodigo = CInt(lblCodigo.Caption)
    If bGuardar_UsuarioPersonal(CodPaisTrabajo, iCodigo, CInt(moDepartamentos.TextoLinCol(moDepartamentos.LineaSeleccionada, 0)), sUsuario, sJerarquia, sNombre, sDNI, sCargo, seMail, sCuentaBancaria, sid_SAP, sid_Epsilon, sCECO, sMensaje) Then

        CargarUsuariosDepartamento
        
        LimpiarCampos
        txtUsuario.SetFocus

    Else
        MsgBox "No se ha podido realizar la acción solicitada. Error: " & sMensaje, vbApplicationModal + vbExclamation, "Error"
    End If
    
End Sub

Private Sub cmdLimpiar_Click()

    LimpiarCampos
    txtUsuario.SetFocus

End Sub

Private Sub LimpiarCampos()
    lblCodigo.Caption = ""
    txtUsuario.Text = ""
    txtUsuario.Enabled = True
    cmbJerarquia.ListIndex = 0
    txtNombre.Text = ""
    txtDNI.Text = ""
    txteMail.Text = ""
    txtCuentaBancaria.Text = ""
    txtCargo.Text = ""
    txtid_SAP.Text = "0"
    txtid_Epsilon.Text = "0"
    txtCECO.Text = "0"
    cmdRestablecerClave.Visible = False
End Sub

Private Sub cmdListadoUsuarios_Click()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    Dim afoRegAux As ADODB.Recordset
    Dim sDNI As String
    
    Dim iFichero As Integer
    
    sSql = "SELECT a.CodPais,b.Nombre NomPais,a.Departamento,c.Descripcion NomDepartamento,a.Codigo,a.Nombre NomCodigo " & _
           "FROM Usuarios_Personal a " & _
           "    INNER JOIN Paises b ON a.CodPais = b.CodPais " & _
           "    INNER JOIN DepartamentosCL c ON a.CodPais = c.CodPais AND a.Departamento = c.Codigo " & _
           "WHERE a.FechaBaja IS NULL AND a.CodPais <> 92 AND a.Departamento NOT IN (0,999) AND c.Descripcion <> 'Pagos' AND c.Descripcion <> 'Test' AND a.CodPais NOT IN (1,6,7,9,24,42,57,80,90,85,69,93)"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
    
        iFichero = FreeFile
        Open "C:\GeneracionCSV\UsuGastos.csv" For Output As #iFichero
        Print #iFichero, "CodPais;Nombre;Departamento;Nombre;codigo;Nombre;DNI"
    
        While Not afoRegistro.EOF
            
            sDNI = ""
            If bDameDatos_UsuarioPersonal_BDGastos(afoRegistro("CodPais"), afoRegistro("Departamento"), afoRegistro("Codigo"), afoRegAux) Then
                sDNI = Trim(afoRegAux("DNI"))
            End If
            If sDNI <> "" Then sDNI = Encrypt_v2(sDNI, , True)
            Set afoRegAux = Nothing
            
            Print #iFichero, afoRegistro("CodPais") & ";" & Trim(afoRegistro("NomPais")) & ";" & afoRegistro("Departamento") & ";" & Trim(afoRegistro("NomDepartamento")) & ";" & afoRegistro("Codigo") & ";" & Trim(afoRegistro("NomCodigo")) & ";" & sDNI
            
            afoRegistro.MoveNext
            
        Wend
        Close #iFichero
        
        MsgBox "Acabé", vbApplicationModal + vbInformation, "Información"
    
    Else
        MsgBox "No se han encontrado datos", vbApplicationModal + vbExclamation, "Información"
    End If
    Set afoRegistro = Nothing

End Sub

Private Sub cmdReactivar_Click()

    If moDepartamentos.LineaSeleccionada = -1 Then
        MsgBox "No ha seleccionado un departamento.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    If lblCodigo.Caption = "" Then Exit Sub
    
    If bAlta_Baja_UsuariosPersonal(CodPaisTrabajo, CInt(moDepartamentos.TextoLinCol(moDepartamentos.LineaSeleccionada, 0)), CInt(lblCodigo.Caption), 1) Then
        LimpiarCampos
        MsgBox "Acción realizada correctamente.", vbApplicationModal + vbInformation, "Información"
        CargarUsuariosDepartamento
    Else
        MsgBox "No se ha podido llevar a cabo la acción solicitada.", vbApplicationModal + vbExclamation, "E R R O R"
    End If
    
End Sub

Private Sub cmdRestablecerClave_Click()

    Dim iCodigo As Integer
    Dim sDNI As String
    Dim sMensaje As String
    
    If MsgBox("Va a restablecer la clave del usuario, ¿está seguro?", vbApplicationModal + vbQuestion + vbYesNo, "Confirmación") = vbNo Then Exit Sub

    If Trim(lblCodigo.Caption) = "" Then Exit Sub
    sDNI = LimpiarCadena(Trim(txtDNI.Text))
    If Len(Trim(txtDNI.Text)) < 8 Then
        MsgBox "El DNI no es válido por lo que no se puede restablecer la clave del usuario (debe tener al menos 8.", vbApplicationModal + vbExclamation, "Error"
        Exit Sub
    End If
    
    iCodigo = 0
    If lblCodigo.Caption <> "" Then iCodigo = CInt(lblCodigo.Caption)
        
    sMensaje = ""
    If bRestablecerClave_UsuarioPersonal(CodPaisTrabajo, iCodigo, CInt(moDepartamentos.TextoLinCol(moDepartamentos.LineaSeleccionada, 0)), sDNI, sMensaje) Then

        CargarUsuariosDepartamento
        
        LimpiarCampos
        txtUsuario.SetFocus

    Else
        MsgBox "No se ha podido realizar la acción solicitada. Error: " & sMensaje, vbApplicationModal + vbExclamation, "Error"
    End If

End Sub

Private Sub Form_Load()

    shRecuadroForm.Width = Me.Width
    shRecuadroForm.Height = Me.Height
    shRecuadroForm.Left = 0
    shRecuadroForm.Top = 0
    
    moDepartamentos.Create mfgDepartamentos, 2, 2
    moDepartamentos.Formatear 0, 0, "Codigo"
    moDepartamentos.Formatear 1, 4500, "Departamento", 4
    moDepartamentos.SituarEn 0, 1
    CargarDepartamentos
    
    moUsuariosDepartamento.Create mfgUsuariosDepartamento, 13, 2
    moUsuariosDepartamento.Formatear 0, 0, "Codigo"
    moUsuariosDepartamento.Formatear 1, 1500, "Usuario", 4
    moUsuariosDepartamento.Formatear 2, 5000, "Nombre", 4
    moUsuariosDepartamento.Formatear 3, 1250, "Jerarquia", 4, 4
    moUsuariosDepartamento.Formatear 4, 3000, "Cargo", 4
    moUsuariosDepartamento.Formatear 5, 0, "DNI I", 4
    moUsuariosDepartamento.Formatear 6, 1500, "DNI", 4
    moUsuariosDepartamento.Formatear 7, 5000, "email", 4
    moUsuariosDepartamento.Formatear 8, 1500, "Baja", 4, 4
    moUsuariosDepartamento.Formatear 9, 3000, "Cuenta bancaria", 4
    moUsuariosDepartamento.Formatear 10, 1500, "id_SAP", 4
    moUsuariosDepartamento.Formatear 11, 1500, "id_Epsilon", 4
    moUsuariosDepartamento.Formatear 12, 1500, "CECO", 4
    
    moUsuariosDepartamento.SituarEn 0, 1
        
    cMouseW.InitMouseWheel mfgDepartamentos.hwnd
    cMouseW.InitMouseWheel mfgUsuariosDepartamento.hwnd
    
    If CodPaisTrabajo = giPAIS_GES Then
    
        txtUsuario.Enabled = False
        cmbJerarquia.Enabled = False
        txtNombre.Enabled = False
        txtDNI.Enabled = False
        txteMail.Enabled = False
        txtCargo.Enabled = False
        txtCuentaBancaria.Enabled = False
        cmdGuardarUsuario.Enabled = False
        cmdLimpiar.Enabled = False
        cmdBaja.Enabled = False
        cmdReactivar.Enabled = False
        
    End If
    
    If goUsuario.IdUsuario = "lanchi" Then
        cmdEncriptarDNI.Visible = True
        cmdListadoUsuarios.Visible = True
        cmdEncriptarCadena.Visible = True
        txtCadenaEncriptar.Visible = True
        txtCadenaEncriptada.Visible = True
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

Private Sub mfgDepartamentos_Click()
    LimpiarCampos
    If moDepartamentos.bLineaActualVacia Then
        moDepartamentos.SituarEn 0, 1
        Exit Sub
    End If
    CargarUsuariosDepartamento
    moDepartamentos.ResaltarLinea
End Sub

Private Sub mfgUsuariosDepartamento_Click()
    If moUsuariosDepartamento.bLineaActualVacia Then
        moUsuariosDepartamento.SituarEn 0, 1
        Exit Sub
    End If
    moUsuariosDepartamento.ResaltarLinea
End Sub

Private Sub mfgUsuariosDepartamento_DblClick()
    
    Dim iFila As Long
    
    If moUsuariosDepartamento.bLineaActualVacia Then
        moUsuariosDepartamento.SituarEn 0, 1
        Exit Sub
    End If
    
    iFila = moUsuariosDepartamento.Grid.Row
    lblCodigo.Caption = moUsuariosDepartamento.TextoLinCol(iFila, 0)
    txtUsuario.Text = moUsuariosDepartamento.TextoLinCol(iFila, 1)
    txtUsuario.Enabled = False
    txtNombre.Text = moUsuariosDepartamento.TextoLinCol(iFila, 2)
    cmbJerarquia.Text = moUsuariosDepartamento.TextoLinCol(iFila, 3)
    txtCargo.Text = moUsuariosDepartamento.TextoLinCol(iFila, 4)
    txtDNI.Text = moUsuariosDepartamento.TextoLinCol(iFila, 6)
    txteMail.Text = moUsuariosDepartamento.TextoLinCol(iFila, 7)
    txtCuentaBancaria.Text = moUsuariosDepartamento.TextoLinCol(iFila, 9)
    txtid_SAP.Text = moUsuariosDepartamento.TextoLinCol(iFila, 10)
    txtid_Epsilon.Text = moUsuariosDepartamento.TextoLinCol(iFila, 11)
    txtCECO.Text = moUsuariosDepartamento.TextoLinCol(iFila, 12)
    
    cmdRestablecerClave.Visible = True
    
    moUsuariosDepartamento.ResaltarLinea

End Sub

Private Sub txtCargo_GotFocus()
    TextGotFocus txtCargo
End Sub

Private Sub txtCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtCargo_LostFocus()
    TextLostFocus txtCargo
End Sub

Private Sub txtCECO_GotFocus()
    TextGotFocus txtCECO
End Sub

Private Sub txtCECO_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtCECO_LostFocus()
    TextLostFocus txtCECO
End Sub

Private Sub txtCuentaBancaria_GotFocus()
    TextGotFocus txtCuentaBancaria
End Sub

Private Sub txtCuentaBancaria_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtCuentaBancaria_LostFocus()
    TextLostFocus txtCuentaBancaria
End Sub

Private Sub txtDNI_GotFocus()
    TextGotFocus txtDNI
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtDNI_LostFocus()
    TextLostFocus txtDNI
End Sub

Private Sub txteMail_GotFocus()
    TextGotFocus txteMail
End Sub

Private Sub txteMail_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txteMail_LostFocus()
    TextLostFocus txteMail
End Sub

Private Sub txtid_Epsilon_GotFocus()
    TextGotFocus txtid_Epsilon
End Sub

Private Sub txtid_Epsilon_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
End Sub

Private Sub txtid_Epsilon_LostFocus()
    TextLostFocus txtid_Epsilon
End Sub

Private Sub txtid_SAP_GotFocus()
    TextGotFocus txtid_SAP
End Sub

Private Sub txtid_SAP_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    EntraNumero KeyAscii, txtid_SAP, 15
End Sub

Private Sub txtid_SAP_LostFocus()
    TextLostFocus txtid_SAP
End Sub

Private Sub txtNombre_GotFocus()
    TextGotFocus txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtNombre_LostFocus()
    TextLostFocus txtNombre
End Sub

Private Sub txtUsuario_GotFocus()
    TextGotFocus txtUsuario
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtUsuario_LostFocus()
    TextLostFocus txtUsuario
End Sub

