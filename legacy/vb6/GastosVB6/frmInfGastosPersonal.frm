VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfGastosPersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Gastos Internos"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfGastosPersonal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInforme 
      Caption         =   "Tipo de informe"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   7215
      Begin VB.CheckBox chkUsuarios 
         Caption         =   "Desglosar por usuario"
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cmbTipoComida 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cmbDepartamento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   920
         Width           =   4215
      End
      Begin VB.ComboBox cmbTiposGasto 
         Height          =   315
         ItemData        =   "frmInfGastosPersonal.frx":08CA
         Left            =   2400
         List            =   "frmInfGastosPersonal.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   440
         Width           =   3135
      End
      Begin VB.OptionButton optDepartamento 
         Caption         =   "Por departamento"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optTipoGasto 
         Caption         =   "Por tipo de gasto"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Label lblTipoInforme 
         BackColor       =   &H008080FF&
         Caption         =   "1"
         Height          =   135
         Left            =   4080
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   615
      Left            =   6120
      Picture         =   "frmInfGastosPersonal.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2325
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   615
      Left            =   4680
      Picture         =   "frmInfGastosPersonal.frx":0E58
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2325
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFechaDesde 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   20381699
      CurrentDate     =   38755
   End
   Begin MSComCtl2.DTPicker dtpFechaHasta 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   20381699
      CurrentDate     =   38755
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   2565
      Width           =   540
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "hasta"
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   8
      Top             =   2565
      Width           =   465
   End
End
Attribute VB_Name = "frmInfGastosPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mcltTiposGasto As Collection

Private Sub cmbTiposGasto_Click()
    
    If cmbTiposGasto.ListIndex = 2 Then
        cmbTipoComida.ListIndex = 0
        cmbTipoComida.Visible = True
    Else
        cmbTipoComida.Visible = False
    End If
    
End Sub

Private Sub cmdConsultar_Click()

    On Error GoTo Error_Rutina
    
    Select Case lblTipoInforme
    Case "1"
        InformePorTipoGasto
    Case "2"
        InformePorDepartamento
    Case Else
        
    End Select

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub InformePorDepartamento()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    Dim sFichero As String
    Dim iFichero As Integer
    Dim sLinea As String
    
    Dim dGastoGasolina As Single
    Dim dGastoKilometros As Single
    Dim dGastoComidasInvitaciones As Single
    Dim dGastoSalones As Single
    Dim dGastoHoteles As Single
    Dim dGastoTelefono As Single
    Dim dGastoBilletes As Single
    Dim dGastoTaxis As Single
    Dim dGastoPeajes As Single
    Dim dGastoOtrosGastos As Single
    Dim dGastoParking As Single
    Dim dGastoMuestras As Single
    Dim dGastoFormacion As Single
    Dim iDepartamento As Integer
    Dim sDepartamento As String
    Dim iCodigoUsuario As Integer
    Dim sCodigoUsuario As String
        
    On Error GoTo Error_Rutina

    If cmbDepartamento.Text = "" Then
        cmbDepartamento.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = HOURGLASS
    
    sLinea = ""
    If Mid(cmbDepartamento.Text, 1, 1) <> "*" Then sLinea = "AND g.Departamento = " & Trim(Mid(cmbDepartamento.Text, 1, InStr(1, cmbDepartamento.Text, "-") - 1)) & " "
    
    If chkUsuarios.Value = vbChecked Then
        sSql = "SELECT g.TipoGasto,SUM(g.Importe) SumImporte,g.Departamento,d.Descripcion,g.Codigo CodUsu,u.Nombre Usu " & _
               "FROM Gastos_Internos g,DepartamentosCL d,Usuarios_Personal u " & _
               "WHERE g.CodPais = d.CodPais AND g.CodPais = d.CodPais AND g.CodPais = " & CodPaisTrabajo & _
               " AND g.Departamento = d.Codigo AND d.Codigo = u.Departamento " & _
               "AND g.Codigo = u.Codigo " & _
               "AND Fecha BETWEEN '" & Format(dtpFechaDesde, "ddMMyyyy") & _
               "' AND '" & Format(dtpFechaHasta, "ddMMyyyy") & _
               "' " & sLinea & "AND g.TipoGasto <> 2 " & _
               "GROUP BY g.TipoGasto,g.Departamento,d.Descripcion,g.Codigo,u.Nombre " & _
               "ORDER BY d.Descripcion,g.Codigo,g.TipoGasto"
    Else
        sSql = "SELECT g.TipoGasto,SUM(g.Importe) SumImporte,g.Departamento,d.Descripcion,0 CodUsu,'usu' Usu " & _
               "FROM Gastos_Internos g,DepartamentosCL d " & _
               "WHERE g.CodPais = d.CodPais AND g.CodPais = " & CodPaisTrabajo & " AND g.Departamento = d.Codigo " & _
               "AND Fecha BETWEEN '" & Format(dtpFechaDesde, "ddMMyyyy") & _
               "' AND '" & Format(dtpFechaHasta, "ddMMyyyy") & _
               "' " & sLinea & "AND g.TipoGasto <> 2 " & _
               "GROUP BY g.TipoGasto,g.Departamento,d.Descripcion " & _
               "ORDER BY d.Descripcion,g.TipoGasto"
    End If
    'Excluimos Kilometros al ser el importe=importe*numero, y lo tenemos que calcular aparte
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Screen.MousePointer = DEFAULT
        MsgBox "No se han encontrado datos con las condiciones introducidas.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    
    sFichero = Dir("C:\Listados_GI", vbDirectory)
    If sFichero = "" Then MkDir "C:\Listados_GI"
    
    sFichero = "C:\Listados_GI\Lis_GI_GD_" & QuitarCaracteres(QuitarCaracteres(Trim(Mid(cmbDepartamento.Text, InStr(1, cmbDepartamento.Text, "-") + 1)), " ", ""), ":", "_") & "_" & Format(dtpFechaDesde, "ddMMyy") & "_" & Format(dtpFechaHasta, "ddMMyy") & ".csv"
    iFichero = FreeFile
    Open sFichero For Output As #iFichero
    sLinea = "Departamento;Gasolina;Kilómetros;Comidas/Invitaciones;Salones;Hoteles;Teléfono;Billetes;Taxis;Peajes;Parking;Otros Gastos;Muestras"
    If chkUsuarios.Value = vbChecked Then sLinea = "Departamento;Usuario;Gasolina;Kilómetros;Comidas/Invitaciones;Salones;Hoteles;Teléfono;Billetes;Taxis;Peajes;Parking;Otros Gastos;Muestras;Formación"
    Print #iFichero, sLinea
    
    dGastoGasolina = 0
    dGastoKilometros = 0
    dGastoComidasInvitaciones = 0
    dGastoSalones = 0
    dGastoHoteles = 0
    dGastoTelefono = 0
    dGastoBilletes = 0
    dGastoTaxis = 0
    dGastoPeajes = 0
    dGastoParking = 0
    dGastoOtrosGastos = 0
    dGastoMuestras = 0
    dGastoFormacion = 0
    iDepartamento = -1
    sDepartamento = ""
    iCodigoUsuario = -1
    sCodigoUsuario = ""
    While Not afoRegistro.EOF
    
        If afoRegistro("Departamento") <> iDepartamento Or afoRegistro("CodUsu") <> iCodigoUsuario Then
            
            If iDepartamento <> -1 Then
               sLinea = sDepartamento & ";" & dGastoGasolina & ";" & dGastoKilometros & ";" & dGastoComidasInvitaciones & ";" & dGastoSalones & ";" & dGastoHoteles & ";" & dGastoTelefono & ";" & dGastoBilletes & ";" & dGastoTaxis & ";" & dGastoPeajes & ";" & dGastoParking & ";" & dGastoOtrosGastos & ";" & dGastoMuestras
               If chkUsuarios.Value = vbChecked Then sLinea = sDepartamento & ";" & sCodigoUsuario & ";" & dGastoGasolina & ";" & dGastoKilometros & ";" & dGastoComidasInvitaciones & ";" & dGastoSalones & ";" & dGastoHoteles & ";" & dGastoTelefono & ";" & dGastoBilletes & ";" & dGastoTaxis & ";" & dGastoPeajes & ";" & dGastoParking & ";" & dGastoOtrosGastos & ";" & dGastoMuestras & ";" & dGastoFormacion
               Print #iFichero, sLinea
            Else
                If iCodigoUsuario <> -1 And iCodigoUsuario <> 0 Then
                    sLinea = sDepartamento & ";" & sCodigoUsuario & ";" & dGastoGasolina & ";" & dGastoKilometros & ";" & dGastoComidasInvitaciones & ";" & dGastoSalones & ";" & dGastoHoteles & ";" & dGastoTelefono & ";" & dGastoBilletes & ";" & dGastoTaxis & ";" & dGastoPeajes & ";" & dGastoParking & ";" & dGastoOtrosGastos & ";" & dGastoMuestras
                    Print #iFichero, sLinea
                End If
            End If
            
            dGastoGasolina = 0
            dGastoComidasInvitaciones = 0
            dGastoSalones = 0
            dGastoHoteles = 0
            dGastoTelefono = 0
            dGastoBilletes = 0
            dGastoTaxis = 0
            dGastoPeajes = 0
            dGastoParking = 0
            dGastoOtrosGastos = 0
            dGastoMuestras = 0
            dGastoFormacion = 0
            iDepartamento = afoRegistro("Departamento")
            sDepartamento = Trim(afoRegistro("Descripcion"))
            iCodigoUsuario = afoRegistro("CodUsu")
            sCodigoUsuario = Trim(afoRegistro("Usu"))
            dGastoKilometros = -1
            
        End If
        If dGastoKilometros = -1 Then dGastoKilometros = dCalcularGastoKilometros(iDepartamento, iCodigoUsuario)
        
        Select Case afoRegistro("TipoGasto")
        Case 1
            dGastoGasolina = afoRegistro("SumImporte")
        Case 3
            dGastoComidasInvitaciones = afoRegistro("SumImporte")
        Case 4
            dGastoSalones = afoRegistro("SumImporte")
        Case 5
            dGastoHoteles = afoRegistro("SumImporte")
        Case 6
            dGastoTelefono = afoRegistro("SumImporte")
        Case 7
            dGastoBilletes = afoRegistro("SumImporte")
        Case 8
            dGastoTaxis = afoRegistro("SumImporte")
        Case 9
            dGastoPeajes = afoRegistro("SumImporte")
        Case 10
            dGastoOtrosGastos = afoRegistro("SumImporte")
        Case 11
            dGastoParking = afoRegistro("SumImporte")
        Case 12
            dGastoMuestras = afoRegistro("SumImporte")
        Case 13
            dGastoFormacion = afoRegistro("SumImporte")
        End Select
        
        afoRegistro.MoveNext
    
    Wend
    sLinea = sDepartamento & ";" & dGastoGasolina & ";" & dGastoKilometros & ";" & dGastoComidasInvitaciones & ";" & dGastoSalones & ";" & dGastoHoteles & ";" & dGastoTelefono & ";" & dGastoBilletes & ";" & dGastoTaxis & ";" & dGastoPeajes & ";" & dGastoParking & ";" & dGastoOtrosGastos & ";" & dGastoMuestras & ";" & dGastoFormacion
    If chkUsuarios.Value = vbChecked Then sLinea = sDepartamento & ";" & sCodigoUsuario & ";" & dGastoGasolina & ";" & dGastoKilometros & ";" & dGastoComidasInvitaciones & ";" & dGastoSalones & ";" & dGastoHoteles & ";" & dGastoTelefono & ";" & dGastoBilletes & ";" & dGastoTaxis & ";" & dGastoPeajes & ";" & dGastoParking & ";" & dGastoOtrosGastos & ";" & dGastoMuestras & ";" & dGastoFormacion
    Print #iFichero, sLinea
    
    Set afoRegistro = Nothing
    Close #iFichero
    Screen.MousePointer = DEFAULT
    MsgBox "Listado creado correctamente en '" & sFichero & "'.", vbApplicationModal + vbInformation, "Información"
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    Screen.MousePointer = DEFAULT
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Function dCalcularGastoKilometros(iDepartamento As Integer, _
    iCodigoUsuario As Integer) As Single

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    On Error GoTo Error_Rutina

    dCalcularGastoKilometros = 0
    sSql = "SELECT SUM(Importe*Numero) SumImporte " & _
           "FROM Gastos_Internos " & _
           "WHERE Fecha BETWEEN '" & Format(dtpFechaDesde, "ddMMyyyy") & _
           "' AND '" & Format(dtpFechaHasta, "ddMMyyyy") & _
           "' AND TipoGasto = 2 AND Departamento = " & iDepartamento
    If iCodigoUsuario <> 0 Then sSql = sSql & " AND Codigo = " & iCodigoUsuario
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Function
    End If
    If IsNull(afoRegistro("SumImporte")) Then
        Set afoRegistro = Nothing
        Exit Function
    End If
    dCalcularGastoKilometros = afoRegistro("SumImporte")
    Set afoRegistro = Nothing

Exit_Rutina:
    Exit Function
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Function

Private Sub InformePorTipoGasto()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    Dim iFichero As Integer
    Dim sFichero As String
    Dim sLinea As String
    Dim bPrimeraLinea As Boolean
    Dim sLineaTitulo As String
    
    On Error GoTo Error_Rutina

    If cmbTiposGasto = "" Then
        cmbTiposGasto.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = HOURGLASS
    sSql = "SELECT g.TipoGasto,g.Tipo_Com_Inv,g.Fecha,g.Importe,g.Numero,g.Numero2,g.Donde," & _
           "g.Visa,g.Validado,g.Pagado,g.Moneda,u.Nombre,d.Descripcion,g.Factura,NVL(g.Personas,'') Personas,NVL(e.Empresa,'') Empresa " & _
           "FROM Gastos_Internos g,Usuarios_Personal u,DepartamentosCL d,OUTER Gastos_Empresas e " & _
           "WHERE g.CodPais = u.CodPais AND g.Codigo = u.Codigo AND " & _
           "g.Jerarquia = u.Jerarquia AND g.Departamento = u.Departamento AND " & _
           "g.Codpais = d.CodPais AND g.Departamento = d.Codigo AND " & _
           "Fecha BETWEEN '" & Format(dtpFechaDesde, "ddMMyyyy") & "' AND '" & Format(dtpFechaHasta, "ddMMyyyy") & _
           "' AND g.CodPais = " & CodPaisTrabajo & " AND TipoGasto = " & mcltTiposGasto(cmbTiposGasto.Text) & _
           " AND NVL(g.id_Empresa,0) = e.id"
    If cmbTipoComida.Visible Then
        Select Case cmbTipoComida.ListIndex
        Case 1
            sSql = sSql & " AND Tipo_Com_Inv = 'I'"
        Case 2
            sSql = sSql & " AND Tipo_Com_Inv = 'N'"
        End Select
    End If
    
    If goUsuario.IdUsuario = "ldelgado" Then sSql = sSql & " AND g.Departamento = 1201"
    
    
    sSql = sSql & " ORDER BY g.Fecha,d.Descripcion,u.Nombre"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Screen.MousePointer = DEFAULT
        MsgBox "No se han encontrado datos con las condiciones introducidas.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    
    sFichero = Dir("C:\Listados_GI", vbDirectory)
    If sFichero = "" Then MkDir "C:\Listados_GI"
    
    sFichero = "C:\Listados_GI\Lis_GI_TG_" & cmbTiposGasto.Text & "_" & Format(dtpFechaDesde, "ddMMyy") & "_" & Format(dtpFechaHasta, "ddMMyy") & ".csv"
    iFichero = FreeFile
    Open sFichero For Output As #iFichero
    bPrimeraLinea = True
    While Not afoRegistro.EOF
    
        Select Case cmbTiposGasto.ListIndex
        Case 0  'gasolina
            sLineaTitulo = "Fecha;Departamento;Usuario;Itinerario;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Importe") & ";" & Trim(afoRegistro("Moneda")) & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 1  'kilómetros
            sLineaTitulo = "Fecha;Departamento;Usuario;Itinerario;Kilómetros;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Numero") & ";" & afoRegistro("Importe") & ";" & Trim(afoRegistro("Moneda")) & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 2  'comidas/invitaciones
            sLineaTitulo = "Fecha;Departamento;Usuario;Donde/Concepto/Tipo Comida;Num. Personas;Personas;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Numero") & ";" & Trim(afoRegistro("Personas")) & ";" & afoRegistro("Importe") & ";" & Trim(afoRegistro("Moneda")) & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 3  'salones
            sLineaTitulo = "Fecha;Departamento;Usuario;Donde;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Importe") & ";" & Trim(afoRegistro("Moneda")) & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 4  'hoteles
            sLineaTitulo = "Fecha;Departamento;Usuario;Donde;Noches;Personas;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Numero") & ";" & afoRegistro("Numero2") & ";" & afoRegistro("Importe") & ";" & Trim(afoRegistro("Moneda")) & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 5  'telefono
            sLineaTitulo = "Fecha;Departamento;Usuario;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & afoRegistro("Importe") & ";" & Trim(afoRegistro("Moneda")) & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 6  'billetes
            sLineaTitulo = "Fecha;Departamento;Usuario;Itinerario;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Importe") & ";" & Trim(afoRegistro("Moneda")) & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 7  'taxis
            sLineaTitulo = "Fecha;Departamento;Usuario;Itinerario;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Importe") & ";" & Trim(afoRegistro("Moneda")) & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 8  'peajes
            sLineaTitulo = "Fecha;Departamento;Usuario;Itinerario;Tickets;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Numero") & ";" & afoRegistro("Importe") & ";" & Trim(afoRegistro("Moneda")) & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 10  'parking
            sLineaTitulo = "Fecha;Departamento;Usuario;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Importe") & ";" & Trim(afoRegistro("Moneda")) & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 11  'Otros gastos
            sLineaTitulo = "Fecha;Departamento;Usuario;Concepto;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Importe") & ";" & afoRegistro("Moneda") & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 12  'Muestras
            sLineaTitulo = "Fecha;Departamento;Usuario;Concepto;Importe;Numero;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Importe") & ";" & afoRegistro("Numero") & ";" & afoRegistro("Moneda") & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        Case 13  'Formación
            sLineaTitulo = "Fecha;Departamento;Usuario;Concepto;Importe;Moneda;VISA;Validado;Pagado;Factura;Empresa"
            sLinea = Format(afoRegistro("Fecha"), "dd/MM/yy") & ";" & Trim(afoRegistro("Descripcion")) & ";" & Trim(afoRegistro("Nombre")) & ";" & Trim(afoRegistro("Donde")) & ";" & afoRegistro("Importe") & ";" & afoRegistro("Moneda") & ";" & NoNull(afoRegistro("VISA"), "T") & ";" & NoNull(afoRegistro("Validado"), "T") & ";" & NoNull(afoRegistro("Pagado"), "T")
        End Select
        sLinea = sLinea & ";" & NoNull(afoRegistro("Factura"), "A")
        If bPrimeraLinea Then
            bPrimeraLinea = False
            Print #iFichero, sLineaTitulo
        End If
        sLinea = sLinea & ";" & Trim(NoNull(afoRegistro("Empresa"), "T"))
        Print #iFichero, sLinea
        
        afoRegistro.MoveNext
        
    Wend
    Set afoRegistro = Nothing
    Close #iFichero
    Screen.MousePointer = DEFAULT
    MsgBox "Listado creado correctamente en '" & sFichero & "'.", vbApplicationModal + vbInformation, "Información"

Exit_Rutina:
    Exit Sub
Error_Rutina:
    Screen.MousePointer = DEFAULT
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdVolver_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    On Error GoTo Error_Rutina
    
    dtpFechaDesde = Now
    dtpFechaHasta = Now
'    dtpFechaDesde.MaxDate = Now
'    dtpFechaHasta.MaxDate = Now
    
    Set mcltTiposGasto = New Collection
    mcltTiposGasto.Add 1, "Gasolina"
    cmbTiposGasto.AddItem "Gasolina"
    mcltTiposGasto.Add 2, "Kilómetros"
    cmbTiposGasto.AddItem "Kilómetros"
    mcltTiposGasto.Add 3, "Comidas-Invitaciones"
    cmbTiposGasto.AddItem "Comidas-Invitaciones"
    mcltTiposGasto.Add 4, "Salones"
    cmbTiposGasto.AddItem "Salones"
    mcltTiposGasto.Add 5, "Hoteles"
    cmbTiposGasto.AddItem "Hoteles"
    mcltTiposGasto.Add 6, "Teléfono"
    cmbTiposGasto.AddItem "Teléfono"
    mcltTiposGasto.Add 7, "Billetes"
    cmbTiposGasto.AddItem "Billetes"
    mcltTiposGasto.Add 8, "Taxis"
    cmbTiposGasto.AddItem "Taxis"
    mcltTiposGasto.Add 9, "Peajes"
    cmbTiposGasto.AddItem "Peajes"
    mcltTiposGasto.Add 11, "Parking"
    cmbTiposGasto.AddItem "Parking"
    mcltTiposGasto.Add 10, "Otros Gastos"
    cmbTiposGasto.AddItem "Otros Gastos"
    mcltTiposGasto.Add 12, "Muestras"
    cmbTiposGasto.AddItem "Muestras"
    mcltTiposGasto.Add 13, "Formación"
    cmbTiposGasto.AddItem "Formación"
    
    cmbTipoComida.AddItem "Todas"
    cmbTipoComida.AddItem "Internas"
    cmbTipoComida.AddItem "Negocios"
    
    sSql = "SELECT Codigo,Descripcion FROM DepartamentosCL " & _
           "WHERE CodPais = " & CodPaisTrabajo
    If goUsuario.IdUsuario = "ldelgado" Then sSql = sSql & " AND Codigo = 1201"
    sSql = sSql & " ORDER BY Descripcion"
    
    If goUsuario.IdUsuario <> "ldelgado" Then cmbDepartamento.AddItem "* - Todos"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        While Not afoRegistro.EOF
            cmbDepartamento.AddItem afoRegistro("Codigo") & " - " & Trim(afoRegistro("Descripcion"))
            afoRegistro.MoveNext
        Wend
    End If
    Set afoRegistro = Nothing

Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub optDepartamento_Click()
    cmbTiposGasto.ListIndex = -1
    cmbTiposGasto.Enabled = False
    cmbDepartamento.Enabled = True
    chkUsuarios.Enabled = True
    cmbTipoComida.Visible = False
    lblTipoInforme = "2"
End Sub

Private Sub optTipoGasto_Click()
    cmbDepartamento.ListIndex = -1
    cmbDepartamento.Enabled = False
    chkUsuarios.Value = 0
    chkUsuarios.Enabled = False
    cmbTiposGasto.Enabled = True
    lblTipoInforme = "1"
End Sub
