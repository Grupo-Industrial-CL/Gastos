VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIntroducirGastosInternos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Introducir gastos"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIntroducirGastosInternos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContinuar 
      Caption         =   "Introducir gasto"
      Height          =   615
      Left            =   6360
      Picture         =   "frmIntroducirGastosInternos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   615
      Left            =   6360
      Picture         =   "frmIntroducirGastosInternos.frx":0EF4
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame fraDatosGasto 
      Enabled         =   0   'False
      Height          =   5655
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   7815
      Begin VB.TextBox txtPersonas 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   255
         TabIndex        =   18
         Top             =   5040
         Width           =   3375
      End
      Begin VB.TextBox txtProveedor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   16
         Top             =   4200
         Width           =   3375
      End
      Begin VB.ComboBox cmbFamiliaGasto 
         Height          =   315
         ItemData        =   "frmIntroducirGastosInternos.frx":147E
         Left            =   1680
         List            =   "frmIntroducirGastosInternos.frx":1480
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtFactura 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   17
         Top             =   4640
         Width           =   1815
      End
      Begin VB.ComboBox cmbEntrada 
         Height          =   315
         ItemData        =   "frmIntroducirGastosInternos.frx":1482
         Left            =   3840
         List            =   "frmIntroducirGastosInternos.frx":148C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1280
         Width           =   1215
      End
      Begin VB.TextBox txtNumPersonas 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   14
         Top             =   3435
         Width           =   615
      End
      Begin VB.TextBox txtNumTickets 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   13
         Top             =   3075
         Width           =   615
      End
      Begin VB.TextBox txtNumKms 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         MaxLength       =   5
         TabIndex        =   12
         Top             =   2715
         Width           =   615
      End
      Begin VB.TextBox txtNumNoches 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   11
         Top             =   2355
         Width           =   615
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   15
         Top             =   3795
         Width           =   1095
      End
      Begin VB.TextBox txtConcepto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3435
         Width           =   3375
      End
      Begin VB.TextBox txtItinerario 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2715
         Width           =   3375
      End
      Begin VB.TextBox txtDonde 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2355
         Width           =   3375
      End
      Begin VB.CommandButton cmdOtroEmpleado 
         Caption         =   "Nuevo empleado"
         Height          =   615
         Left            =   5760
         Picture         =   "frmIntroducirGastosInternos.frx":14A1
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdGuardarGasto 
         Caption         =   "Guardar"
         Height          =   615
         Left            =   5760
         Picture         =   "frmIntroducirGastosInternos.frx":1D6B
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox cmbTiposGasto 
         Height          =   315
         ItemData        =   "frmIntroducirGastosInternos.frx":2635
         Left            =   1680
         List            =   "frmIntroducirGastosInternos.frx":2637
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
      Begin VB.ComboBox cmbTipoComida 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmIntroducirGastosInternos.frx":2639
         Left            =   1680
         List            =   "frmIntroducirGastosInternos.frx":263B
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3075
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1275
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   140574723
         CurrentDate     =   38755
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Personas"
         Height          =   195
         Index           =   16
         Left            =   360
         TabIndex        =   40
         Top             =   5085
         Width           =   660
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   17
         Left            =   360
         TabIndex        =   39
         Top             =   4245
         Width           =   750
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de gasto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   15
         Left            =   360
         TabIndex        =   38
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   37
         Top             =   4680
         Width           =   555
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Entrada"
         Height          =   195
         Index           =   13
         Left            =   3120
         TabIndex        =   36
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Index           =   12
         Left            =   360
         TabIndex        =   35
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   34
         Top             =   3840
         Width           =   570
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nº personas"
         Height          =   195
         Index           =   10
         Left            =   5520
         TabIndex        =   33
         Top             =   3480
         Width           =   885
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   32
         Top             =   3480
         Width           =   690
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo comida"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   31
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nº tickets"
         Height          =   195
         Index           =   7
         Left            =   5520
         TabIndex        =   30
         Top             =   3120
         Width           =   690
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Kilómetros"
         Height          =   195
         Index           =   6
         Left            =   5520
         TabIndex        =   29
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nº noches"
         Height          =   195
         Index           =   5
         Left            =   5520
         TabIndex        =   28
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Itinerario"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   27
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Donde"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   26
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Concepto Gasto"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   23
         Top             =   405
         Width           =   1155
      End
   End
   Begin VB.ComboBox cmbEmpleado 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin VB.ComboBox cmbDepartamento 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   315
      Width           =   4215
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Empleado"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   1005
      Width           =   690
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Departamento"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   360
      Width           =   1035
   End
End
Attribute VB_Name = "frmIntroducirGastosInternos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mcltTiposGasto As Collection
Dim dImporteKM As Double

Private Sub cmbDepartamento_Click()
    
    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    Dim iCont As Integer
        
    On Error GoTo Error_Rutina
    
    For iCont = 0 To cmbEmpleado.ListCount - 1
        cmbEmpleado.RemoveItem 0
    Next iCont
    cmbEmpleado.Clear
    If cmbDepartamento.Text = "" Then Exit Sub
    
    sSql = "SELECT * FROM Usuarios_Personal " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " AND Departamento = " & Trim(Mid(cmbDepartamento.Text, 1, InStr(1, cmbDepartamento.Text, "-") - 1)) & _
           " AND FechaBaja IS NULL " & _
           "ORDER BY Nombre"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Sub
    End If
    While Not afoRegistro.EOF
        cmbEmpleado.AddItem afoRegistro("Codigo") & " - " & Trim(afoRegistro("Nombre")) & " (" & Trim(afoRegistro("Jerarquia")) & ")"
        afoRegistro.MoveNext
    Wend
    Set afoRegistro = Nothing
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmbEntrada_Click()
    
'    txtFactura = ""
'    txtFactura.Enabled = False
'    If cmbEntrada = "Facturado" Then txtFactura.Enabled = True
    Tabular
    
End Sub

Private Sub cmbEntrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub cmbTipoComida_Click()
    Tabular
End Sub

Private Sub cmbTipoComida_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub cmbTiposGasto_Click()

    VaciarCampos
    
    Select Case cmbTiposGasto.Text
    Case "Gasolina"
        txtItinerario.Enabled = True
        txtImporte.Enabled = True
    Case "Kilómetros"
        txtImporte = dImporteKM
        txtItinerario.Enabled = True
        txtNumKms.Enabled = True
    Case "Comidas-Invitaciones"
        txtDonde.Enabled = True
        cmbTipoComida.Enabled = True
        txtConcepto.Enabled = True
        txtNumPersonas.Enabled = True
        txtImporte.Enabled = True
        txtPersonas.Enabled = True
    Case "Salones"
        txtDonde.Enabled = True
        txtImporte.Enabled = True
    Case "Hoteles"
        txtDonde.Enabled = True
        txtNumNoches.Enabled = True
        txtNumPersonas.Enabled = True
        txtImporte.Enabled = True
    Case "Teléfono"
        txtImporte.Enabled = True
    Case "Billetes"
        txtItinerario.Enabled = True
        txtImporte.Enabled = True
    Case "Taxis"
        txtItinerario.Enabled = True
        txtImporte.Enabled = True
    Case "Peajes"
        txtItinerario.Enabled = True
        txtNumTickets.Enabled = True
        txtImporte.Enabled = True
    Case "Parking"
        txtDonde.Enabled = True
        txtImporte.Enabled = True
    Case "Otros Gastos"
        txtConcepto.Enabled = True
        txtImporte.Enabled = True
    Case Else
        Exit Sub
    End Select
    txtFactura.Enabled = True
    txtProveedor.Enabled = True
    Tabular

End Sub

Private Sub cmdContinuar_Click()

    If cmbDepartamento.Text = "" Then
        cmbDepartamento.SetFocus
        Exit Sub
    End If
    If cmbEmpleado.Text = "" Then
        cmbEmpleado.SetFocus
        Exit Sub
    End If
    cmbDepartamento.Enabled = False
    cmbEmpleado.Enabled = False
    cmdContinuar.Enabled = False
    fraDatosGasto.Enabled = True
    cmbTiposGasto.SetFocus

End Sub

Private Sub cmdGuardarGasto_Click()

    Dim sSql As String
    
    Dim iNumero As Long
    Dim iNumero2 As Integer
    Dim sCadena As String

    On Error GoTo Error_Rutina
    
    If cmbTiposGasto.Text = "" Then
        cmbTiposGasto.SetFocus
        Exit Sub
    End If
    If cmbFamiliaGasto.Text = "" Then
        cmbFamiliaGasto.SetFocus
        Exit Sub
    End If
    If cmbEntrada.Text = "" Then
        cmbEntrada.SetFocus
        Exit Sub
    End If
    
    txtDonde = Trim(Replace(txtDonde, "'", " "))
    txtItinerario = Trim(Replace(txtItinerario, "'", " "))
    txtConcepto = Trim(Replace(txtConcepto, "'", " "))
    txtNumNoches = Trim(txtNumNoches)
    txtNumKms = Trim(txtNumKms)
    txtNumTickets = Trim(txtNumTickets)
    txtNumPersonas = Trim(txtNumPersonas)
    txtFactura = Trim(Replace(txtFactura, "'", " "))
    txtImporte = Trim(txtImporte)
    If txtImporte = "" Then
        txtImporte.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtImporte) Then
        txtImporte = ""
        txtImporte.SetFocus
        Exit Sub
    End If
    
    iNumero = 0
    iNumero2 = 0
    sCadena = ""
    Select Case cmbTiposGasto.Text
    Case "Gasolina"
    
        If txtItinerario = "" Then
            txtItinerario.SetFocus
            Exit Sub
        End If
        sCadena = txtItinerario
        
    Case "Kilómetros"
    
        If txtItinerario = "" Then
            txtItinerario.SetFocus
            Exit Sub
        End If
        If txtNumKms = "" Then
            txtNumKms.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtNumKms) Then
            txtNumKms = ""
            txtNumKms.SetFocus
            Exit Sub
        End If
        
        iNumero = txtNumKms
        sCadena = txtItinerario
        
    Case "Comidas-Invitaciones"
    
        If txtDonde = "" Then
            txtDonde.SetFocus
            Exit Sub
        End If
        If cmbTipoComida.Text = "" Then
            cmbTipoComida.SetFocus
            Exit Sub
        End If
        If txtNumPersonas = "" Then
            txtNumPersonas.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtNumPersonas) Then
            txtNumPersonas = ""
            txtNumPersonas.SetFocus
            Exit Sub
        End If
        
        iNumero = txtNumPersonas
        sCadena = "(" & cmbTipoComida & ") " & txtDonde
        If txtConcepto <> "" Then sCadena = sCadena & " / " & txtConcepto
        
        txtPersonas = Trim(txtPersonas)
        If iNumero > 1 Then
            If txtPersonas = "" Then
                txtPersonas.SetFocus
                Exit Sub
            End If
        End If
        
    Case "Salones"
    
        If txtDonde = "" Then
            txtDonde.SetFocus
            Exit Sub
        End If
        
        sCadena = txtDonde
        
    Case "Hoteles"
    
        If txtDonde = "" Then
            txtDonde.SetFocus
            Exit Sub
        End If
        If txtNumNoches = "" Then
            txtNumNoches.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtNumNoches) Then
            txtNumNoches = ""
            txtNumNoches.SetFocus
            Exit Sub
        End If
        If txtNumPersonas = "" Then
            txtNumPersonas.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtNumPersonas) Then
            txtNumPersonas = ""
            txtNumPersonas.SetFocus
            Exit Sub
        End If
    
        sCadena = txtDonde
        iNumero = txtNumNoches
        iNumero2 = txtNumPersonas
        
    Case "Teléfono"
    
    Case "Billetes", "Taxis"
        
        If txtItinerario = "" Then
            txtItinerario.SetFocus
            Exit Sub
        End If
        
        sCadena = txtItinerario
        
    Case "Peajes"
    
        If txtItinerario = "" Then
            txtItinerario.SetFocus
            Exit Sub
        End If
        If txtNumTickets = "" Then
            txtNumTickets.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtNumTickets) Then
            txtNumTickets = ""
            txtNumTickets.SetFocus
            Exit Sub
        End If
        
        sCadena = txtItinerario
        iNumero = txtNumTickets
        
    Case "Parking"
    
        If txtDonde = "" Then
            txtDonde.SetFocus
            Exit Sub
        End If
        
        sCadena = txtDonde
        
    Case "Otros Gastos"
        
        If txtConcepto = "" Then
            txtConcepto.SetFocus
            Exit Sub
        End If
        
        sCadena = txtConcepto
        
    Case Else
        Exit Sub
    End Select
    If cmbEntrada.Text = "Facturado" And txtProveedor = "" Then
        txtProveedor.SetFocus
        Exit Sub
    End If
    If cmbEntrada.Text = "Facturado" And txtFactura = "" Then
        txtFactura.SetFocus
        Exit Sub
    End If
    If cmbEntrada.Text <> "Facturado" Then
        txtProveedor = ""
        txtFactura = ""
    End If
    
    sCadena = Replace(sCadena, "'", " ")
    sCadena = Replace(sCadena, """", " ")
    
    sSql = "INSERT INTO Gastos_Internos(CodPais,Codigo,Departamento,Jerarquia,TipoGasto,Fecha,Importe,Numero,Numero2,Donde,Tipo_Com_Inv,VISA,Validado,Validacion,Pagado,Moneda,Factura,FamiliaGasto,Proveedor,Personas) " & _
           " VALUES(" & CodPaisTrabajo & "," & _
           Trim(Mid(cmbEmpleado.Text, 1, InStr(1, cmbEmpleado.Text, "-") - 1)) & "," & _
           Trim(Mid(cmbDepartamento.Text, 1, InStr(1, cmbDepartamento.Text, "-") - 1)) & ",'" & _
           Mid(cmbEmpleado.Text, Len(cmbEmpleado.Text) - 2, 2) & "'," & _
           mcltTiposGasto.Item(cmbTiposGasto.Text) & ",'" & _
           Format(dtpFecha, "ddMMyyyy") & "'," & txtImporte & "," & _
           iNumero & "," & iNumero2 & ",'" & sCadena & "','" & _
           Mid(cmbTipoComida.Text, 1, 1) & "','"
    If cmbEntrada.Text = "VISA" Then
        sSql = sSql & "S"
    Else
        sSql = sSql & "F"
    End If
    sSql = sSql & "','S','','S','EUR','" & txtFactura & "','" & Mid(cmbFamiliaGasto.Text, 1, 1) & "','" & txtProveedor & "','" & txtPersonas & "')"
    
    If Not gBdGALILEO.bExecute(sSql) Then
        MsgBox "Error al intentar guardar la información", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    'VaciarCampos
    cmbTiposGasto_Click
    cmbTiposGasto.SetFocus
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdOtroEmpleado_Click()

    VaciarCampos
    dtpFecha = Now
    fraDatosGasto.Enabled = False
    cmbDepartamento.Enabled = True
    cmbEmpleado.Enabled = True
    cmdContinuar.Enabled = True

End Sub

Private Sub cmdVolver_Click()
    Set mcltTiposGasto = Nothing
    Unload Me
End Sub

Private Sub Form_Load()

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset
    
    On Error GoTo Error_Rutina
    
    dtpFecha = Now
'    dtpFecha.MaxDate = Now
    
    Set mcltTiposGasto = New Collection
    mcltTiposGasto.Add 1, "Gasolina"
    cmbTiposGasto.AddItem "Gasolina"
    
    Select Case CodPaisTrabajo
    Case 34, 39, 51, 85, 56, 57, 12
    
        mcltTiposGasto.Add 2, "Kilómetros"
        cmbTiposGasto.AddItem "Kilómetros"
        Select Case CodPaisTrabajo
        Case 34
            dImporteKM = 0.03
        Case 39
            dImporteKM = 0.08
        Case 51
            dImporteKM = 0.1
        Case 85
            dImporteKM = 2
        Case 56
            '06-06-17 Benitez indica que el importe del km se va a pagar a 30 pesos en Chile
            'dImporteKM = 15
            dImporteKM = 30
        Case 57
            dImporteKM = 75
        Case 12
            dImporteKM = 0.33
        End Select
        
    End Select
    
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
    
    cmbTipoComida.AddItem "Interno"
    cmbTipoComida.AddItem "Negocios"
    
    cmbFamiliaGasto.AddItem "Ordinario"
    cmbFamiliaGasto.AddItem "Formación"
    cmbFamiliaGasto.AddItem "Presentación"
    
    sSql = "SELECT Codigo,Descripcion FROM DepartamentosCL " & _
           "WHERE CodPais = " & CodPaisTrabajo & _
           " ORDER BY Descripcion"
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

Private Sub txtConcepto_GotFocus()
    TextGotFocus txtConcepto
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtConcepto_LostFocus()
    TextLostFocus txtConcepto
End Sub

Private Sub txtDonde_GotFocus()
    TextGotFocus txtDonde
End Sub

Private Sub txtDonde_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtDonde_LostFocus()
    TextLostFocus txtDonde
End Sub

Private Sub txtFactura_GotFocus()
    TextGotFocus txtFactura
End Sub

Private Sub txtFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtFactura_LostFocus()
    TextLostFocus txtFactura
End Sub

Private Sub txtImporte_GotFocus()
    TextGotFocus txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    EntraNumero KeyAscii, txtImporte, 7, 2
End Sub

Private Sub txtImporte_LostFocus()
    TextLostFocus txtImporte
End Sub

Private Sub txtItinerario_GotFocus()
    TextGotFocus txtItinerario
End Sub

Private Sub txtItinerario_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtItinerario_LostFocus()
    TextLostFocus txtItinerario
End Sub

Private Sub txtNumKms_GotFocus()
    TextGotFocus txtNumKms
End Sub

Private Sub txtNumKms_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    EntraNumero KeyAscii, txtNumKms, 3
End Sub

Private Sub txtNumKms_LostFocus()
    TextLostFocus txtNumKms
End Sub

Private Sub txtNumNoches_GotFocus()
    TextGotFocus txtNumNoches
End Sub

Private Sub txtNumNoches_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    EntraNumero KeyAscii, txtNumNoches, 2
End Sub

Private Sub txtNumNoches_LostFocus()
    TextLostFocus txtNumNoches
End Sub

Private Sub txtNumPersonas_GotFocus()
    TextGotFocus txtNumPersonas
End Sub

Private Sub txtNumPersonas_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    EntraNumero KeyAscii, txtNumPersonas, 2
End Sub

Private Sub txtNumPersonas_LostFocus()
    TextLostFocus txtNumPersonas
End Sub

Private Sub txtNumTickets_GotFocus()
    TextGotFocus txtNumTickets
End Sub

Private Sub txtNumTickets_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then
        Tabular
        Exit Sub
    End If
    EntraNumero KeyAscii, txtNumTickets, 2
End Sub

Private Sub txtNumTickets_LostFocus()
    TextLostFocus txtNumTickets
End Sub

Private Sub VaciarCampos()
    
    cmbEntrada.ListIndex = -1
    txtDonde = ""
    txtDonde.Enabled = False
    txtItinerario = ""
    txtItinerario.Enabled = False
    cmbTipoComida.ListIndex = -1
    cmbTipoComida.Enabled = False
    txtConcepto = ""
    txtConcepto.Enabled = False
    txtNumNoches = ""
    txtNumNoches.Enabled = False
    txtNumKms = ""
    txtNumKms.Enabled = False
    txtNumTickets = ""
    txtNumTickets.Enabled = False
    txtNumPersonas = ""
    txtNumPersonas.Enabled = False
    txtImporte = ""
    txtImporte.Enabled = False
    txtFactura = ""
    txtFactura.Enabled = False
    txtProveedor = ""
    txtProveedor.Enabled = False
    txtPersonas = ""
    txtPersonas.Enabled = False
    
End Sub

Private Sub txtProveedor_GotFocus()
    TextGotFocus txtProveedor
End Sub

Private Sub txtProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtProveedor_LostFocus()
    TextLostFocus txtProveedor
End Sub

Private Sub txtPersonas_GotFocus()
    TextGotFocus txtPersonas
End Sub

Private Sub txtPersonas_KeyPress(KeyAscii As Integer)
    If KeyAscii = ENTER Then Tabular
End Sub

Private Sub txtPersonas_LostFocus()
    TextLostFocus txtPersonas
End Sub

