VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConciliarProveedores 
   BorderStyle     =   0  'None
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19425
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
   ScaleHeight     =   7365
   ScaleWidth      =   19425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConciliar 
      Caption         =   "Asignar"
      Height          =   855
      Left            =   9720
      Picture         =   "frmConciliarProveedores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid mfgParrillaGastos 
      Height          =   5535
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9763
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid mfgParrillaProveedores 
      Height          =   5535
      Left            =   11280
      TabIndex        =   3
      Top             =   1440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9763
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Proveedores / Vendors"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   11280
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Gastos sin proveedor/vendor conciliado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4365
   End
   Begin VB.Shape shRecuadroForm 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   18720
      Picture         =   "frmConciliarProveedores.frx":08CA
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Conciliar gastos / proveedores"
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
      Width           =   5475
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   14400
      Picture         =   "frmConciliarProveedores.frx":0DD6
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmConciliarProveedores.frx":10E0
      Top             =   60
      Width           =   480
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
      Width           =   19425
   End
End
Attribute VB_Name = "frmConciliarProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miCodPais As Integer
Dim mafoRegistro As ADODB.Recordset

Dim moGastos As New CFlexGrid
Dim moProveedores As New CFlexGrid

Dim miFilaProveedores As Integer

Public Sub Create(iCodPais As Integer)
    miCodPais = iCodPais
End Sub

Private Sub cmdConciliar_Click()

    Dim iFila As Integer
    Dim idProveedor As String
    Dim bEnc As Boolean
    Dim sFilasEliminar As String
    Dim sIDFiscal As String
    
    If miFilaProveedores = -1 Then
        MsgBox "Debe seleccionar un proveedor.", vbApplicationModal + vbExclamation, "Información"
        Exit Sub
    End If
    idProveedor = moProveedores.TextoLinCol(miFilaProveedores, 0)
    sIDFiscal = moProveedores.TextoLinCol(miFilaProveedores, 2)
    
    bEnc = False
    sFilasEliminar = ""
    For iFila = 1 To moGastos.Grid.Rows - 1
        If moGastos.TextoLinCol(iFila, 2) = "-->" Then If bAsignarProveedorGasto(CodPaisTrabajo, CLng(moGastos.TextoLinCol(iFila, 0)), idProveedor, sIDFiscal) = False Then bEnc = True
    Next
    
    If bEnc = True Then
        MsgBox "Hay gastos que no han podido ser asignados a un proveedor.", vbApplicationModal + vbExclamation, "Error"
    Else
        MsgBox "Gastos asignados al proveedor correctamente.", vbApplicationModal + vbInformation, "Información"
    End If
    Rellenar_GastosProveedores_SinConciliar
    moProveedores.TextoLinCol(miFilaProveedores, 1) = ""
    moProveedores.SituarEn 0, 1
    miFilaProveedores = -1

End Sub

Private Sub Form_Load()

    If DisplayMonitorCount = 1 Then
        Me.Left = (Screen.Width / 2) - (Me.Width / 2)
        Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Else
        If MDIPrincipal.Left + 120 < Screen.Width Then
        'El MDI está en la pantalla 1, lo centro en la pantalla 1
            Me.Left = (Screen.Width / 2) - (Me.Width / 2)
            Me.Top = (Screen.Height / 2) - (Me.Height / 2)
        Else
            Me.Left = gAnchoPantallaPrincipal + (((gAnchoPantallaVirtual - gAnchoPantallaPrincipal) - Me.Width) \ 2)
            Me.Top = (Screen.Height - Me.Height) \ 2
        End If
    End If

    shRecuadroForm.Width = Me.Width - 5
    shRecuadroForm.Height = Me.Height - 5
    shRecuadroForm.Left = 5
    shRecuadroForm.Top = 5

    moGastos.Create mfgParrillaGastos, 8, 2
    moGastos.Formatear 0, 0, "Contador"
    moGastos.Formatear 1, 0, "Ruta"
    moGastos.Formatear 2, 350, "", 4, 4
    moGastos.Formatear 3, 1000, "Fecha", 4, 4
    moGastos.Formatear 4, 4500, "Usuario gasto / Concepto", 4
    moGastos.Formatear 5, 1250, "Importe", 4, 7
    moGastos.Formatear 6, 750, "Divisa", 4, 4
    moGastos.Formatear 7, 1000, "", 4
    moGastos.SituarEn 0, 1
    
    moProveedores.Create mfgParrillaProveedores, 4, 2
    moProveedores.Formatear 0, 0, "id"
    moProveedores.Formatear 1, 350, "", 4, 4
    moProveedores.Formatear 2, 1500, "ID Fiscal", 4, 7, True
    moProveedores.Formatear 3, 4000, "Proveedor / Vendor", 4, , True
    moProveedores.SituarEn 0, 1
    
    Rellenar_GastosProveedores_SinConciliar
    Rellenar_Proveedores
    
    miFilaProveedores = -1

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub Rellenar_Proveedores()

    If Not bProveedoresGES(mafoRegistro) Then
        Set mafoRegistro = Nothing
        Exit Sub
    End If
    
    While Not mafoRegistro.EOF
        moProveedores.TextoFila moProveedores.LineaLibre, mafoRegistro("id"), "", Trim(mafoRegistro("idFiscal")), Trim(mafoRegistro("Nombre"))
        mafoRegistro.MoveNext
    Wend
    moProveedores.SituarEn 0, 1
    Set mafoRegistro = Nothing

End Sub

Private Sub Rellenar_GastosProveedores_SinConciliar()

    Dim sVerTicket As String
    Dim iFila As Integer

    moGastos.Vaciar
    moGastos.SituarEn 0, 1

    If Not bGastosProveedores_SinConciliar(miCodPais, mafoRegistro) Then
        Set mafoRegistro = Nothing
        Exit Sub
    End If
    
    While Not mafoRegistro.EOF
    
        sVerTicket = ""
        If Trim(mafoRegistro("Ruta_Ticket")) <> "" Then sVerTicket = "Ver ticket"
    
        iFila = moGastos.LineaLibre
        moGastos.TextoFila iFila, mafoRegistro("Contador"), Trim(mafoRegistro("Ruta_Ticket")), "", Format(mafoRegistro("Fecha"), "dd/MM/yy"), _
            "(" & Trim(mafoRegistro("Codigo")) & ") " & Trim(mafoRegistro("Nombre")) & " / " & Trim(mafoRegistro("Donde")), mafoRegistro("Importe_Ori"), _
            mafoRegistro("Moneda"), sVerTicket
        moGastos.SituarEn 4, iFila
        moGastos.Grid.CellFontSize = 7
        
        mafoRegistro.MoveNext
        
    Wend
    Set mafoRegistro = Nothing
    moGastos.SituarEn 0, 1
    
End Sub

Private Sub imgSalir_Click()
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub mfgParrillaGastos_Click()

    Dim sRuta As String
    
    If moGastos.Grid.col = 7 Then
        If moGastos.TextoLinCol(moGastos.Grid.Row, 7) = "Ver ticket" Then
            
            sRuta = LCase(Trim(moGastos.TextoLinCol(moGastos.Grid.Row, 1)))
            sRuta = Replace(sRuta, "\", "/")
            sRuta = Replace(sRuta, "c:/ficherosgv/", "")
            sRuta = "../Marte/" & sRuta
    
            If LCase(Mid(sRuta, Len(sRuta) - 3)) = ".pdf" Then
                sRuta = "personal.cristianlayusers.com\Gastos\VisorTickets.asp?P=" & sRuta
            Else
                sRuta = "personal.cristianlayusers.com\Gastos\VisorTickets.asp?I=" & sRuta
            End If
            frmExplorer.Create sRuta
            frmExplorer.Show MODAL
            
        End If
    End If
    
    If moGastos.bLineaActualVacia Then
        moGastos.SituarEn 0, 1
        Exit Sub
    End If
    
    moGastos.ResaltarLinea
    
End Sub

Private Sub mfgParrillaGastos_DblClick()

    If moGastos.bLineaActualVacia Then
        moGastos.SituarEn 0, 1
        Exit Sub
    End If
    
    If moGastos.TextoLinCol(moGastos.Grid.Row, 2) = "" Then
        moGastos.TextoLinCol(moGastos.Grid.Row, 2) = "-->"
    Else
        moGastos.TextoLinCol(moGastos.Grid.Row, 2) = ""
    End If
    moGastos.ResaltarLinea

End Sub

Private Sub mfgParrillaGastos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.MousePointer = DEFAULT
    
    With mfgParrillaGastos
        If .MouseRow > 0 And (.MouseCol = 4) Then
            .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
            If .ToolTipText = "X" Then .ToolTipText = vbNullString
        Else
            .ToolTipText = vbNullString
        End If
    End With

End Sub

Private Sub mfgParrillaProveedores_Click()
    
    Dim iFila As Integer
    
    If moProveedores.Grid.MouseRow = 0 Then
        If miFilaProveedores <> -1 Then
            moProveedores.TextoLinCol(miFilaProveedores, 1) = ""
            miFilaProveedores = -1
        End If
        moProveedores.OrdenarSiCorresponde
        moProveedores.SituarEn 0, 1
        Exit Sub
    End If
    
    If moProveedores.bLineaActualVacia Then
    
        If miFilaProveedores <> -1 Then
            moProveedores.TextoLinCol(miFilaProveedores, 1) = ""
            miFilaProveedores = -1
        End If
        moProveedores.SituarEn 0, 1
        Exit Sub
        
    End If
    
    iFila = moProveedores.Grid.Row
    If miFilaProveedores <> -1 Then
        moProveedores.TextoLinCol(miFilaProveedores, 1) = ""
        miFilaProveedores = -1
    End If
    moProveedores.SituarEn 0, iFila
    moProveedores.TextoLinCol(iFila, 1) = "-->"
    moProveedores.ResaltarLinea
    miFilaProveedores = iFila
    
End Sub
