VERSION 5.00
Begin VB.Form frmCargarFicheroORBIS 
   BorderStyle     =   0  'None
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9225
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
   ScaleHeight     =   4290
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCargar 
      Caption         =   "C A R G A R    D A T O S"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5040
      Picture         =   "frmCargarFicheroORBIS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2685
      Width           =   3615
   End
   Begin VB.CommandButton cmdComprobarFichero 
      Caption         =   "COMPROBAR FICHERO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      Picture         =   "frmCargarFicheroORBIS.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2685
      Width           =   3615
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Caption         =   "Cargando fila"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   3165
      Visible         =   0   'False
      Width           =   8985
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Caption         =   "Los datos de las facturas LayTours a cargar deben estar en el fichero 'C:\GeneracionCSV\LayTours.csv'."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1050
      Width           =   8985
   End
   Begin VB.Image imgVolver 
      Height          =   360
      Left            =   8640
      Picture         =   "frmCargarFicheroORBIS.frx":1794
      ToolTipText     =   "Volver"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgIcono 
      Height          =   480
      Left            =   120
      Picture         =   "frmCargarFicheroORBIS.frx":1EFE
      Top             =   35
      Width           =   480
   End
   Begin VB.Label lblTexto 
      BackColor       =   &H00E2C07A&
      Caption         =   "Carga fichero Lay Tours"
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
      Height          =   495
      Index           =   8
      Left            =   720
      TabIndex        =   5
      Top             =   75
      Width           =   4455
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   5640
      Picture         =   "frmCargarFicheroORBIS.frx":27C8
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Caption         =   $"frmCargarFicheroORBIS.frx":2AD2
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1845
      Width           =   8385
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Caption         =   "No debe tener cabecera y seguir el siguiente formato:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1485
      Width           =   8985
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
      Width           =   9225
   End
End
Attribute VB_Name = "frmCargarFicheroORBIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCargar_Click()
    
    Dim sSql As String
    Dim afoRegDatos As ADODB.Recordset
    Dim afoRegAux As ADODB.Recordset
    
    Dim sFichero As String
    Dim iFichero As Integer
    Dim iFila As Long
    
    Dim sLinea As String
    Dim sFactura As String
    Dim sConcepto As String
    Dim sImporte As String
    Dim sUsuarioEpsilon As String
    Dim sNumeroAlbaran As String
    Dim sDetalle As String
    Dim sFecha As String
    Dim sNumeroDias As String
    Dim sDestino As String
    
    Dim bTransaccionAbierta As Boolean
    
    Dim sidTipoGasto As String
    Dim sTipoGasto As String
    Dim sDondeGasto As String
    Dim dImporteGasto As Double
    Dim iNumero As Integer
    Dim iNumero2 As Integer
    
    Dim iCodPais As Integer
    
    Dim iNumPuntoycoma As Integer
    Dim sLineaAux As String
    
    On Error GoTo Error_Rutina
    
    bTransaccionAbierta = False
    iFila = 0
    
    sFichero = Dir(gsRutaEquipo & "GeneracionCSV\LayTours.csv")
    If sFichero = "" Then
        MsgBox "No existe el fichero '" & gsRutaEquipo & "GeneracionCSV\LayTours.csv'", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    
    iFichero = FreeFile
    Open gsRutaEquipo & "GeneracionCSV\LayTours.csv" For Input As #iFichero
    bTransaccionAbierta = True
    gBdGALILEO.BeginTrans True
    Me.MousePointer = HOURGLASS
    cmdCargar.Visible = False
    cmdComprobarFichero.Visible = False
    lblTexto(3).Visible = True
    
    Do While Not EOF(iFichero)
    
        iFila = iFila + 1
        lblTexto(3).Caption = "Cargando la fila " & iFila
        Refresh

        Line Input #iFichero, sLinea
        
        sLinea = Trim(sLinea)
        If sLinea = "" Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Línea vacía en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLineaAux = sLinea
        iNumPuntoycoma = 0
        While InStr(1, sLineaAux, ";")
            iNumPuntoycoma = iNumPuntoycoma + 1
            sLineaAux = Mid(sLineaAux, InStr(1, sLineaAux, ";") + 1)
        Wend
        If iNumPuntoycoma <> 8 Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Número de columnas incorrectas en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sFactura = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If sFactura = "" Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Invoice number' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sConcepto = LCase(Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)))
        Select Case sConcepto
            Case "avion", "gasto gestion", "hotel", "mayoristas internac.", "mayorista internac.", "mayoristas nacional", "mayorista nacional", "no usar  avion / tren / barco", "otros", "rent a car", "servicios de tierra", "servicio de tierra", "tren"
                sConcepto = sConcepto
            Case Else
                sConcepto = ""
        End Select
        If sConcepto = "" Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Accounting concept' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sImporte = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If Not IsNumeric(sImporte) Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Amount' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sUsuarioEpsilon = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If Not IsNumeric(sUsuarioEpsilon) Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'User ID' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sNumeroAlbaran = LimpiarCadena(Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)))
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sDetalle = LimpiarCadena(Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)))
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sFecha = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If IsNumeric(sFecha) Then
            If Len(sFecha) = 7 Then sFecha = "0" & sFecha
            If Len(sFecha) = 8 Then
                Select Case CInt(Mid(sFecha, 3, 2))
                    Case 2
                        If (CInt(Mid(sFecha, 5, 4)) Mod 4) Then
                            If CInt(Mid(sFecha, 1, 2)) < 1 Or CInt(Mid(sFecha, 1, 2)) > 29 Then sFecha = ""
                        Else
                            If CInt(Mid(sFecha, 1, 2)) < 1 Or CInt(Mid(sFecha, 1, 2)) > 28 Then sFecha = ""
                        End If
                    Case 1, 3, 5, 7, 8, 10, 12
                        If CInt(Mid(sFecha, 1, 2)) < 1 Or CInt(Mid(sFecha, 1, 2)) > 31 Then sFecha = ""
                    Case 4, 6, 9, 11
                        If CInt(Mid(sFecha, 1, 2)) < 1 Or CInt(Mid(sFecha, 1, 2)) > 30 Then sFecha = ""
                    Case Else
                        sFecha = ""
                End Select
            Else
                sFecha = ""
            End If
        Else
            sFecha = ""
        End If
        If sFecha = "" Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Fecha inicio servicio' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sNumeroDias = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If Not IsNumeric(sNumeroDias) Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Num de dias' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
                
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        If InStr(1, sLinea, ";") > 0 Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Más campos de los esperados en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        sDestino = LimpiarCadena(sLinea)
        
        iCodPais = CodPaisTrabajo
        If Not bDameDatosUsuarioEpsilon(sUsuarioEpsilon, iCodPais, afoRegDatos) Then
            gBdGALILEO.RollbackTrans
            Set afoRegDatos = Nothing
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Se ha encontrado un usuario de Epsilon sin asignar en el sistema de gastos en la fila " & iFila & ".", vbApplicationModal + vbInformation, "Información"
            Exit Sub
        End If
        
        Select Case iCodPais
        Case 34, 39, 51, 52, 56, 12, 16, 18
            sidTipoGasto = "I"
        Case Else
            
            sidTipoGasto = "I"
            
            sSql = "SELECT id FROM Gastos_Empresas WHERE CodPais = " & iCodPais
            If gBdGALILEO.bRecordset(sSql, afoRegAux) Then sidTipoGasto = afoRegAux("id")
            Set afoRegAux = Nothing
            
        End Select
        
        iNumero = 0
        iNumero2 = 0
        sTipoGasto = ""
        Select Case sConcepto
            Case "avion", "no usar  avion / tren / barco", "tren"
            'Billete
                sSql = "Billetes"
                sDondeGasto = sNumeroAlbaran & "/" & sDestino
            Case "hotel"
            'Hotel
                sSql = "Hoteles"
                iNumero = CInt(sNumeroDias)
                iNumero2 = 1
                sDondeGasto = sNumeroAlbaran & "/" & sDetalle
            Case Else '"gasto gestion", "mayorista internac.", "mayoristas nacional", "mayorista nacional", "mayoristas internac.", "otros", "rent a car", "servicio de tierra", "servicios de tierra"
            'Otros gastos
                sSql = "Otros gastos"
                sDondeGasto = sNumeroAlbaran & "/" & sDetalle
        End Select
        sSql = "SELECT TipoGasto FROM TiposGasto WHERE Tipo = '" & sidTipoGasto & "' AND NombreGasto = '" & sSql & "'"
        If gBdGALILEO.bRecordset(sSql, afoRegAux) Then sTipoGasto = afoRegAux("TipoGasto")
        Set afoRegAux = Nothing
        If sTipoGasto = "" Then
            gBdGALILEO.RollbackTrans
            Set afoRegDatos = Nothing
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "No se ha encontrado el tipo de gasto en la fila " & iFila & ".", vbApplicationModal + vbInformation, "Información"
            Exit Sub
        End If
        
        sDondeGasto = Replace(Mid(sDondeGasto, 1, 50), "'", " ")
            
        sSql = "INSERT INTO Gastos_Internos(CodPais,Codigo,Departamento,Jerarquia,TipoGasto,Fecha,Importe,Numero,Numero2,Donde,Tipo_Com_Inv,VISA,Validado,Validacion,Pagado," & _
               "Moneda,Factura,FamiliaGasto,Fecha_Insert) " & _
               " VALUES(" & iCodPais & "," & afoRegDatos("Codigo") & "," & afoRegDatos("Departamento") & ",'" & afoRegDatos("Jerarquia") & "'," & _
               sTipoGasto & ",'" & sFecha & "'," & sImporte & "," & iNumero & "," & iNumero2 & ",'" & sDondeGasto & "','','F','S','Carga Orbis','S','EUR','" & sFactura & "','O',CURRENT YEAR TO MINUTE)"
        If Not gBdGALILEO.bExecute(sSql) Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Set afoRegDatos = Nothing
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Error al intentar guardar datos de la fila " & iFila & " en Gastos_Internos.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        Set afoRegDatos = Nothing
        
    Loop
    gBdGALILEO.CommitTrans
    bTransaccionAbierta = False
    Close #iFichero
    
    Me.MousePointer = DEFAULT
    lblTexto(3).Caption = ""
    cmdCargar.Visible = True
    cmdComprobarFichero.Visible = True
    lblTexto(3).Visible = False
    Refresh
    MsgBox "Carga realizada correctamente.", vbApplicationModal + vbInformation, "Información"
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    Me.MousePointer = DEFAULT
    If bTransaccionAbierta Then
        gBdGALILEO.RollbackTrans
        bTransaccionAbierta = False
    End If
    Close
    lblTexto(3).Caption = ""
    cmdCargar.Visible = True
    cmdComprobarFichero.Visible = True
    lblTexto(3).Visible = False
    Refresh
    sLinea = Err.Number & " - " & Err.Description
    If iFila <> 0 Then sLinea = "Error en la línea " & iFila & ", no se ha guardado ningún dato. " & sLinea
    MsgBox sLinea, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdComprobarFichero_Click()

    Dim sSql As String
    Dim afoRegDatos As ADODB.Recordset
    
    Dim sFichero As String
    Dim iFichero As Integer
    Dim iFicheroSalida As Integer
    Dim iFila As Long
    
    Dim sLinea As String
    Dim sFactura As String
    Dim sConcepto As String
    Dim sImporte As String
    Dim sUsuarioEpsilon As String
    Dim sNumeroAlbaran As String
    Dim sDetalle As String
    Dim sFecha As String
    Dim sNumeroDias As String
    Dim sDestino As String
    
    Dim bUsuarioEpsilonNoEncontrado As Boolean
    
    Dim iNumPuntoycoma As Integer
    Dim sLineaAux As String
    
    On Error GoTo Error_Rutina
    
    iFila = 0
    bUsuarioEpsilonNoEncontrado = False
    
    sFichero = Dir(gsRutaEquipo & "GeneracionCSV\LayTours.csv")
    If sFichero = "" Then
        MsgBox "No existe el fichero '" & gsRutaEquipo & "GeneracionCSV\LayTours.csv'", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    
    iFichero = FreeFile
    Open gsRutaEquipo & "GeneracionCSV\LayTours.csv" For Input As #iFichero
    iFicheroSalida = FreeFile
    Open gsRutaEquipo & "GeneracionCSV\LayTours_Usuarios_Incorrectos.csv" For Output As #iFicheroSalida
    Me.MousePointer = HOURGLASS
    cmdCargar.Visible = False
    cmdComprobarFichero.Visible = False
    lblTexto(3).Visible = True
    lblTexto(3) = ""
    Refresh
    
    Do While Not EOF(iFichero)
        
        iFila = iFila + 1
        lblTexto(3).Caption = "Comprobando la fila " & iFila
        Refresh

        Line Input #iFichero, sLinea
        
        sLinea = Trim(sLinea)
        If sLinea = "" Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Línea vacía en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLineaAux = sLinea
        iNumPuntoycoma = 0
        While InStr(1, sLineaAux, ";")
            iNumPuntoycoma = iNumPuntoycoma + 1
            sLineaAux = Mid(sLineaAux, InStr(1, sLineaAux, ";") + 1)
        Wend
        If iNumPuntoycoma <> 8 Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Número de columnas incorrectas en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sFactura = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If sFactura = "" Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Invoice number' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sConcepto = LCase(Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)))
        Select Case sConcepto
            Case "avion", "gasto gestion", "hotel", "mayorista internac.", "mayoristas internac.", "mayoristas nacional", "mayorista nacional", "no usar  avion / tren / barco", "otros", "rent a car", "servicios de tierra", "servicio de tierra", "tren"
                sConcepto = sConcepto
            Case Else
                sConcepto = ""
        End Select
        If sConcepto = "" Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Accounting concept' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sImporte = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If Not IsNumeric(sImporte) Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Amount' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sUsuarioEpsilon = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If Not IsNumeric(sUsuarioEpsilon) Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'User ID' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sNumeroAlbaran = LimpiarCadena(Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)))
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sDetalle = LimpiarCadena(Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)))
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sFecha = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If IsNumeric(sFecha) Then
            If Len(sFecha) = 7 Then sFecha = "0" & sFecha
            If Len(sFecha) = 8 Then
                Select Case CInt(Mid(sFecha, 3, 2))
                    Case 2
                        If (CInt(Mid(sFecha, 5, 4)) Mod 4) Then
                            If CInt(Mid(sFecha, 1, 2)) < 1 Or CInt(Mid(sFecha, 1, 2)) > 29 Then sFecha = ""
                        Else
                            If CInt(Mid(sFecha, 1, 2)) < 1 Or CInt(Mid(sFecha, 1, 2)) > 28 Then sFecha = ""
                        End If
                    Case 1, 3, 5, 7, 8, 10, 12
                        If CInt(Mid(sFecha, 1, 2)) < 1 Or CInt(Mid(sFecha, 1, 2)) > 31 Then sFecha = ""
                    Case 4, 6, 9, 11
                        If CInt(Mid(sFecha, 1, 2)) < 1 Or CInt(Mid(sFecha, 1, 2)) > 30 Then sFecha = ""
                    Case Else
                        sFecha = ""
                End Select
            Else
                sFecha = ""
            End If
        Else
            sFecha = ""
        End If
        If sFecha = "" Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Fecha inicio servicio' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sNumeroDias = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If Not IsNumeric(sNumeroDias) Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Num de dias' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
                
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        If InStr(1, sLinea, ";") > 0 Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarFichero.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Más campos de los esperados en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        sDestino = LimpiarCadena(sLinea)
        
        If Not bDameDatosUsuarioEpsilon(sUsuarioEpsilon, CodPaisTrabajo, afoRegDatos) Then
            Print #iFicheroSalida, "Usuario Epsilon no asignado: " & sUsuarioEpsilon
            bUsuarioEpsilonNoEncontrado = True
        End If
        Set afoRegDatos = Nothing
        
    Loop
    Close #iFichero
    Close #iFicheroSalida
    
    If bUsuarioEpsilonNoEncontrado Then
        Me.MousePointer = DEFAULT
        lblTexto(3).Caption = ""
        cmdCargar.Visible = True
        cmdComprobarFichero.Visible = True
        lblTexto(3).Visible = False
        Refresh
        MsgBox "Se han encontrado usuarios de Epsilon sin asignar en el sistema de gastos, revise el fichero '" & gsRutaEquipo & "GeneracionCSV\LayTours_Usuarios_Incorrectos.csv'.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    
    Me.MousePointer = DEFAULT
    lblTexto(3).Caption = ""
    cmdCargar.Visible = True
    cmdComprobarFichero.Visible = True
    lblTexto(3).Visible = False
    Refresh
    MsgBox "Comprobación realizada correctamente.", vbApplicationModal + vbInformation, "Información"
    
Exit_Rutina:
    Exit Sub
Error_Rutina:
    Me.MousePointer = DEFAULT
    Close
    lblTexto(3).Caption = ""
    cmdCargar.Visible = True
    cmdComprobarFichero.Visible = True
    lblTexto(3).Visible = False
    Refresh
    sLinea = Err.Number & " - " & Err.Description
    If iFila <> 0 Then sLinea = "Error en la línea " & iFila & ". " & sLinea
    MsgBox sLinea, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub imgVolver_Click()
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgVolver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub
