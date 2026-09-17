VERSION 5.00
Begin VB.Form frmCargaFicheroSolred2 
   BorderStyle     =   0  'None
   ClientHeight    =   4275
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
   Moveable        =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdComprobarTarjetas 
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
      Picture         =   "frmCargaFicheroSolred2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   3615
   End
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
      Picture         =   "frmCargaFicheroSolred2.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
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
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   8985
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
      TabIndex        =   4
      Top             =   1440
      Width           =   8985
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Caption         =   $"frmCargaFicheroSolred2.frx":2B94
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
      TabIndex        =   3
      Top             =   1800
      Width           =   8385
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   5640
      Picture         =   "frmCargaFicheroSolred2.frx":2C7E
      Top             =   40
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTexto 
      BackColor       =   &H00E2C07A&
      Caption         =   "Carga fichero Solred"
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
      Left            =   960
      TabIndex        =   2
      Top             =   75
      Width           =   4455
   End
   Begin VB.Image imgIcono 
      Height          =   720
      Left            =   120
      Picture         =   "frmCargaFicheroSolred2.frx":2F88
      Top             =   -40
      Width           =   720
   End
   Begin VB.Image imgVolver 
      Height          =   360
      Left            =   8640
      Picture         =   "frmCargaFicheroSolred2.frx":4C52
      ToolTipText     =   "Volver"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Caption         =   "Los datos de las tarjetas SOLRED a cargar deben estar en el fichero 'C:\GeneracionCSV\Solred.csv'."
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
      TabIndex        =   1
      Top             =   1005
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
Attribute VB_Name = "frmCargaFicheroSolred2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCargar_Click()
    
    Dim sSql As String
    Dim afoRegTarjeta As ADODB.Recordset
    Dim afoRegAux As ADODB.Recordset
    
    Dim sFichero As String
    Dim iFichero As Integer
    Dim iFila As Long
    
    Dim iFicheroSalida As Integer
    
    Dim sLinea As String
    Dim sTarjeta As String
    Dim sAnyo As String
    Dim sFactura As String
    Dim sFechaOperacion As String
    Dim sHoraOperacion As String
    Dim sMatricula As String
    Dim sConductor As String
    Dim sPoblacionEst As String
    Dim sKms As String
    Dim sDescProd As String
    Dim sLitros As String
    Dim sIVA As String
    Dim sImporteTotal As String
    Dim sImporteTotal_SinIVA As String
    Dim sPrecioLitro As String
    Dim sInfoAux As String
    Dim lidContador As Long
    
    Dim bTransaccionAbierta As Boolean
    
    Dim sTipoGasto As String
    Dim sDondeGasto As String
    Dim dImporteGasto As Double
    Dim dImporteGasto_SinIVA As Double
    Dim iCampanyaGasto As Integer
    Dim iAnyoGasto As Integer
    Dim iNumero As Integer
    
    Dim iCodPais As Integer
    
    On Error GoTo Error_Rutina
    
    bTransaccionAbierta = False
    iFila = 0
    
    sFichero = Dir(gsRutaEquipo & "GeneracionCSV\Solred.csv")
    If sFichero = "" Then
        MsgBox "No existe el fichero '" & gsRutaEquipo & "GeneracionCSV\Solred.csv'", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    
    iFichero = FreeFile
    Open gsRutaEquipo & "GeneracionCSV\Solred.csv" For Input As #iFichero
    bTransaccionAbierta = True
    gBdGALILEO.BeginTrans True
    Me.MousePointer = HOURGLASS
    cmdCargar.Visible = False
    cmdComprobarTarjetas.Visible = False
    lblTexto(3).Visible = True
    
    Do While Not EOF(iFichero)
        
        iFila = iFila + 1
        lblTexto(3).Caption = "Cargando la fila " & iFila
        Refresh

        Line Input #iFichero, sLinea
        
        sLinea = Trim(sLinea)
        If sLinea = "" Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Línea vacía en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sAnyo = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If Not IsNumeric(sAnyo) Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Año Factura' incorrecto en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sFactura = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If sFactura = "" Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Tarjeta' incorrecto en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sTarjeta = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        While Mid(sTarjeta, 1, 1) = 0
            sTarjeta = Mid(sTarjeta, 2)
        Wend
        If InStr(1, sTarjeta, ".") Then
            sTarjeta = Mid(sTarjeta, 1, InStr(1, sTarjeta, ".") - 1)
        End If
        If InStr(1, sTarjeta, ",") Then
            sTarjeta = Mid(sTarjeta, 1, InStr(1, sTarjeta, ",") - 1)
        End If
        If sTarjeta = "" Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Tarjeta' incorrecto en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sMatricula = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sConductor = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sFechaOperacion = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If InStr(1, sFechaOperacion, "/") = 0 Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo fecha incorrecto en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        sFechaOperacion = Replace(sFechaOperacion, "/", "")
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sHoraOperacion = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If InStr(1, sHoraOperacion, ":") = 0 Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo hora incorrecto en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
                
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sPoblacionEst = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sKms = Replace(Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)), ",", ".")
        If Not IsNumeric(sKms) Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Kilómetros' incorrecto en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sDescProd = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sLitros = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If Not IsNumeric(sLitros) Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Nº Litros' incorrecto en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sIVA = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If Not IsNumeric(sIVA) Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'IVA' incorrecto en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sImporteTotal = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If Not IsNumeric(sImporteTotal) Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Importe Total' incorrecto en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sPrecioLitro = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If Not IsNumeric(sPrecioLitro) Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Precio litro' incorrecto en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        sPrecioLitro = Replace(sPrecioLitro, ",", ".")
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sInfoAux = sLinea
        
        iCodPais = giPAIS_ESPAÑA
        'If Not bDameDatosTarjetaSolred(iCodPais, sTarjeta, afoRegTarjeta) Then
        If Not bDameDatosTarjetaSolred(sTarjeta, afoRegTarjeta) Then
            gBdGALILEO.RollbackTrans
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "No se han encontrado datos de la tarjeta Solred en la fila " & iFila & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        iCodPais = afoRegTarjeta("CodPais")
        
        'Miro sí la tarjeta está asignada a un usuario de la red de ventas para hacer el insert en la tabla de gastos
        'Se han encontrado importes en negativos que son descuentos, estos datos no habría que insertarlos en la tabla de Gastos, sí en la de Gastos Solred
        lidContador = -1
        If afoRegTarjeta("Departamento") = -1 Then
        'Red Comercial

            If CDbl(sImporteTotal) > 0 Then
            
                sTipoGasto = "G"
                sDondeGasto = sPoblacionEst
                dImporteGasto = sImporteTotal
                dImporteGasto_SinIVA = Round(dImporteGasto / (1 + (sIVA / 100)), 2)
                iCampanyaGasto = iDameCampanya(giPAIS_ESPAÑA, afoRegTarjeta("Codigo"), Trim(afoRegTarjeta("Jerarquia")), sFechaOperacion, iAnyoGasto)
                If InStr(1, LCase(sDescProd), "aparcamiento") Or InStr(1, LCase(sDescProd), "aparcamien") Then
                    sTipoGasto = "V"
                    sDondeGasto = "Parking " & sPoblacionEst & " (" & sInfoAux & ")"
                End If
                If InStr(1, LCase(sDescProd), "estradas") Or InStr(1, LCase(sDescProd), "autopista") Or InStr(1, LCase(sDescProd), "teroperable") Or InStr(1, LCase(sDescProd), "autoestrad") Then
                    sTipoGasto = "P"
                    sDondeGasto = sPoblacionEst & " (" & sInfoAux & ")"
                End If
                sDondeGasto = Replace(Mid(sDondeGasto, 1, 50), "'", " ")
                
                'paso el importe del gasto a pesetas
                dImporteGasto_SinIVA = Round(dConvierteMoneda(CStr(dImporteGasto_SinIVA), Euros_A_Local, giPAIS_ESPAÑA), 0)
                sSql = "INSERT INTO Gastos (CodPais,Codigo,Jerarquia,TipoGasto,Fecha,Anyo,Campanya," & _
                       "Importe,Numero,Donde,TipoPago,Identificador,Chequeado,Validado,Validacion,Pagado,Hora) " & _
                       "VALUES (" & giPAIS_ESPAÑA & "," & afoRegTarjeta("Codigo") & ",'" & _
                       Trim(afoRegTarjeta("Jerarquia")) & "','" & sTipoGasto & "','" & sFechaOperacion & "'," & _
                       iAnyoGasto & "," & iCampanyaGasto & "," & dImporteGasto_SinIVA & ",1,'" & _
                       sDondeGasto & "','S','-999999','S','S','Carga SolRed','S'," & Mid(sHoraOperacion, 1, 2) & ")"
                If Not gBdGALILEO.bExecute(sSql) Then
                    gBdGALILEO.RollbackTrans
                    Close #iFichero
                    Me.MousePointer = DEFAULT
                    lblTexto(3).Caption = ""
                    cmdCargar.Visible = True
                    cmdComprobarTarjetas.Visible = True
                    lblTexto(3).Visible = False
                    Refresh
                    MsgBox "Error al intentar guardar datos de la fila " & iFila & " en Gastos_Solred.", vbApplicationModal + vbExclamation, "E R R O R"
                    Exit Sub
                End If
                
                sSql = "SELECT Contador FROM Gastos " & _
                       "WHERE CodPais = " & giPAIS_ESPAÑA & " AND Codigo = " & afoRegTarjeta("Codigo") & " AND Jerarquia = '" & Trim(afoRegTarjeta("Jerarquia")) & "' " & _
                       "AND TipoGasto = '" & sTipoGasto & "' AND Fecha = '" & sFechaOperacion & "' AND Anyo = " & iAnyoGasto & " AND Campanya = " & iCampanyaGasto & _
                       " AND Importe = " & dImporteGasto_SinIVA & " AND Donde = '" & sDondeGasto & "' ORDER BY Contador DESC"
                If gBdGALILEO.bRecordset(sSql, afoRegAux) Then lidContador = afoRegAux("Contador")
                Set afoRegAux = Nothing
            
            End If
            
        Else
        'Interno
        
            If NoNull(afoRegTarjeta("Ins_Gasto_Int"), "T") = "S" And CDbl(sImporteTotal) > 0 Then
            'solo insertamos sí la tarjeta tiene marcado como que inserte gastos
            
                sTipoGasto = "1"    'Gasolina
                iNumero = 0
                sDondeGasto = sPoblacionEst
                dImporteGasto = sImporteTotal
                dImporteGasto_SinIVA = Round(dImporteGasto / (1 + (sIVA / 100)), 2)

                If InStr(1, LCase(sDescProd), "aparcamiento") Or InStr(1, LCase(sDescProd), "aparcamien") Then
                    sTipoGasto = "11"   'Parking
                    sDondeGasto = "Parking " & sPoblacionEst & " (" & sInfoAux & ")"
                    iNumero = 0
                End If
                If InStr(1, LCase(sDescProd), "estradas") Or InStr(1, LCase(sDescProd), "autopista") Or InStr(1, LCase(sDescProd), "teroperable") Or InStr(1, LCase(sDescProd), "autoestrad") Then
                    sTipoGasto = "9"    'Peaje
                    sDondeGasto = sPoblacionEst & " (" & sInfoAux & ")"
                    iNumero = 1
                End If
                If InStr(1, LCase(sDescProd), "lavados") Then
                    sTipoGasto = "10"    'Otros gastos
                    iNumero = 1
                End If
                sDondeGasto = Replace(Mid(sDondeGasto, 1, 50), "'", " ")
                
                sSql = "INSERT INTO Gastos_Internos(CodPais,Codigo,Departamento,Jerarquia,TipoGasto,Fecha,Importe,Numero,Numero2,Donde,Tipo_Com_Inv,VISA,Validado,Validacion,Pagado,Moneda,FamiliaGasto) " & _
                       " VALUES(" & iCodPais & "," & afoRegTarjeta("Codigo") & "," & afoRegTarjeta("Departamento") & ",'" & afoRegTarjeta("Jerarquia") & "'," & _
                       sTipoGasto & ",'" & sFechaOperacion & "'," & dImporteGasto_SinIVA & "," & iNumero & ",0,'" & sDondeGasto & "','','R','S','Carga SolRed','S','EUR','O')"
                If Not gBdGALILEO.bExecute(sSql) Then
                    gBdGALILEO.RollbackTrans
                    Close #iFichero
                    Me.MousePointer = DEFAULT
                    lblTexto(3).Caption = ""
                    cmdCargar.Visible = True
                    cmdComprobarTarjetas.Visible = True
                    lblTexto(3).Visible = False
                    Refresh
                    MsgBox "Error al intentar guardar datos de la fila " & iFila & " en Gastos_Solred.", vbApplicationModal + vbExclamation, "E R R O R"
                    Exit Sub
                End If
                
                sSql = "SELECT Contador FROM Gastos_Internos " & _
                       "WHERE CodPais = " & iCodPais & " AND Codigo = " & afoRegTarjeta("Codigo") & " AND Jerarquia = '" & Trim(afoRegTarjeta("Jerarquia")) & "' " & _
                       " AND Departamento = " & afoRegTarjeta("Departamento") & " AND TipoGasto = " & sTipoGasto & " AND Fecha = '" & sFechaOperacion & _
                       "' AND Importe = " & dImporteGasto_SinIVA & " AND Donde = '" & sDondeGasto & "' ORDER BY Contador DESC"
                If gBdGALILEO.bRecordset(sSql, afoRegAux) Then lidContador = afoRegAux("Contador")
                Set afoRegAux = Nothing
            
            End If
        
        End If
        
        'Insertamos los detalles del movimiento en la tarjeta solred
        If sKms > 9999 Then sKms = "0"
        dImporteGasto = sImporteTotal
        dImporteGasto_SinIVA = Round(dImporteGasto / (1 + (sIVA / 100)), 2)
        sSql = "INSERT INTO Gastos_Solred (Cont_GastoRed,Num_Tarjeta,Anyo_Fact,Factura,Matricula,Conductor,Fec_Operacion,Hor_Operacion,Poblacion,Km,Des_Prod,Num_Litros,IVA,Imp_Total,Imp_Total_SinIVA,Precio_Litro,Info_Aux) " & _
               "VALUES (" & lidContador & ",'" & sTarjeta & "'," & sAnyo & ",'" & sFactura & "','" & sMatricula & "','" & sConductor & "','" & sFechaOperacion & "','" & _
               sHoraOperacion & "','" & sPoblacionEst & "'," & Replace(sKms, ",", ".") & ",'" & sDescProd & "'," & sLitros & "," & sIVA & "," & dImporteGasto & "," & dImporteGasto_SinIVA & "," & sPrecioLitro & ",'" & Replace(sInfoAux, "'", " ") & "')"
        If Not gBdGALILEO.bExecute(sSql) Then
            gBdGALILEO.RollbackTrans
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Error al intentar guardar datos de la fila " & iFila & " en Gastos_Solred.", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
    
    Loop
    gBdGALILEO.CommitTrans
    bTransaccionAbierta = False
    Close #iFichero
    Close #iFicheroSalida
    
    Me.MousePointer = DEFAULT
    lblTexto(3).Caption = ""
    cmdCargar.Visible = True
    cmdComprobarTarjetas.Visible = True
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
    cmdComprobarTarjetas.Visible = True
    lblTexto(3).Visible = False
    Refresh
    sLinea = Err.Number & " - " & Err.Description
    If iFila <> 0 Then sLinea = "Error en la línea " & iFila & ", no se ha guardado ningún dato. " & sLinea
    MsgBox sLinea, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdComprobarTarjetas_Click()

    Dim sSql As String
    Dim afoRegTarjeta As ADODB.Recordset
    
    Dim sFichero As String
    Dim iFichero As Integer
    Dim iFicheroSalida As Integer
    Dim iFila As Long
    
    Dim sLinea As String
    Dim sTarjeta As String
    Dim sAnyo As String
    Dim sFactura As String
    Dim sFechaOperacion As String
    Dim sHoraOperacion As String
    Dim sMatricula As String
    Dim sConductor As String
    Dim sPoblacionEst As String
    Dim sKms As String
    Dim sDescProd As String
    Dim sLitros As String
    Dim sIVA As String
    Dim sImporteTotal As String
    Dim sImporteTotal_SinIVA As String
    Dim sPrecioLitro As String
    Dim sInfoAux As String
    
    Dim bTarjetaNoEncontrada As Boolean
    
    On Error GoTo Error_Rutina
    
    iFila = 0
    bTarjetaNoEncontrada = False
    
    sFichero = Dir(gsRutaEquipo & "GeneracionCSV\Solred.csv")
    If sFichero = "" Then
        MsgBox "No existe el fichero '" & gsRutaEquipo & "GeneracionCSV\Solred.csv'", vbApplicationModal + vbExclamation, "E R R O R"
        Exit Sub
    End If
    
    iFichero = FreeFile
    Open gsRutaEquipo & "GeneracionCSV\Solred.csv" For Input As #iFichero
    iFicheroSalida = FreeFile
    Open gsRutaEquipo & "GeneracionCSV\Sored_Tarjetas_Incorrectas.csv" For Output As #iFicheroSalida
    Me.MousePointer = HOURGLASS
    cmdCargar.Visible = False
    cmdComprobarTarjetas.Visible = False
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
            Close #iFichero
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Línea vacía en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sAnyo = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If Not IsNumeric(sAnyo) Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Año Factura' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sFactura = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If sFactura = "" Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Tarjeta' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sTarjeta = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If InStr(1, sTarjeta, ".") Then
            sTarjeta = Mid(sTarjeta, 1, InStr(1, sTarjeta, ".") - 1)
        End If
        If InStr(1, sTarjeta, ",") Then
            sTarjeta = Mid(sTarjeta, 1, InStr(1, sTarjeta, ",") - 1)
        End If
        While Mid(sTarjeta, 1, 1) = 0
            sTarjeta = Mid(sTarjeta, 2)
        Wend
        If sTarjeta = "" Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Tarjeta' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sMatricula = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sConductor = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sFechaOperacion = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If InStr(1, sFechaOperacion, "/") = 0 Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo fecha incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        sFechaOperacion = Replace(sFechaOperacion, "/", "")
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sHoraOperacion = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If InStr(1, sHoraOperacion, ":") = 0 Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo hora incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
                
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sPoblacionEst = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sKms = Replace(Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)), ",", ".")
        If Not IsNumeric(sKms) Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Kilómetros' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sDescProd = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sLitros = Trim(Mid(sLinea, 1, InStr(1, sLinea, ";") - 1))
        If Not IsNumeric(sLitros) Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Nº Litros' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sIVA = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If Not IsNumeric(sIVA) Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'IVA' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sImporteTotal = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If Not IsNumeric(sImporteTotal) Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Importe Total' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sPrecioLitro = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If Not IsNumeric(sPrecioLitro) Then
            Close
            Me.MousePointer = DEFAULT
            lblTexto(3).Caption = ""
            cmdCargar.Visible = True
            cmdComprobarTarjetas.Visible = True
            lblTexto(3).Visible = False
            Refresh
            MsgBox "Campo 'Precio litro' incorrecto en la fila " & iFila & ".", vbApplicationModal + vbExclamation, "E R R O R"
            Exit Sub
        End If
        sPrecioLitro = Replace(sPrecioLitro, ",", ".")
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sInfoAux = sLinea
        
        'If Not bDameDatosTarjetaSolred(giPAIS_ESPAÑA, sTarjeta, afoRegTarjeta) Then
        If Not bDameDatosTarjetaSolred(sTarjeta, afoRegTarjeta) Then
            Print #iFicheroSalida, "Tarjeta errónea: " & sTarjeta
            bTarjetaNoEncontrada = True
        End If
        Set afoRegTarjeta = Nothing
        
    Loop
    Close #iFichero
    Close #iFicheroSalida
    
    If bTarjetaNoEncontrada Then
        Me.MousePointer = DEFAULT
        lblTexto(3).Caption = ""
        cmdCargar.Visible = True
        cmdComprobarTarjetas.Visible = True
        lblTexto(3).Visible = False
        Refresh
        MsgBox "Se han encontrado tarjetas sin asignar, revise el fichero '" & gsRutaEquipo & "GeneracionCSV\Sored_Tarjetas_Incorrectas.csv'.", vbApplicationModal + vbInformation, "Información"
        Exit Sub
    End If
    
    Me.MousePointer = DEFAULT
    lblTexto(3).Caption = ""
    cmdCargar.Visible = True
    cmdComprobarTarjetas.Visible = True
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
    cmdComprobarTarjetas.Visible = True
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


