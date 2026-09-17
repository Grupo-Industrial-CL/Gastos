VERSION 5.00
Begin VB.Form frmCargarGastosFichero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar gastos desde fichero"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCargarGastosFichero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVolver 
      Caption         =   "V O L V E R"
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
      Picture         =   "frmCargarGastosFichero.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   9615
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "C A R G A R     G A S T O S"
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
      Picture         =   "frmCargarGastosFichero.frx":1522
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   9615
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "TIPOS PAGO:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   6360
      TabIndex        =   11
      Top             =   1800
      Width           =   1170
   End
   Begin VB.Label etQueTipoPago 
      AutoSize        =   -1  'True
      Caption         =   "M->Metálico V->Visa       S->Solred    X->Varios"
      ForeColor       =   &H000000FF&
      Height          =   780
      Left            =   6360
      TabIndex        =   10
      Top             =   2040
      Width           =   1260
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "SU-> Supervisoras    DA-> Dir. Area    DR-> Dir. Región    DP-> Dir. Apoyo"
      ForeColor       =   &H000000FF&
      Height          =   1035
      Left            =   480
      TabIndex        =   9
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "España-> Pesetas    Portugal-> Escudos    Italia-> Liras        Resto-> Moneda local"
      ForeColor       =   &H000000FF&
      Height          =   1020
      Index           =   1
      Left            =   8160
      TabIndex        =   8
      Top             =   2040
      Width           =   1980
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "IMPORTE GASTO:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   8160
      TabIndex        =   7
      Top             =   1800
      Width           =   1530
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "TIPOS GASTO:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   2880
      TabIndex        =   6
      Top             =   1800
      Width           =   1290
   End
   Begin VB.Label Label5 
      Caption         =   "P->Peaje       F->Telefono V->Varios O->Comida  E->Cena       A->Cafeteria"
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   4560
      TabIndex        =   5
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "B->Billetes       H->Hotel          I->Invitaciones S->Salones     T->Transporte G->Gasolina  K->Kms         "
      ForeColor       =   &H000000FF&
      Height          =   1380
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   1380
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "JERARQUÍA:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Caption         =   $"frmCargarGastosFichero.frx":23EC
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   915
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   9825
   End
End
Attribute VB_Name = "frmCargarGastosFichero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCargar_Click()

    Dim iFichero As Integer
    Dim sLinea As String
    Dim iLinea As Integer
    
    Dim bTransaccionAbierta As Boolean
    
    Dim iCodPais As Integer
    Dim lCodigo As Long
    Dim sJerarquia As String
    Dim sTipoGasto As String
    Dim sFecha As String
    Dim iAnyo As Integer
    Dim iCampanya As Integer
    Dim dImporte As Double
    Dim dNumero As Double
    Dim sTipoPago As String
    Dim sDonde As String
        
    On Error GoTo Error_Rutina
    
    Me.MousePointer = HOURGLASS
    cmdCargar.Enabled = False
    cmdVolver.Enabled = False
    
    iFichero = FreeFile
    Open "C:\GastosFichero.csv" For Input As #iFichero
    
    iLinea = 0
    bTransaccionAbierta = True
    gBdGALILEO.BeginTrans True
    Do While Not EOF(iFichero)
    
        iLinea = iLinea + 1
        Line Input #iFichero, sLinea
        
        sLinea = Trim(sLinea)
        If sLinea = "" Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            MsgBox "Línea vacía en la línea " & iLinea & ". No se ha guardado ningún dato.", vbApplicationModal + vbExclamation, "E R R O R"
            cmdCargar.Enabled = True
            cmdVolver.Enabled = True
            Exit Sub
        End If
        
        iCodPais = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If Not bCodPaisCorrecto(iCodPais) Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            MsgBox "País incorrecto en la línea " & iLinea & ". No se ha guardado ningún dato del fichero.", vbApplicationModal + vbExclamation, "E R R O R"
            cmdCargar.Enabled = True
            cmdVolver.Enabled = True
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        lCodigo = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sJerarquia = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If sJerarquia <> "SU" And sJerarquia <> "DA" And sJerarquia <> "DR" And sJerarquia <> "DP" Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            MsgBox "Jerarquía incorrecta en la línea " & iLinea & ". No se ha guardado ningún dato del fichero.", vbApplicationModal + vbExclamation, "E R R O R"
            cmdCargar.Enabled = True
            cmdVolver.Enabled = True
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sFecha = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        sFecha = Replace(sFecha, "/", "")
        sFecha = Replace(sFecha, "-", "")
                
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        iAnyo = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If iAnyo < 2000 Or iAnyo > 2050 Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            MsgBox "Año no válido en la línea " & iLinea & ". No se ha guardado ningún dato del fichero.", vbApplicationModal + vbExclamation, "E R R O R"
            cmdCargar.Enabled = True
            cmdVolver.Enabled = True
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        iCampanya = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If iCampanya < 1 Or iCampanya > 26 Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            MsgBox "Campaña no válida en la línea " & iLinea & ". No se ha guardado ningún dato del fichero.", vbApplicationModal + vbExclamation, "E R R O R"
            cmdCargar.Enabled = True
            cmdVolver.Enabled = True
            Exit Sub
        End If
        
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sDonde = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        sDonde = Replace(sDonde, "'", " ")
        sDonde = Replace(sDonde, """", " ")
                
        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sTipoGasto = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If sTipoGasto <> "B" And sTipoGasto <> "H" And sTipoGasto <> "I" And sTipoGasto <> "S" And _
           sTipoGasto <> "T" And sTipoGasto <> "G" And sTipoGasto <> "K" And sTipoGasto <> "P" And _
           sTipoGasto <> "F" And sTipoGasto <> "V" And sTipoGasto <> "O" And sTipoGasto <> "E" And sTipoGasto <> "A" Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            MsgBox "Tipo gasto incorrecto en la línea " & iLinea & ". No se ha guardado ningún dato del fichero.", vbApplicationModal + vbExclamation, "E R R O R"
            cmdCargar.Enabled = True
            cmdVolver.Enabled = True
            Exit Sub
        End If

        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        sTipoPago = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        If sTipoPago <> "M" And sTipoPago <> "V" And sTipoPago <> "S" And sTipoPago <> "X" Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            MsgBox "Tipo pago incorrecto en la línea " & iLinea & ". No se ha guardado ningún dato del fichero.", vbApplicationModal + vbExclamation, "E R R O R"
            cmdCargar.Enabled = True
            cmdVolver.Enabled = True
            Exit Sub
        End If

        sLinea = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        dImporte = Mid(sLinea, 1, InStr(1, sLinea, ";") - 1)
        dNumero = Mid(sLinea, InStr(1, sLinea, ";") + 1)
        
        sSql = "INSERT INTO Gastos (CodPais,Codigo,Jerarquia,TipoGasto,Fecha,Anyo,Campanya," & _
               "Importe,Numero,Donde,TipoPago,Identificador,Chequeado,Validado,Validacion,Pagado,Hora,idGrupo,idModulo) " & _
               "VALUES (" & iCodPais & "," & lCodigo & ",'" & sJerarquia & "','" & sTipoGasto & "','" & sFecha & "'," & _
               iAnyo & "," & iCampanya & "," & dImporte & "," & dNumero & ",'" & sDonde & "','" & sTipoPago & "','" & _
               Format(Now, "ddMMyyyy") & "','S','S','" & goUsuario.IdUsuario & "','S',1,0,0)"
        If Not gBdGALILEO.bExecute(sSql) Then
            gBdGALILEO.RollbackTrans
            Close
            Me.MousePointer = DEFAULT
            MsgBox "Error insert en la línea " & iLinea & ". No se ha guardado ningún dato del fichero.", vbApplicationModal + vbExclamation, "E R R O R"
            cmdCargar.Enabled = True
            cmdVolver.Enabled = True
            Exit Sub
        End If
        
        'País - Código - Jerarquía - Fecha - Año - Campaña - Donde - Tipo gasto - Tipo pago - Importe gasto (moneda local) - Número
    Loop
    gBdGALILEO.CommitTrans
    Close
    
    Me.MousePointer = DEFAULT
    MsgBox "Fichero cargado correctamente.", vbApplicationModal + vbInformation, "Información"
    cmdCargar.Enabled = True
    cmdVolver.Enabled = True
        
Exit_Rutina:
    Exit Sub
Error_Rutina:
    If bTransaccionAbierta Then gBdGALILEO.RollbackTrans
    Close
    cmdCargar.Enabled = True
    cmdVolver.Enabled = True
    Me.MousePointer = DEFAULT
    If iLinea > 0 Then
        MsgBox "Error en la línea " & iLinea & ". No se ha guardado ningún dato del fichero.", vbApplicationModal + vbExclamation, "E R R O R"
    Else
        MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    End If
    Resume Exit_Rutina
    Resume Next
End Sub

Private Sub cmdVolver_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Function bCodPaisCorrecto(iCodPais As Integer) As Boolean

    Dim sSql As String
    Dim afoRegistro As ADODB.Recordset

    bCodPaisCorrecto = False

    sSql = "SELECT CodPais FROM Paises WHERE CodPais = " & iCodPais
    bCodPaisCorrecto = gBdGALILEO.bRecordset(sSql, afoRegistro)
    Set afoRegistro = Nothing

End Function



