VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlexGrid2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mrMiGrid As MSFlexGrid
Private miFilasIniciales As Long
Private miLineaSeleccionada As Long
Private miLineaLibre As Long
Private mcltColOrder As New Collection

'Se utilizan para la búsqueda en columna
Private msBufferTeclado As String
Private mlPulsacionAnterior As Long

Public LargeCopy As Boolean

Public Sub aClipBoard(Index As Integer)

   Dim CopyString As String, CellString As String
   Dim GridRLoop As Long, GridCLoop As Long, ColCount As Long
   Dim RowStart As Long, RowEnd As Long, RowCount As Long
   Dim ColStart As Long, ColEnd As Long, LastCol As Long
   Dim FirstRow As Boolean
   
   
   Dim Msg As String
   Dim Title As String
   
   On Error GoTo CopyErr
   Screen.MousePointer = vbHourglass
   Select Case Index
       Case 0 'Copy Selected
           RowStart = mrMiGrid.RowSel
           RowEnd = mrMiGrid.Row
           ColStart = mrMiGrid.ColSel
           ColEnd = mrMiGrid.col
           If RowStart > RowEnd Then
               GridRLoop = RowEnd
               RowEnd = RowStart
               RowStart = GridRLoop
           End If
           If ColStart > ColEnd Then
               GridCLoop = ColEnd
               ColEnd = ColStart
               ColStart = GridCLoop
           End If
       Case Else 'Copy All
           RowStart = 0
           RowEnd = mrMiGrid.Rows - 1
           ColStart = 0
           ColEnd = mrMiGrid.Cols - 1
   End Select
   LastCol = ColEnd
   For GridCLoop = ColEnd To ColStart Step -1
       If mrMiGrid.ColWidth(GridCLoop) > 1 Then
           Exit For
       Else
           LastCol = LastCol - 1
       End If
   Next
   FirstRow = True
   For GridRLoop = RowStart To RowEnd
       If mrMiGrid.RowHeight(GridRLoop) > 1 Then
           For GridCLoop = ColStart To LastCol
               If mrMiGrid.ColWidth(GridCLoop) > 1 Then
                   If FirstRow Then
                       ColCount = ColCount + 1
                   End If
                   CellString = mrMiGrid.TextMatrix(GridRLoop, GridCLoop)
                   If mrMiGrid.MergeRow(GridRLoop) And GridCLoop > ColStart Then
                       If CellString = mrMiGrid.TextMatrix(GridRLoop, (GridCLoop - 1)) Then
                           'we will want to skip the content
                           CellString = ""
                       End If
                   End If
                   If mrMiGrid.MergeCol(GridCLoop) And GridRLoop > RowStart Then
                       If CellString = mrMiGrid.TextMatrix((GridRLoop - 1), GridCLoop) Then
                           CellString = ""
                       End If
                   End If
                   CopyString = CopyString & CellString
                   If GridCLoop = LastCol Then
                       If GridRLoop < RowEnd Then
                           CopyString = CopyString & vbLf
                       End If
                   Else
                       CopyString = CopyString & vbTab
                   End If
               End If
           Next
           RowCount = RowCount + 1
           FirstRow = False
       End If
   Next
   Clipboard.Clear
   Clipboard.SetText CopyString
   Screen.MousePointer = vbDefault
   If Len(CopyString) > 2000 Then
       LargeCopy = True
       MsgBox CStr(RowCount) & " Filas y " & CStr(ColCount) & " Columnas", vbInformation + vbOKOnly, "Copia Completada"
   Else
       LargeCopy = False
   End If
   Exit Sub
CopyErr:
   CopyString = ""
   Select Case Err
       Case 7, 44 'memory and string space problems
           Msg = "Lo siento la selección de lineas excede"
           Msg = Msg & vbLf & "la cantidad disponible de memoria de Windows."
           Msg = Msg & vbLf & "Por favor copie una selección mas pequeña."
           Title = "Recursos de Windows insuficiente"
       Case Else
           Msg = "Lo siento, ha ocurrido un error y la copia"
           Msg = Msg & vbLf & "no se ha completado."
           Msg = Msg & vbLf & "Codigo de error = " & CStr(Err) & vbLf & Error$
           Title = "Copia fallida"
   End Select
   Screen.MousePointer = vbDefault
   MsgBox Msg, vbExclamation, Title
   Exit Sub
End Sub


Public Sub BuscarEnColumna( _
    KeyAscii As Integer, _
    iColumna As Long)
    
' Se utiliza para buscar un texto en una columna.
' Hay un buffer de teclado, que se va incrementando si ejecutamos
' pulsaciones de teclas separadas en el tiempo en menos de un
' segundo. Si se encuentra el  texto, se sitúa el foco en la
' fila donde se halle.
' NOTA: Esta funcion debe introducise dentro
' del evento Keypress del FlexGrid

    Dim iLinea As Long
    Dim iLongBuffer As Integer
    
    If (Timer - mlPulsacionAnterior) < 1 Then
        msBufferTeclado = msBufferTeclado & Chr$(KeyAscii)
    Else
        msBufferTeclado = Chr$(KeyAscii)
    End If
    
    iLongBuffer = Len(msBufferTeclado)
    
    For iLinea = 1 To NumLineasOcupadas
        If LCase$(Mid$(msBufferTeclado, 1, iLongBuffer)) = _
        LCase$(Mid$(TextoLinCol(iLinea, iColumna), 1, iLongBuffer)) Then
            SeleccionarLinea iLinea
            ResaltarLinea
            mrMiGrid.TopRow = iLinea
            Exit For
        End If
    Next
    
    mlPulsacionAnterior = Timer

End Sub


Public Function iBuscarLineaPorTextoCol(sTextoCol As String, iColumna As Long) As Long

    Dim iLinea As Long
    
    iBuscarLineaPorTextoCol = -1
    For iLinea = 1 To NumLineasOcupadas
        If Trim$(TextoLinCol(iLinea, iColumna)) = Trim$(sTextoCol) Then
            iBuscarLineaPorTextoCol = iLinea
            Exit For
        End If
    Next

End Function

Public Sub Create( _
    elGrid As MSFlexGrid, _
    yCols As Long, _
    Optional yRows As Long = 1, _
    Optional yFixedCols As Long = 0, _
    Optional yFixedRows As Long = 1)
        
    Set mrMiGrid = elGrid
    
    With mrMiGrid
        .Cols = yCols
        .Rows = yRows
        .FixedCols = yFixedCols
        .FixedRows = yFixedRows
    End With
    
    miLineaSeleccionada = -1
    miLineaLibre = yFixedRows
    miFilasIniciales = yRows
    mlPulsacionAnterior = 0
End Sub

Public Sub Formatear_2( _
    col As Long, _
    Fil As Long, _
    Ancho As Integer, _
    Titulo As String, _
    Optional FixAling As Integer = 0, _
    Optional ColAling As Integer = 0, _
    Optional ColOrder As Boolean = False, _
    Optional lColor As Long)
    
    mrMiGrid.Row = Fil
    mrMiGrid.col = col
    mrMiGrid.FixedAlignment(col) = FixAling
    mrMiGrid.TextMatrix(Fil, col) = Titulo
    If col >= mrMiGrid.FixedCols Then
        mrMiGrid.ColAlignment(col) = ColAling
    End If
    mrMiGrid.ColWidth(col) = Ancho
    If ColOrder Then
        mcltColOrder.Add col
    End If
    
    If lColor Then
        mrMiGrid.FillStyle = flexFillRepeat
        mrMiGrid.ColSel = col
        mrMiGrid.CellBackColor = lColor
    End If
    
End Sub

Public Sub Formatear( _
    ByVal col As Long, _
    ByVal Ancho As Integer, _
    ByVal Titulo As String, _
    Optional ByVal FixAling As Integer = 0, _
    Optional ByVal ColAling As Integer = 0, _
    Optional ByVal ColOrder As Boolean = False)
    
    mrMiGrid.Row = 0
    mrMiGrid.col = col
    mrMiGrid.FixedAlignment(col) = FixAling
    mrMiGrid.TextMatrix(0, col) = Titulo
    If col >= mrMiGrid.FixedCols Then
        mrMiGrid.ColAlignment(col) = ColAling
    End If
    mrMiGrid.ColWidth(col) = Ancho
    If ColOrder Then
        mcltColOrder.Add col
    End If
    
End Sub

Public Sub CambioColumnaEuros(iColumna As Long, iCodPais As Integer)
        
    Dim iLinea As Long
    Dim dCambio As Double
    Dim dValorinter As Double
    
    Dim sValorinter As String
    
    Select Case iCodPais
        Case 34, 99
            dCambio = 166.386
        Case 51
            dCambio = 200.482
        Case 39
            dCambio = 1936.27
            'Nacho rufo: añado esto para que se pueda llamar por todas y solo cambie en España, Portugal e Italia
        Case Else
            dCambio = 1
    End Select
    'Nacho rufo: condición para el nuevo cambio
    If dCambio <> 1 Then
        For iLinea = 1 To NumLineasOcupadas
            If IsNumeric(mrMiGrid.TextMatrix(iLinea, iColumna)) Then
                sValorinter = Format$(mrMiGrid.TextMatrix(iLinea, iColumna), "#########0.00")
              '  dValorinter = sValorinter
                dValorinter = Val(sValorinter)
                mrMiGrid.TextMatrix(iLinea, iColumna) = _
                    Format$(Round(Val(Format$(dValorinter / dCambio, "#.000")), 2), "#.00")
                
            End If
        Next
    End If

End Sub
Public Sub CambioColumnaMlocal2(iColumna As Long, iCodPais As Integer)
        
    Dim iLinea As Long
    Dim dCambio As Double
    Dim dValorinter As Double
    
    Dim sValorinter As String
    
    Select Case iCodPais
        Case 34, 99
            dCambio = 166.386
        Case 51
            dCambio = 200.482
        Case 39
            dCambio = 1936.27
    End Select
    
    For iLinea = 1 To NumLineasOcupadas
        If IsNumeric(mrMiGrid.TextMatrix(iLinea, iColumna)) Then
            sValorinter = Format$(mrMiGrid.TextMatrix(iLinea, iColumna), "#########0.00")
          '  dValorinter = sValorinter
            dValorinter = Val(sValorinter)
            mrMiGrid.TextMatrix(iLinea, iColumna) = _
                Format$(Round(Val(Format$(dValorinter * dCambio, "#.000")), 2), "#.00")
            
        End If
    Next

End Sub


Public Sub CambioColumnaMonedaLocal(iColumna As Long, iCodPais As Integer, iSalto As Integer)
        
    Dim iLinea As Long
    Dim dCambio As Double
    
    Select Case iCodPais
        Case 34, 99
            dCambio = 166.386
        Case 51
            dCambio = 200.482
        Case 39
            dCambio = 1936.27
    End Select
    
    For iLinea = 1 To NumLineasOcupadas
        If IsNumeric(mrMiGrid.TextMatrix(iLinea, iColumna)) Then
            
       '     mrMiGrid.TextMatrix(iLinea, iColumna) = _
               Format((Val(Val(mrMiGrid.TextMatrix(iLinea, iColumna)) * dCambio)), "###,##0")
            
          mrMiGrid.TextMatrix(iLinea, iColumna) = mrMiGrid.TextMatrix(iLinea, iColumna + iSalto)
            
        End If
    Next

End Sub

Public Sub AddLinea(iFila As Long, ParamArray aTexts() As Variant)

    Dim i As Long
    Dim bAvisado As Boolean
    Dim objTexts
    Dim iTotalParametros As Integer
    
    bAvisado = False
    For Each objTexts In aTexts
        iTotalParametros = iTotalParametros + 1
    Next
    
    With mrMiGrid
        If iFila > .Rows - .FixedRows Then
            .Rows = .Rows + 1
        End If
        
        .AddItem "", iFila
        For i = 0 To .Cols - 1
            If i < iTotalParametros Then
                If IsNull(aTexts(i)) Then
                    .TextMatrix(iFila, i) = ""
                Else
                    .TextMatrix(iFila, i) = aTexts(i)
                End If
            Else
                .TextMatrix(iFila, i) = ""
            End If
        Next
    End With
    
    If miLineaLibre < 32766 Then
        miLineaLibre = mrMiGrid.Rows
    Else
        If Not bAvisado Then
            MsgBox "Se ha excedido el tamaño del Grid. Se verán 32767 filas"
        End If
    End If
End Sub

Public Sub TextoFila(iFila As Long, ParamArray aTexts() As Variant)
    
    Dim i As Long
    Dim bAvisado As Boolean
    Dim objTexts
    Dim iTotalParametros As Integer
    
    bAvisado = False
    iTotalParametros = 0
    For Each objTexts In aTexts
        iTotalParametros = iTotalParametros + 1
    Next

    With mrMiGrid
        If iFila > .Rows - .FixedRows Then
            .Rows = .Rows + 1
        End If
        
        For i = 0 To .Cols - 1
            If i < iTotalParametros Then
                If IsNull(aTexts(i)) Then
                    .TextMatrix(iFila, i) = ""
                Else
                    .TextMatrix(iFila, i) = aTexts(i)
                End If
            Else
                .TextMatrix(iFila, i) = ""
            End If
        Next
    End With
    
    If miLineaLibre > 32766 Then
   '     miLineaLibre = mrMiGrid.Rows
   ' Else
            MsgBox "Se ha excedido el tamaño del Grid. Se verán 32767 filas"
     End If
   ' End If
       miLineaLibre = iFila + 1
End Sub

Public Sub TextoFila2(iFila As Long, ParamArray aTexts() As Variant)
    
    Dim i As Long
    Dim bAvisado As Boolean
    Dim objTexts
    Dim iTotalParametros As Integer
    
    bAvisado = False
    iTotalParametros = 0
    For Each objTexts In aTexts
        iTotalParametros = iTotalParametros + 1
    Next

    With mrMiGrid
        If iFila > .Rows - 1 Then
            .Rows = .Rows + 1
        End If
        
        For i = 0 To .Cols - 1
            If i < iTotalParametros Then
                If IsNull(aTexts(i)) Then
                    .TextMatrix(iFila, i) = ""
                Else
                    .TextMatrix(iFila, i) = aTexts(i)
                End If
            Else
                .TextMatrix(iFila, i) = ""
            End If
        Next
    End With
    
    If miLineaLibre > 32766 Then
   '     miLineaLibre = mrMiGrid.Rows
   ' Else
            MsgBox "Se ha excedido el tamaño del Grid. Se verán 32767 filas"
     End If
   ' End If
       miLineaLibre = iFila + 1
End Sub

Public Sub ColorCelda(iFila As Long, iColumna As Long, lColor As Long)
    
    With mrMiGrid
        .Row = iFila
        .col = iColumna
        .FillStyle = flexFillRepeat
        .ColSel = iColumna
        .CellBackColor = lColor
    End With
    
End Sub

Public Sub OcultaFila(iFila As Long)
' Mod Julián Petición Posibles Distribuidoras
    mrMiGrid.RowHeight(iFila) = 0
End Sub

Public Sub ColorFila(iFila As Long, lColor As Long)
    
    With mrMiGrid
        .Row = iFila
        .col = .FixedCols
        .FillStyle = flexFillRepeat
        .ColSel = .Cols - 1
        .CellBackColor = lColor
    End With
    
End Sub

Public Sub ColorFilaActual(lColor As Long)
    
    With mrMiGrid
        '.Row = iFila
        .col = .FixedCols
        .FillStyle = flexFillRepeat
        .ColSel = .Cols - 1
        .CellBackColor = lColor
    End With
    
End Sub

Public Sub ColorColumna(iFila As Long, iColumna As Long, lColor As Long)
    
    With mrMiGrid
        .Row = iFila
        .col = iColumna
        .FillStyle = flexFillRepeat
        .ColSel = iColumna
        .CellBackColor = lColor
    End With
    
End Sub

Public Sub ColorColumnaEntera(iColumna As Long, lColor As Long)
    
    With mrMiGrid
        .Row = .FixedRows
        .col = iColumna
        .FillStyle = flexFillRepeat
        .RowSel = .Rows - 1
        .CellBackColor = lColor
    End With
    
End Sub

Public Function DameColorCelda(iFila As Long, iColumna As Long) As Long
    
    With mrMiGrid
        .Row = iFila
        .col = iColumna
        DameColorCelda = .CellBackColor
    End With
    
End Function

Public Sub ResaltarLinea()
    
    With mrMiGrid
        .HighLight = flexHighlightNever
        .col = 0
        .ColSel = .Cols - 1
        .HighLight = flexHighlightAlways
        miLineaSeleccionada = .Row
    End With
    
End Sub

Public Sub MarcarLinea(Optional sMarca As String = "->")

    Dim iFila As Long
    
    With mrMiGrid
    
        .col = 0
        If miLineaSeleccionada <> -1 Then
            iFila = .Row
            .Row = miLineaSeleccionada
            .Text = ""
            .Row = iFila
        End If
        .Text = sMarca
        miLineaSeleccionada = .Row
        
    End With
    
End Sub

Public Sub ColocarMousePointer()

    Dim vCol As Variant
    
    If mrMiGrid.MouseRow = 0 Then
        For Each vCol In mcltColOrder
            If mrMiGrid.MouseCol = vCol Then
                mrMiGrid.MousePointer = vbUpArrow
                mrMiGrid.ToolTipText = "Ordenar"
                Exit For
            Else
                mrMiGrid.MousePointer = vbDefault
                mrMiGrid.ToolTipText = ""
            End If
        Next
    Else
        mrMiGrid.MousePointer = vbDefault
        mrMiGrid.ToolTipText = ""
    End If

End Sub

Public Sub OrdenarSiCorresponde(Optional iModoOrden As Integer = 1)

    Dim vCol As Variant
    Dim iNumColumna As Long
    Dim i As Long
    
  '  iNumColumna = mrMiGrid.MouseCol
    iNumColumna = mrMiGrid.col
    If mrMiGrid.MouseRow = 0 Then
        For i = 0 To miLineaLibre - 1
            mrMiGrid.Row = i
            If bLineaActualVacia Then
                mrMiGrid.Rows = i
                miLineaLibre = i
                Exit For
            End If
        Next
        For Each vCol In mcltColOrder
           If iNumColumna = vCol Then
                mrMiGrid.col = vCol
                mrMiGrid.Sort = iModoOrden
                Exit For
            Else
                mrMiGrid.MousePointer = vbDefault
            End If
        Next
    Else
        mrMiGrid.MousePointer = vbDefault
    End If
    If miLineaLibre < miFilasIniciales Then
        miLineaLibre = miFilasIniciales
        mrMiGrid.Rows = miFilasIniciales
    End If

End Sub

Public Function bLineaActualVacia()
       
    Dim i As Long
    
    If mrMiGrid.Row <= mrMiGrid.FixedRows - 1 Then
        bLineaActualVacia = False
    Else
        bLineaActualVacia = True
            
        i = 1
        Do While i <= mrMiGrid.Cols - 1
            mrMiGrid.col = i
            If mrMiGrid.Text <> "" Then
                bLineaActualVacia = False
                Exit Do
            End If
            i = i + 1
        Loop
    End If
  
End Function

Sub Vaciar()

    With mrMiGrid
        .Rows = .FixedRows
        .Rows = miFilasIniciales
    End With
    
    miLineaLibre = mrMiGrid.FixedRows
    miLineaSeleccionada = -1
        
End Sub

Public Property Get TextoCol(iCol As Long) As String

    TextoCol = mrMiGrid.TextMatrix(mrMiGrid.Row, iCol)
    
End Property

Public Property Let TextoCol(iCol As Long, vValor As String)

    mrMiGrid.TextMatrix(mrMiGrid.Row, iCol) = vValor
    
End Property

Public Property Let TextoLinCol( _
    ByVal iLin As Long, _
    ByVal iCol As Long, _
    sTexto As String)
    
    mrMiGrid.TextMatrix(iLin, iCol) = sTexto
    
End Property

Public Property Get TextoLinCol( _
    ByVal iLin As Long, _
    ByVal iCol As Long) As String
    
    TextoLinCol = mrMiGrid.TextMatrix(iLin, iCol)

End Property

Public Function sTextoColLineaSel(iCol As Long)
    
    If miLineaSeleccionada <> -1 Then
         sTextoColLineaSel = mrMiGrid.TextMatrix(miLineaSeleccionada, iCol)
    Else
        sTextoColLineaSel = ""
    End If
    
End Function

Public Sub NoResaltarLineas()
    mrMiGrid.HighLight = flexHighlightNever
    miLineaSeleccionada = -1
End Sub

Public Sub ElimLinea()

    Dim iLineaBorrar As Long
    Dim iCol As Long
    
    iLineaBorrar = miLineaSeleccionada
    If iLineaBorrar <> -1 Then
        If mrMiGrid.Rows = mrMiGrid.FixedRows + 1 Then
            mrMiGrid.Row = mrMiGrid.FixedRows
            For iCol = 0 To mrMiGrid.Cols - 1
                mrMiGrid.col = iCol
                mrMiGrid.Text = ""
            Next
            miLineaLibre = mrMiGrid.FixedRows
            ColorFila miLineaLibre, mrMiGrid.BackColor
        Else
            mrMiGrid.RemoveItem iLineaBorrar
            miLineaLibre = mrMiGrid.Rows
        End If
        miLineaSeleccionada = -1
        mrMiGrid.HighLight = flexHighlightNever
    End If

End Sub

Public Property Get LineaSeleccionada() As Long
    LineaSeleccionada = miLineaSeleccionada
End Property

Public Property Get LineaLibre() As Long
    LineaLibre = miLineaLibre
End Property

Public Property Get NumLineasOcupadas() As Long

    If miLineaLibre = mrMiGrid.FixedRows Then
        NumLineasOcupadas = 0
    Else
        NumLineasOcupadas = mrMiGrid.Rows - mrMiGrid.FixedRows
    End If
    
End Property

Public Sub SeleccionarLinea(iLinea As Long)
    mrMiGrid.Row = iLinea
End Sub

Public Property Get Grid() As MSFlexGrid
    Set Grid = mrMiGrid
End Property

Public Property Get ColorDeFila(iFila As Long) As Long
    
    mrMiGrid.Row = iFila
    mrMiGrid.col = 1
    ColorDeFila = mrMiGrid.CellBackColor
    
End Property

Sub GeneraCSV(Optional sSeparador As String = ";", Optional sNombreFichero As String = "")

' Genera Fichero csv con  el contenido de un grid

   Dim iColumna As Long
   Dim iFila As Long
   Dim sLinea As String
   Dim sArchivo As String
   Dim sFichero As String
   


    On Error GoTo error_generacsv
    
    Screen.MousePointer = vbHourglass
    
    'sFichero = mrMiGrid.Name
    sFichero = Trim$(sNombreFichero)
    If sFichero = "" Then sFichero = mrMiGrid.Name
    
    sArchivo = Dir$("C:\GeneracionCSV", vbDirectory)
    If sArchivo = "" Then
        MkDir ("C:\GeneracionCSV")
    End If
    
    If sFichero = mrMiGrid.Name Then
        sArchivo = "C:\GeneracionCSV\" & sFichero & "_" & Format$(Date, "dd-mm-yy") & _
            "_" & Format$(Time, "hh-mm") & ".csv"
    Else
        sArchivo = "C:\GeneracionCSV\" & sFichero & ".csv"
    End If
    
    Open sArchivo For Output As #1

    ' Genera lineas
    ' --------------------------------------------------------
    For iFila = 0 To mrMiGrid.Rows - 1
        mrMiGrid.Row = iFila
        sLinea = ""
        For iColumna = 0 To mrMiGrid.Cols - 1
           mrMiGrid.col = iColumna
           sLinea = sLinea & Trim$(mrMiGrid.Text) & sSeparador
        Next
        Print #1, sLinea
     Next
       Screen.MousePointer = vbDefault
       Close #1
    'Exit Sub
    
    MsgBox "Fichero generado en: " & sArchivo, vbApplicationModal + vbInformation, "Información"
      
error_generacsv:
        If Err.Number <> 0 Then
            MsgBox "Error " & Err.Number & " GenerandoCSV." & vbCrLf & _
                    "Descripcion: " & Err.Description, vbCritical
            Screen.MousePointer = vbDefault
        End If
End Sub

Sub GeneraCSVCitrix2(Optional sSeparador As String = ";", Optional sNombreFichero As String = "")

' Genera Fichero csv con  el contenido de un grid

   Dim iColumna As Long
   Dim iFila As Long
   Dim sLinea As String
   Dim sArchivo As String
   Dim sFichero As String
   


    On Error GoTo error_generacsv
    
    Screen.MousePointer = vbHourglass
    
    'sFichero = mrMiGrid.Name
    sFichero = Trim$(sNombreFichero)
    If sFichero = "" Then sFichero = mrMiGrid.Name
    
    sArchivo = Dir$("V:\GeneracionCSV", vbDirectory)
    If sArchivo = "" Then
        MkDir ("V:\GeneracionCSV")
    End If
    
    If sFichero = mrMiGrid.Name Then
        sArchivo = "V:\GeneracionCSV\" & sFichero & "_" & Format$(Date, "dd-mm-yy") & _
            "_" & Format$(Time, "hh-mm") & ".csv"
    Else
        sArchivo = "V:\GeneracionCSV\" & sFichero & ".csv"
    End If
    
    Open sArchivo For Output As #1

    ' Genera lineas
    ' --------------------------------------------------------
    For iFila = 0 To mrMiGrid.Rows - 1
        mrMiGrid.Row = iFila
        sLinea = ""
        For iColumna = 0 To mrMiGrid.Cols - 1
           mrMiGrid.col = iColumna
           sLinea = sLinea & Trim$(mrMiGrid.Text) & sSeparador
        Next
        Print #1, sLinea
     Next
       Screen.MousePointer = vbDefault
       Close #1
    Exit Sub
    
    MsgBox "Fichero generado en: " & sArchivo, vbApplicationModal + vbInformation, "Información"
      
error_generacsv:
        MsgBox "Error " & Err.Number & " GenerandoCSV." & vbCrLf & _
               "Descripcion: " & Err.Description, vbCritical
        Screen.MousePointer = vbDefault
End Sub

Sub GeneraCSVCitrix()

' Genera Fichero csv con  el contenido de un grid

   Dim iColumna As Long
   Dim iFila As Long
   Dim sLinea As String
   Dim sArchivo As String
   Dim sFichero As String
   


    On Error GoTo error_generacsvC
    
    Screen.MousePointer = vbHourglass
    
    sFichero = mrMiGrid.Name
    
    
    sArchivo = Dir$("V:\GeneracionCSV", vbDirectory)
    If sArchivo = "" Then
        MkDir ("V:\GeneracionCSV")
    End If
    
    sArchivo = "V:\GeneracionCSV\" & sFichero & "_" & Format$(Date, "dd-mm-yy") & _
        "_" & Format$(Time, "hh-mm") & ".csv"
    
    Open sArchivo For Output As #1

    ' Genera lineas
    ' --------------------------------------------------------
    For iFila = 0 To mrMiGrid.Rows - 1
        mrMiGrid.Row = iFila
        sLinea = ""
        For iColumna = 0 To mrMiGrid.Cols - 1
           mrMiGrid.col = iColumna
           sLinea = sLinea & Trim$(mrMiGrid.Text) & "; "
        Next
        Print #1, sLinea
     Next
       Screen.MousePointer = vbDefault
       Close #1
       
       MsgBox "Fichero generado en: " & sArchivo, vbApplicationModal + vbInformation, "Información"
    Exit Sub
    
error_generacsvC:
        MsgBox "Error " & Err.Number & " GenerandoCSV." & vbCrLf & _
               "Descripcion: " & Err.Description, vbCritical
        Screen.MousePointer = vbDefault
End Sub


Sub Imprime(sTitulo1 As String, sTitulo2 As String, _
    Optional bApaisado As Boolean = True, _
    Optional lSize As Long = 10, _
    Optional sSegundaLinea As String = "")

' Imprime el contenido de un grid en tipo de letra courier
' para impresoras seriales

   Dim iColumna As Long
   Dim iFila As Long
   Dim sLinea As String
   Dim iAncho As Integer
   Dim iAlinea As Integer
   Dim iMargen As Integer
   Dim iMediomargen As Integer
   Dim sSubrayado As String
   
' Imprime cabecera
' --------------------------------------------------------

    On Error GoTo error_imprime_grid
    
    Screen.MousePointer = vbHourglass
    If bApaisado Then
        Printer.Orientation = vbPRORLandscape
    Else
        Printer.Orientation = vbPRORPortrait
    End If
    Printer.Font.Name = "Courier New"
    Printer.FontSize = lSize
  
    Printer.FontBold = True
    Printer.Print sTitulo2; "     "; sTitulo1; "       a  " & Format$(Now, "dd/mm/yyyy")
    Printer.Print " "
    Printer.FontBold = False
    If sSegundaLinea <> "" Then
        Printer.Print sSegundaLinea
        Printer.Print " "
    End If
    ' Imprime líneas
    ' --------------------------------------------------------
    For iFila = 0 To mrMiGrid.Rows - 1
        mrMiGrid.Row = iFila
        sLinea = ""
        For iColumna = 0 To mrMiGrid.Cols - 1
            '''RAUL'If iColumna = 2 Or iColumna = 4 Or iColumna = 5 Or iColumna = 6 Or iColumna = 12 Then
            mrMiGrid.col = iColumna
            ' Posiciona texto
            iAncho = Int(mrMiGrid.ColWidth(iColumna) / 70)
            If iAncho <> 0 Then '****************
            iMargen = iAncho - Len(Trim$(mrMiGrid.Text))
            If iMargen < 1 Then iMargen = 0
            iMediomargen = Int(iMargen / 2)
            If iColumna >= mrMiGrid.FixedCols And iFila >= mrMiGrid.FixedRows Then
               iAlinea = mrMiGrid.ColAlignment(iColumna)
            Else
               iAlinea = mrMiGrid.FixedAlignment(iColumna)
            End If
            Select Case iAlinea
               Case flexAlignLeftCenter, flexAlignLeftTop, flexAlignLeftBottom
                  sLinea = sLinea & Trim$(mrMiGrid.Text) & Space$(iMargen + 1)
               Case flexAlignCenterCenter, flexAlignCenterTop, flexAlignCenterBottom
                  sLinea = sLinea & Space$(iMediomargen) & Trim$(mrMiGrid.Text) & Space$(iMargen - iMediomargen + 1)
               Case flexAlignRightCenter, flexAlignRightTop, flexAlignRightBottom
                  sLinea = sLinea & Space$(iMargen) & Trim$(mrMiGrid.Text) & Space$(1)
            End Select
            End If
            ' Carga subrayado de los titulos
            If iFila = 0 And mrMiGrid.FixedRows > 0 Then
               sSubrayado = sSubrayado & String$(iAncho, "_") & Space$(1)
            End If
            '''RAUL'End If
          Next
          Printer.Print sLinea
    ' Comprueba si es línea de títulos y deja una línea en blanco
    ' -----------------------------------------------------------
          If mrMiGrid.FixedRows = iFila + 1 Then
             Printer.Print sSubrayado
             Printer.Print " "
             Printer.Print " "
          End If
       Next
       Printer.EndDoc
       Screen.MousePointer = vbDefault
    Exit Sub
    
error_imprime_grid:
        MsgBox "Error " & Err.Number & " imprimendo grid." & vbCrLf & _
               "Descripcion: " & Err.Description, vbCritical
        Screen.MousePointer = vbDefault
End Sub

Public Property Get UltimaLineaRellena() As Long

Dim iFila As Long

UltimaLineaRellena = 0

While UltimaLineaRellena = 0 And iFila < mrMiGrid.Rows
    iFila = iFila + 1
    mrMiGrid.Row = iFila
    mrMiGrid.col = 1
        
    If mrMiGrid.Text = "" Then UltimaLineaRellena = iFila
    
Wend

    If UltimaLineaRellena = 0 Then UltimaLineaRellena = mrMiGrid.Rows + 1

End Property

Public Function ObtenerTextoFila(ByVal iFila As Long) As Variant
    
    Dim i As Long
    Dim bAvisado As Boolean
    Dim aTexts() As Variant
    
    bAvisado = False
    With mrMiGrid
        If iFila > .Rows - .FixedRows Then
            .Rows = .Rows + 1
        End If
        
        ReDim aTexts(.Cols - 1) As Variant
        For i = 0 To .Cols - 1
            aTexts(i) = .TextMatrix(iFila, i)
        Next
    End With
    
    If miLineaLibre < 32766 Then
        miLineaLibre = mrMiGrid.Rows
    Else
        If Not bAvisado Then
            MsgBox "Se ha excedido el tamaño del Grid. Se verán 32767 filas"
        End If
    End If
    
    ObtenerTextoFila = aTexts()
End Function


Public Function ObtenerTextoCelda(ByVal iFila As Long, ByVal sNombreColumna As String) As String
    
    Dim i As Long
    Dim aTexts() As Variant
    Dim iColumnaValida As Long
    
    ObtenerTextoCelda = ""
    
    With mrMiGrid
        If .FixedRows < 1 Then
            MsgBox "No se pueden obtener los datos de celda sin una fila de titulos.", vbApplicationModal
            Exit Function
        End If
        
        iColumnaValida = 1000
        ReDim aTexts(.Cols - 1) As Variant
        aTexts = ObtenerTextoFila(0)
        For i = 0 To .Cols - 1
            If Trim$(UCase$(aTexts(i))) = Trim$(UCase$(sNombreColumna)) Then
                iColumnaValida = i
                Exit For
            End If
        Next
        
        If iColumnaValida < 1000 Then
            aTexts = ObtenerTextoFila(iFila)
            ObtenerTextoCelda = aTexts(iColumnaValida)
        Else
            MsgBox "No hay ninguna columna en el Grid '" & .Name & "' llamada '" & sNombreColumna & "'.", vbApplicationModal
            Exit Function
        End If
    End With
    
    If miLineaLibre < 32766 Then
        miLineaLibre = mrMiGrid.Rows
    Else
        MsgBox "Se ha excedido el tamaño del Grid. Se verán 32767 filas"
    End If
End Function

Public Sub OrdenarPorUnaColumna(Optional iModoOrden As Integer = 0)

    Dim vCol As Variant
    Dim iNumColumna As Long
    Dim i As Long
    
    mrMiGrid.AddItem "", 1
    iNumColumna = mrMiGrid.MouseCol
    If mrMiGrid.MouseRow = 0 Then
        For i = 0 To miLineaLibre - 1
            mrMiGrid.Row = i
            If bLineaActualVacia Then
                mrMiGrid.Rows = i
                miLineaLibre = i
                Exit For
            End If
        Next
        For Each vCol In mcltColOrder
           If iNumColumna = vCol Then
                mrMiGrid.col = vCol
                mrMiGrid.Sort = iModoOrden
                Exit For
            Else
                mrMiGrid.MousePointer = vbDefault
            End If
        Next
    Else
        mrMiGrid.MousePointer = vbDefault
    End If
    If miLineaLibre < miFilasIniciales Then
        miLineaLibre = miFilasIniciales
        mrMiGrid.Rows = miFilasIniciales
    End If

End Sub

Public Sub OrdenarPorUnaColumnaDada(iCol_Orden As Integer, Optional iModoOrden As Integer = 1)

    Dim vCol As Variant
    
    For Each vCol In mcltColOrder
        If vCol = iCol_Orden Then
            mrMiGrid.col = vCol
            mrMiGrid.Sort = iModoOrden
            Exit For
        Else
            mrMiGrid.MousePointer = vbDefault
        End If
    Next
    
    If miLineaLibre < miFilasIniciales Then
        miLineaLibre = miFilasIniciales
        mrMiGrid.Rows = miFilasIniciales
    End If

End Sub

Public Sub SituarEn(iCol As Long, iFil As Long)
    
    mrMiGrid.col = iCol
    mrMiGrid.Row = iFil

End Sub
Public Sub Desplazar(iFil As Long)
    
    If mrMiGrid.Row > iFil Then
        mrMiGrid.TopRow = mrMiGrid.Row - iFil
    End If

End Sub


Public Sub TextoVacioFila(iFila As Long)
    
    Dim i As Long
    Dim bAvisado As Boolean
    
    bAvisado = False
    With mrMiGrid
        If iFila > .Rows - .FixedRows Then
            .Rows = .Rows + 1
        End If
        
        For i = 0 To .Cols - 1
            .TextMatrix(iFila, i) = ""
        Next
    End With
    
    If miLineaLibre > 32766 Then
   '     miLineaLibre = mrMiGrid.Rows
   ' Else
            MsgBox "Se ha excedido el tamaño del Grid. Se verán 32767 filas"
     End If
   ' End If
       miLineaLibre = iFila + 1
End Sub

Public Function iBuscarLinea(sTexto1 As String, sTexto2 As String) As Long

     Dim iLinea As Long
    
     iBuscarLinea = -1
    
     For iLinea = 1 To NumLineasOcupadas
     
          If Trim$(TextoLinCol(iLinea, 2)) = Trim$(sTexto1) And _
             Trim$(TextoLinCol(iLinea, 3)) = Trim$(sTexto2) Then
               
               iBuscarLinea = iLinea
               Exit For
            
          End If
     Next
     
End Function



