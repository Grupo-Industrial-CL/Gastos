VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBaseDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Envolvera una conexión a base de datos, proporcionando metodos para facilitar el acceso."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''---------------- CBaseDatos ----------------------------------'pp
'Envuelve una conexión a una base de datos, proporcionando
'métodos que facilitan los accesos y controlan los errores
'---------------------------------------------------------------
Option Explicit

Private mCn As ADODB.Connection 'conexión a la base de datos
'Indica si hay que mostrar los mensajes de error
Private mbModoMostrarErrores As Boolean
'Indica si hay que capturar los errores en los acceso a la BD
Private mbModoCapturarErrores As Boolean
'Registros afectados en una modificación
Private mlRegistrosAfectados As Long

Private mbCreadoConExito As Boolean

Private Sub Class_Initialize()
    mbModoMostrarErrores = False
    mbModoCapturarErrores = True
End Sub

Public Sub CreatePorCadenaDeConexion(sCadenaConexion As String)

    Dim cn As ADODB.Connection
    On Error GoTo err_conexion
        
    Set cn = New ADODB.Connection
    
    cn.ConnectionString = sCadenaConexion
    cn.ConnectionTimeout = 30
    cn.CommandTimeout = 100
    cn.CursorLocation = 3
    cn.Open
    Set mCn = cn
    mbCreadoConExito = True
    
   Exit Sub
   
err_conexion:
    GestionarError
End Sub

Public Sub CreatePorUDL(sFicheroUDL As String)
    
    Dim cn As ADODB.Connection
    On Error GoTo err_conexion
        
    Set cn = New ADODB.Connection
    
    cn.ConnectionString = "FILE NAME=" & sFicheroUDL
    cn.ConnectionTimeout = 30
    cn.CommandTimeout = 100
    cn.CursorLocation = 3
    cn.Open
    Set mCn = cn
    mbCreadoConExito = True
    
   Exit Sub
   
err_conexion:
    GestionarError

End Sub

Public Function cnAbrirConexionImasD( _
    sDsn As String, _
    sBD As String, _
    sUid As String, _
    sPWD As String, _
    Optional iCursorLocation = adUseClient)
    
    Dim cn As ADODB.Connection
    On Error GoTo err_conexion
         
    Set cn = New ADODB.Connection
    
    cn.ConnectionString = "provider=SQLNCLI10;Data Source = " & sDsn & "; InitialCatalog = " & sBD & "; user id = " & sUid & "; password = " & sPWD
    'cn.ConnectionString = "provider=SQLNCLI10;Data Source = " & "1.0.6.87" & "; InitialCatalog = " & sBD & "; user id = " & sUid & "; password = " & sPWD
    
    cn.ConnectionTimeout = 30
    cn.CursorLocation = iCursorLocation
    cn.Open
    
    Set mCn = cn

    mbCreadoConExito = True

   Exit Function
   
err_conexion:
    GestionarError

End Function

Public Sub CreateOLEDB( _
    sProveedor As String, _
    sSrv As String, _
    sBD As String, _
    sUid As String, _
    Optional sPWD As String = "", _
    Optional iCursorLocation = adUseClient)
    
    Dim cn As ADODB.Connection
    On Error GoTo err_conexion
        
    Set cn = New ADODB.Connection
        
 
    cn.ConnectionString = "Provider=" & sProveedor & ";Persist Security Info=False;User ID=" & sUid & ";Password=" & sPWD & ";Initial Catalog=" & sBD & ";Data Source=" & sSrv
    
    cn.ConnectionTimeout = 30
    cn.CursorLocation = iCursorLocation
    cn.Open
    Set mCn = cn
    mbCreadoConExito = True
    
   Exit Sub
   
err_conexion:
    GestionarError
    
End Sub

Public Sub Create( _
    sDsn As String, _
    sBD As String, _
    sUid As String, _
    sPWD As String, _
    Optional iCursorLocation = adUseClient)
    
    Dim cn As ADODB.Connection
    On Error GoTo err_conexion
        
    Set cn = New ADODB.Connection
    
    cn.ConnectionString = "DSN=" & sDsn & ";UID=" & sUid & _
        ";PWD=" & sPWD & ";DATABASE=" & sBD
    cn.ConnectionTimeout = 30
    cn.CommandTimeout = 100
    cn.CursorLocation = iCursorLocation
    cn.Open
    Set mCn = cn
    mbCreadoConExito = True
    
   Exit Sub
   
err_conexion:
    GestionarError
    
End Sub

Public Function bRecordset( _
    sSql As String, _
    arsRecordset As ADODB.Recordset, _
    Optional TipoCursor As CursorTypeEnum = adOpenForwardOnly, _
    Optional TipoBloqueo As LockTypeEnum = adLockReadOnly, _
    Optional ByRef CursorSinDatos As Boolean = True) As Boolean

    On Error GoTo error_consulta
    
    Set arsRecordset = New ADODB.Recordset
    CursorSinDatos = True

    arsRecordset.Open sSql, mCn, TipoCursor, TipoBloqueo, adCmdText
    If Not arsRecordset.EOF Then
        bRecordset = True
        CursorSinDatos = False
    End If
    
    Exit Function
    
error_consulta:
    CursorSinDatos = False
    GestionarError
    
End Function

Public Function bExecute(ByRef sSql As String, Optional lRegAfectados As Long, Optional lSegundos As Single, Optional sMensajeError As String) As Boolean

    On Error GoTo error_execute
    
    Dim InicioTime As Single
    
    mlRegistrosAfectados = 0
    sMensajeError = ""
    InicioTime = Timer
    
    mCn.Execute sSql, mlRegistrosAfectados, adCmdText
    lSegundos = Timer - InicioTime
    lRegAfectados = mlRegistrosAfectados
    'DoEvents
    bExecute = True
    
    Exit Function
    
error_execute:
    sMensajeError = Err.Number & " - " & Err.Description
    GestionarError

End Function

'Ejecuta un Procedimiento Almacenado
'Se pasa el Nombre del PA, su numero de parametros
'y una lista de los parametros del PA
'Devuelve un Recordset con los valores de salida

Public Function arsExecutePA( _
    sNombre As String, _
    iNumParam As Integer, _
    ParamArray aParam() As Variant) As ADODB.Recordset
    
    Dim cmdPA As New ADODB.Command
    Dim i As Integer
    Dim iContadorFor As Integer

    On Error GoTo error_execute
    cmdPA.ActiveConnection = mCn
    cmdPA.CommandText = sNombre
    cmdPA.CommandType = adCmdStoredProc
    
    iContadorFor = iNumParam - 1
    For i = 0 To iContadorFor
        cmdPA.Parameters(i) = aParam(i)
    Next
    Set arsExecutePA = cmdPA.Execute()
    
    Exit Function
    
error_execute:
    Set arsExecutePA = Nothing
    GestionarError

End Function

Public Function arsExecuteCPA( _
    sNombre As String, _
    iNumParam As Integer, _
    ParamArray aParam() As Variant) As ADODB.Recordset
    
    Dim cmdPA As New ADODB.Command
    Dim prmPA As New ADODB.Parameter
    Dim i As Integer
    Dim iContadorFor As Integer

    On Error GoTo error_execute
    cmdPA.ActiveConnection = mCn
    cmdPA.CommandText = sNombre
    cmdPA.CommandType = adCmdStoredProc
    
    iContadorFor = iNumParam - 1
    For i = 0 To iContadorFor Step 2
        Set prmPA = cmdPA.CreateParameter(, aParam(i + 1))
        prmPA.Value = aParam(i)
        cmdPA.Parameters.Append prmPA
    Next
    Set arsExecuteCPA = cmdPA.Execute()
    
    Exit Function
    
error_execute:
    Set arsExecuteCPA = Nothing
    GestionarError

End Function
Public Function bExecutePA( _
    sNombre As String, _
    iNumParam As Integer, _
    ParamArray aParam() As Variant) As Boolean
    
    Dim cmdPA As New ADODB.Command
    Dim i As Integer
    Dim iContadorFor As Integer
    
    On Error GoTo error_executePA
    cmdPA.ActiveConnection = mCn
    cmdPA.CommandText = sNombre
    cmdPA.CommandType = adCmdStoredProc
    
    iContadorFor = iNumParam - 1
    For i = 0 To iContadorFor
        cmdPA.Parameters(i) = aParam(i)
    Next
    cmdPA.Execute
    bExecutePA = True
    
    Exit Function
    
error_executePA:
    bExecutePA = False
    GestionarError

End Function

Public Function HayErrores() As Boolean
    HayErrores = mCn.Errors.Count <> 0
End Function
'Muestra los errores si esta activada el modo de mostrar y si
'esta activado el modo de captura.
'Si el modo de captura esta desactivado vuelve a generar el error
Private Sub GestionarError()
    
    If mbModoCapturarErrores Then
        If mbModoMostrarErrores Then
            MostrarErrores
        End If
    Else
        Err.Raise "Err " & Err.Number, , Err.Description
    End If
    
End Sub
'Muestra los errores producidos en la conexión
Public Sub MostrarErrores()

    Dim aErr As ADODB.error
    Dim sMsg As String
    
    'If Not aErr Is Nothing Then
        For Each aErr In mCn.Errors
            Select Case aErr.NativeError
                Case -206
                    'Ignorar Error DropTable
                    Exit Sub
                Case 32 ' Error de bloqueo de registro
                    sMsg = "El Registro está bloqueado. Intente de nuevo la operación"
                Case Else
                    sMsg = "ERROR DE ACCESO A BASE DE DATOS." & vbCr & vbCr & _
                        aErr.Number & " - " & aErr.Description & vbCr & _
                        "Error Nativo: " & aErr.NativeError
            End Select
            MsgBox sMsg, vbCritical, App.Title & ". Error"
           ' Muestra_Mensaje sMsg
           ' End
        Next
    'End If

End Sub

Public Property Get CreadoConExito() As Boolean
    CreadoConExito = mbCreadoConExito
End Property

Public Property Get ModoMostrarErrores() As Boolean
    ModoMostrarErrores = mbModoMostrarErrores
End Property

Public Property Let ModoMostrarErrores(ByVal vNewValue As Boolean)
    mbModoMostrarErrores = vNewValue
End Property

Public Property Get ModoCapturarErrores() As Boolean
    ModoCapturarErrores = mbModoCapturarErrores
End Property

Public Property Let ModoCapturarErrores(ByVal vNewValue As Boolean)
    mbModoCapturarErrores = vNewValue
End Property

Public Property Get RegistrosAfectados() As Long
    RegistrosAfectados = mlRegistrosAfectados
End Property

Public Property Get cn() As ADODB.Connection
    Set cn = mCn
End Property
Public Sub EstablecerDirtyRead()
    
    Dim sSql As String
        
    sSql = "SET ISOLATION TO DIRTY READ"
    bExecute sSql

End Sub

Public Sub EstablecerLokcMode()
    
    Dim sSql As String
        
    sSql = "SET LOCK MODE TO WAIT"
    bExecute sSql

End Sub

Public Sub BeginTrans(bModoCapturaErrores As Boolean)

'Abre una transaccion y damos opcion a tratar los errores fuera de la clase BD

    ModoCapturarErrores = bModoCapturaErrores
    mCn.BeginTrans
    
End Sub

Public Sub CommitTrans()

'Cerramos transacción y volvemos a activar el trato de errores dentro de la clase

    mCn.CommitTrans
    ModoCapturarErrores = True
    
End Sub

Public Sub RollbackTrans()

'Deshacemos la transacción y volvemos a activar el trato de errores dentro de la clase

    mCn.RollbackTrans
    ModoCapturarErrores = True
    
End Sub

Public Function iValorGramo() As Integer

    Dim sSql As String
    Dim arsValor As ADODB.Recordset
    
    sSql = "SELECT valorgramo From refgeneral2"
    
    If bRecordset(sSql, arsValor) Then
        iValorGramo = arsValor(0)
    Else
        iValorGramo = 0
    End If
    
End Function

Public Function tFechaSistema() As Date

    Dim sSql As String
    Dim arsFecha As ADODB.Recordset
    
    sSql = "SELECT today From SysTables WHERE TabId = 1"
    
    If bRecordset(sSql, arsFecha) Then
        tFechaSistema = arsFecha(0)
    Else
        tFechaSistema = Now
    End If
    
End Function

Public Function idiaSemana() As Integer

    Dim sSql As String
    Dim arsFecha As ADODB.Recordset
    
    sSql = "SELECT weekday(today) From SysTables WHERE TabId = 1"
    
    If bRecordset(sSql, arsFecha) Then
        idiaSemana = arsFecha(0)
    Else
        idiaSemana = -1
    End If
    
End Function
Public Function tFechaHoraSistema() As Date

    Dim sSql As String
    Dim arsFecha As ADODB.Recordset
    
    sSql = "SELECT current year to second From SysTables WHERE TabId = 1"
    
    If bRecordset(sSql, arsFecha) Then
        tFechaHoraSistema = arsFecha(0)
    Else
        tFechaHoraSistema = Now
    End If

End Function

Public Function StaticRecordset( _
       QueryName As String, _
       ParamArray Param() As Variant _
     ) As ADODB.Recordset

    On Error GoTo ErrHandler

    Dim CmD As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim Parametro
    Dim prm As ADODB.Parameter
    
    Set CmD = New ADODB.Command
    Set rs = New ADODB.Recordset
    
    With CmD
        .ActiveConnection = mCn
        .CommandText = QueryName
        .CommandType = adCmdStoredProc
        
        For Each Parametro In Param
            Set prm = CmD.CreateParameter(, adDouble, adParamInput, , Parametro)
            CmD.Parameters.Append prm
            Set prm = Nothing
        Next
    End With

    With rs
        .CursorLocation = adUseClient
        .Open CmD, , adOpenStatic
        Set CmD.ActiveConnection = Nothing
        Set CmD = Nothing
        Set .ActiveConnection = Nothing
    End With

    Set StaticRecordset = rs
    Exit Function

ErrHandler:
    MsgBox Err.Description
    Stop
    Resume
    Set StaticRecordset = Nothing
    'Call ErrHandlerAction(Err)
End Function

Public Function dLookUp(ByVal sNombreTabla As String, ByVal sNombreCampo As String, Optional ByVal sCriterios As String = "", Optional ByVal ValorSiNull As Variant = Null) As Variant

Dim rsCursor As ADODB.Recordset, sSql As String

On Error GoTo ErrHandler

sSql = "SELECT " & sNombreCampo & " AS ASDFG123456 "
sSql = sSql & " FROM " & sNombreTabla
If sCriterios <> "" Then sSql = sSql & " WHERE " & sCriterios
If bRecordset(sSql, rsCursor) Then
    If Not IsNull(rsCursor("ASDFG123456")) Then
        dLookUp = Trim$(rsCursor("ASDFG123456"))
    Else
        dLookUp = ValorSiNull
    End If
Else
    dLookUp = ValorSiNull
End If
Set rsCursor = Nothing
Exit Function

ErrHandler:
    MsgBox Err.Description
    Stop
    Resume
    'Call ErrHandlerAction(Err)
End Function

Public Sub Actualiza_PP_OK(lID As Long, _
                sAlbaran As String, _
                sBulto As String, _
                iStatus As Integer, _
                iCantidad As Integer, _
                sSektion As String, _
                sUsuario As String)

    Dim cmdPrueba1 As ADODB.Command
    Dim cmdPrueba2 As ADODB.Command
    'Dim prmPrueba As ADODB.Parameter
    Set cmdPrueba1 = New ADODB.Command

    cmdPrueba1.CommandType = adCmdStoredProc

    Set cmdPrueba1.ActiveConnection = mCn

    cmdPrueba1.Parameters.Append cmdPrueba1.CreateParameter("@p_pp_ID", adDouble, adParamInput, 6)
    cmdPrueba1.Parameters.Append cmdPrueba1.CreateParameter("@p_pp_mengeist", adDouble, adParamInput, 6)
    cmdPrueba1.Parameters.Append cmdPrueba1.CreateParameter("@p_pp_status", adDouble, adParamInput, 6)
    cmdPrueba1.Parameters.Append cmdPrueba1.CreateParameter("@p_Macode", adVarChar, adParamInput, 10)
    cmdPrueba1.Parameters.Append cmdPrueba1.CreateParameter("@p_pp_lbnr", adVarChar, adParamOutput, 20)
    cmdPrueba1.Parameters.Append cmdPrueba1.CreateParameter("@p_pp_collinr", adDouble, adParamOutput, 6)
    cmdPrueba1.Parameters.Append cmdPrueba1.CreateParameter("@p_pp_SEKTION", adVarChar, adParamOutput, 6)
    cmdPrueba1.Parameters.Append cmdPrueba1.CreateParameter("@P_OUTTXT", adVarChar, adParamOutput, 30)
    cmdPrueba1.Parameters("@p_pp_ID").Value = Val(lID)
    cmdPrueba1.Parameters("@p_pp_MENGEIST").Value = Val(iCantidad)
    cmdPrueba1.Parameters("@p_pp_STATUS").Value = Val(iStatus)
    cmdPrueba1.Parameters("@p_MACODE").Value = sUsuario
    cmdPrueba1.Parameters("@p_pp_LBNR").Value = sAlbaran
    cmdPrueba1.Parameters("@p_pp_COLLINR").Value = Val(sBulto)
    cmdPrueba1.Parameters("@p_pp_SEKTION").Value = sSektion
    cmdPrueba1.Parameters("@p_OUTTXT").Value = ""
    cmdPrueba1.CommandText = "MT_PKG$pp_update"

    cmdPrueba1.CommandTimeout = 15
    cmdPrueba1.Execute
'''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''
    Set cmdPrueba2 = New ADODB.Command

    cmdPrueba2.CommandType = adCmdStoredProc

    Set cmdPrueba2.ActiveConnection = mCn
    
        cmdPrueba2.Parameters.Append cmdPrueba2.CreateParameter("@p_pp_LBNR", adVarChar, adParamInput, 20)
        cmdPrueba2.Parameters.Append cmdPrueba2.CreateParameter("@p_pp_COLLINR", adDouble, adParamInput, 6)
        cmdPrueba2.Parameters.Append cmdPrueba2.CreateParameter("@p_pp_SEKTION", adVarChar, adParamInput, 6)
        cmdPrueba2.Parameters.Append cmdPrueba2.CreateParameter("@p_lb_LHNR", adVarChar, adParamInput, 20)
        cmdPrueba2.Parameters.Append cmdPrueba2.CreateParameter("@p_OUTTXT", adVarChar, adParamOutput, 30)
        cmdPrueba2.Parameters("@p_pp_LBNR").Value = Trim(cmdPrueba1.Parameters("@p_pp_LBNR").Value)
        cmdPrueba2.Parameters("@p_pp_COLLINR").Value = Val(cmdPrueba1.Parameters("@p_pp_COLLINR").Value)
        cmdPrueba2.Parameters("@p_pp_SEKTION").Value = Trim(cmdPrueba1.Parameters("@p_pp_SEKTION").Value)
        cmdPrueba2.Parameters("@p_lb_LHNR").Value = sAlbaran & sBulto
        cmdPrueba2.CommandText = "MT_PKG$pp_Section_Fertig"

    cmdPrueba2.CommandTimeout = 15
    cmdPrueba2.Execute
    
End Sub





