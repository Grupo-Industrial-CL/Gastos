VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msIdUsuario As String
Private msContraseña As String
Private msNombre As String

Private miPaisTrabajo As Integer
Private miIdAplicacion As Integer

Private mcltOperaciones As Collection
Private mcltPaises As Collection

Private mbCreadoConExito As Boolean
Private mbConectadoConExito As Boolean

Private mrCnUsuario As ADODB.Connection

Public Sub Create( _
    sIdUsuario As String, _
    sContraseña As String, _
    iIdAplicacion As Integer, _
    cnUsuario As ADODB.Connection)
    
    Dim afoUsuario As ADODB.Recordset
    Dim afoOperUsuario As ADODB.Recordset
    Dim afoPaisesUsuario As ADODB.Recordset
    
    Set mrCnUsuario = cnUsuario
    
    mbCreadoConExito = False
    mbConectadoConExito = False
    
    If bConsUsuario(afoUsuario, _
                    sIdUsuario) Then
                    
        If Trim$(afoUsuario("Password")) = sContraseña Then
        
            msIdUsuario = sIdUsuario
            msContraseña = sContraseña
            msNombre = Trim$(afoUsuario("Nombre"))
            miIdAplicacion = iIdAplicacion
            miPaisTrabajo = afoUsuario("CodPais_Trabajo")
            
            If bConsOperacionesUsuario(afoOperUsuario, _
                                       msIdUsuario, _
                                       miIdAplicacion) Then
                
                Set mcltOperaciones = New Collection
                
                While Not afoOperUsuario.EOF
                  mcltOperaciones.Add _
                    Str$(afoOperUsuario("Id_Operacion"))
                  afoOperUsuario.MoveNext
                Wend
                
                If bConsPaisesUsuario(afoPaisesUsuario, msIdUsuario) Then
                    Set mcltPaises = New Collection
                    While Not afoPaisesUsuario.EOF
                      mcltPaises.Add _
                        Str$(afoPaisesUsuario("CodPais"))
                      afoPaisesUsuario.MoveNext
                    Wend
                    mbCreadoConExito = True
                Else
                    MsgBox "No se han podido recuperar los paises de actuación del usuario", vbExclamation, "Control de Usuario"
                End If
            Else
                MsgBox "Su usuario no tiene permisos en esta aplicación", vbExclamation, "Control de Usuario"
            End If
        Else
            MsgBox "La contraseña introducida no es correcta", vbExclamation, "Control de Usuario"
        End If
    Else
        MsgBox "El nombre de usuario es desconocido", vbExclamation, "Control de Usuario"
    End If
End Sub
    
Private Function bConsUsuario( _
    afoUsuario As ADODB.Recordset, _
    sIdUsuario As String) As Boolean
    
    Dim sSql As String
    
    sSql = "SELECT * FROM prf_usuarios" & _
            " WHERE Id_Usuario='" & sIdUsuario & "'"
            
 
    On Error GoTo errorConsUsuario
    
    Set afoUsuario = New ADODB.Recordset
    
    afoUsuario.Open sSql, mrCnUsuario, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If afoUsuario.EOF Then
        bConsUsuario = False
    Else
        bConsUsuario = True
    End If
        
Exit Function

errorConsUsuario:
    MsgBox "Error " & Err.Number & " al consultar datos de usuario", vbCritical, "Error de Base de Datos"
    bConsUsuario = False

End Function

Private Function bConsPaisesUsuario( _
    afoPaisesUsuario As ADODB.Recordset, _
    sIdUsuario As String) As Boolean

    Dim sSql As String
    
    sSql = "SELECT codpais FROM prf_RelUsuPaises" & _
            " WHERE id_Usuario='" & sIdUsuario & "'"

    On Error GoTo errorConsPaises
    
    Set afoPaisesUsuario = New ADODB.Recordset
    
    afoPaisesUsuario.Open sSql, mrCnUsuario, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If afoPaisesUsuario.EOF Then
        bConsPaisesUsuario = False
    Else
        bConsPaisesUsuario = True
    End If

Exit Function

errorConsPaises:
    MsgBox "Error " & Err.Number & " al consultar paises de actuación del usuario", vbCritical, "Error de Base de Datos"
    bConsPaisesUsuario = False
    
End Function


Private Function bConsOperacionesUsuario( _
    afoOperUsuario As ADODB.Recordset, _
    sIdUsuario As String, _
    iIdAplicacion As Integer) As Boolean
    
    Dim sSql As String
    
    sSql = "SELECT DISTINCT p.Id_Operacion FROM" & _
    " prf_RelUsuPerfiles u, prf_RelPerfilOper p" & _
    " WHERE u.Id_Perfil=p.Id_Perfil" & _
    " AND u.Id_Usuario='" & sIdUsuario & "'"
            
    On Error GoTo errorConsOpers
    
    Set afoOperUsuario = New ADODB.Recordset
    
    afoOperUsuario.Open sSql, mrCnUsuario, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If afoOperUsuario.EOF Then
        bConsOperacionesUsuario = False
    Else
        bConsOperacionesUsuario = True
    End If

Exit Function

errorConsOpers:
    MsgBox "Error " & Err.Number & " al consultar operaciones de usuario", vbCritical, "Error de Base de Datos"
    bConsOperacionesUsuario = False

End Function

Public Property Get CreadoConExito() As Boolean
    CreadoConExito = mbCreadoConExito
End Property

Public Property Get ConectadoConExito() As Boolean
    ConectadoConExito = mbConectadoConExito
End Property

Public Property Get IdUsuario() As String
    IdUsuario = msIdUsuario
End Property

Public Property Get IdContraseña() As String
    IdContraseña = msContraseña
End Property


Public Property Get Nombre() As String
    Nombre = msNombre
End Property

Public Property Get Permiso(iIdOperacion As Integer) As Boolean

    Dim vOperacion As Variant
    
    Permiso = False
    
    For Each vOperacion In mcltOperaciones
        If Val(vOperacion) = iIdOperacion Then
            Permiso = True
            Exit Property
        End If
    Next vOperacion
    
End Property

Public Property Get Operaciones() As Variant

    Set Operaciones = mcltOperaciones
    
End Property

Public Property Get PermisoPais(iPais As Integer) As Boolean

    Dim vPais As Variant
    
    PermisoPais = False
    
    For Each vPais In mcltPaises
        If Val(vPais) = iPais Then
            PermisoPais = True
            Exit Property
        End If
    Next vPais
    
End Property

Public Property Get PaisesAmbito() As Collection
    Set PaisesAmbito = mcltPaises
End Property

Public Property Get PaisTrabajo() As Integer
    PaisTrabajo = miPaisTrabajo
End Property


