Attribute VB_Name = "Rutinas"
Option Explicit

Dim sSql As String
Dim afoRegistro As ADODB.Recordset

Public Function bDameFechasCampanya(iCodPais As Integer, iAnyo As Integer, yCampanya As Byte, _
    lCodigo As Long, iJerarquia As Byte, sFecIni As String, sFecFin As String) As Boolean

    bDameFechasCampanya = False
    sFecIni = ""
    sFecFin = ""
    
    sSql = "SELECT FecInicio,FecFin " & _
           "FROM Campanyas " & _
           "WHERE CodPais = " & iCodPais & _
           " AND CodCampanya = " & yCampanya & _
           " AND Anyo = " & iAnyo
        
    Select Case iJerarquia
    Case 1
        sSql = sSql & _
              " AND CodDirVtas = (SELECT CodGeografia FROM Directores " & _
                                 "WHERE CodPais = " & iCodPais & _
                                 " AND Codigo = " & lCodigo & _
                                 " AND Jerarquia = 'DN')"
    
    Case 2
        sSql = sSql & _
              " AND CodRegion = (SELECT CodGeografia FROM Directores " & _
                                "WHERE CodPais = " & iCodPais & _
                                " AND Codigo = " & lCodigo & _
                                " AND Jerarquia = 'DR')"
    Case 3
        sSql = sSql & _
              " AND CodRegion = (SELECT CodRegion FROM Areas " & _
                                "WHERE CodPais = " & iCodPais & _
                                "AND CodArea = (SELECT MIN(Area) FROM DirApoyo " & _
                                               "WHERE CodPais = " & iCodPais & _
                                               " AND Codigo = " & lCodigo & "))"
    Case 4
        sSql = sSql & _
              " AND CodRegion = (SELECT CodRegion FROM Areas " & _
                                "WHERE CodPais = " & iCodPais & _
                                "AND CodArea = (SELECT CodGeografia FROM Directores " & _
                                               "WHERE CodPais = " & iCodPais & _
                                               " AND Codigo = " & lCodigo & _
                                               " AND Jerarquia = 'DA'))"
    
    Case 5
        sSql = sSql & _
              " AND CodRegion = (SELECT CodRegion FROM Supervisoras " & _
                                "WHERE CodPais = " & iCodPais & _
                                " AND CodSuperv = " & lCodigo & ")"
    Case Else
        Exit Function
    End Select
        
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        sFecIni = afoRegistro("FecInicio")
        sFecFin = afoRegistro("FecFin")
        bDameFechasCampanya = True
    End If
    Set afoRegistro = Nothing

End Function

'Public Function bDameDatosTarjetaSolred(iCodPais As Integer, sTarjeta As String, afoRegDatos As ADODB.Recordset) As Boolean
Public Function bDameDatosTarjetaSolred(sTarjeta As String, afoRegDatos As ADODB.Recordset) As Boolean

    sSql = sTarjeta
    If Mid(sTarjeta, 1, 7) = "9000000" Then sSql = "9000000000000000"
'14-05-18
'    sSql = "SELECT * FROM Tarjeta_Solred " & _
           "WHERE CodPais = " & iCodPais & " AND NumTarjeta = '" & sSql & "'"
    sSql = "SELECT * FROM Tarjeta_Solred " & _
           "WHERE NumTarjeta = '" & sSql & "'"
    bDameDatosTarjetaSolred = gBdGALILEO.bRecordset(sSql, afoRegDatos)
End Function

Public Function bDameDatosUsuarioEpsilon(sUsuarioEpsilon As String, iCodPais As Integer, ByRef afoRegDatos As ADODB.Recordset)

    sSql = "SELECT * FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " AND id_Epsilon = '" & sUsuarioEpsilon & "'"
    bDameDatosUsuarioEpsilon = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function

Public Function iDameCampanya(iCodPais As Integer, lCodigo As Long, sJerarquia As String, sFecha As String, Optional iAnyoCampanya As Integer = 0) As Integer

    Dim iCodRegion As Integer
    Dim iCampAux As Integer
    
    iDameCampanya = 0
    iAnyoCampanya = 0
    
    iCodRegion = iDameCodRegion(iCodPais, lCodigo, sJerarquia)
    
    sSql = "SELECT DISTINCT Anyo,CodCampanya FROM Campanyas WHERE CodPais = " & iCodPais & " AND CodRegion = " & iCodRegion & " AND FecInicio <= '" & sFecha & "' AND FecFin >= '" & sFecha & "'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        iCampAux = afoRegistro("CodCampanya")
        iAnyoCampanya = afoRegistro("Anyo")
    End If
    Set afoRegistro = Nothing
    
    If iCampAux = 0 Then
        
        sSql = "SELECT UNIQUE Anyo,CodCampanya FROM Campanyas WHERE CodPais = " & iCodPais & " AND FecInicio <= '" & sFecha & "' AND FecFin >= '" & sFecha & "'"
        If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
            iDameCampanya = afoRegistro("CodCampanya")
            iAnyoCampanya = afoRegistro("Anyo")
        End If
        Set afoRegistro = Nothing
        
    Else
        iDameCampanya = iCampAux
    End If

End Function

Public Function iDameCodRegion(iCodPais As Integer, lCodigo As Long, sJerarquia As String) As Integer

    iDameCodRegion = 0
    If Trim(sJerarquia) = "" Then Exit Function
    
    Select Case sJerarquia
    Case "DI"
        sSql = "SELECT CodRegion FROM Distribuidoras " & _
               "WHERE CodPais = " & iCodPais & _
               " AND CodDistrib = " & lCodigo
    Case "JG"
        sSql = "SELECT CodRegion FROM JefesGrupo " & _
               "WHERE CodPais = " & iCodPais & _
               " AND CodJefeGru = " & lCodigo
    Case "DC", "SU"
        sSql = "SELECT CodRegion FROM Supervisoras " & _
               "WHERE CodPais = " & iCodPais & _
               " AND CodSuperv = " & lCodigo
    Case "DA"
        sSql = "SELECT CodRegion FROM Areas " & _
               "WHERE CodPais = " & iCodPais & _
               " AND CodArea = (" & _
                    "SELECT CodGeografia FROM Directores " & _
                    "WHERE CodPais = " & iCodPais & _
                    " AND Codigo = " & lCodigo & _
                    " AND Jerarquia = 'DA')"
    Case "DR"
        sSql = "SELECT CodGeografia FROM Directores " & _
               "WHERE CodPais = " & iCodPais & _
               " AND Codigo = " & lCodigo & _
               " AND Jerarquia = 'DR'"
    
    Case "DP", "DS"
        sSql = "SELECT CodRegion FROM Areas " & _
               "WHERE CodPais = " & iCodPais & _
               " AND CodArea = (" & _
                    "SELECT MIN(Area) FROM DirApoyo " & _
                    "WHERE CodPais = " & iCodPais & _
                    " AND Codigo = " & lCodigo & ")"
    Case Else
        Exit Function
    End Select
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then iDameCodRegion = afoRegistro(0)
    Set afoRegistro = Nothing

End Function

Public Function bDameJefesDepartamento(iCodPais As Integer, iCodDepartamento As Integer, afoRegistro As ADODB.Recordset) As Boolean

    sSql = "SELECT * FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " " & _
           "AND Departamento = " & iCodDepartamento & " " & _
           "AND Jerarquia = 'JD' AND FechaBaja IS NULL " & _
           "ORDER BY Nombre"
    bDameJefesDepartamento = gBdGALILEO.bRecordset(sSql, afoRegistro)
    
End Function

Public Function bDameUsuariosDepartamento(iCodPais As Integer, iCodDepartamento As Integer, afoRegistro As ADODB.Recordset) As Boolean

    sSql = "SELECT * FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " " & _
           "AND Departamento = " & iCodDepartamento & _
           "ORDER BY LOWER(Nombre)"
    bDameUsuariosDepartamento = gBdGALILEO.bRecordset(sSql, afoRegistro)
    
End Function

Public Function bDameDatos_UsuarioPersonal_BDGastos(iCodPais As Integer, iCodDepartamento As Integer, iCodigo As Long, afoRegistro As ADODB.Recordset) As Boolean
    
    sSql = "SELECT * FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " " & _
           "AND Departamento = " & iCodDepartamento & " " & _
           "AND Codigo = " & iCodigo
    bDameDatos_UsuarioPersonal_BDGastos = gBDGASTOS.bRecordset(sSql, afoRegistro)

End Function

Public Function bInsertarJefeDepartamento(iCodPais As Integer, iCodDepartamento As Integer, sUsuario As String, sClave As String, sNombre As String, sDNI As String, seMail As String, ByRef sMensaje As String) As Boolean

    Dim iCodigo As Long
    
    On Error GoTo Error_Rutina
    
    bInsertarJefeDepartamento = False
    
    
    
    Exit Function



    sMensaje = ""
    
    sSql = "SELECT Codigo FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " " & _
           "AND Usuario = '" & sUsuario & "'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        sMensaje = "El usuario ya existe en el sistema."
        bInsertarJefeDepartamento = False
        Exit Function
    End If
    Set afoRegistro = Nothing
    
    sSql = "SELECT NVL(MAX(Codigo),0)+1 FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " " & _
           "AND Departamento = " & iCodDepartamento
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        sMensaje = "No se ha podido obtener código para el nuevo usuario."
        bInsertarJefeDepartamento = False
        Exit Function
    End If
    iCodigo = afoRegistro(0)
    Set afoRegistro = Nothing
    
    sSql = "INSERT INTO Usuarios_Personal (Usuario,Clave,CodPais,Jerarquia,Codigo,Nombre,DNI,Departamento,Cargo,Dir_eMail,FechaBaja,DependedeMarketing,ConCal,idPerfil,Evaluador," & _
           "    Usu_Vali_Vaca,Manager_GES,Bloqueado_Por,id_Ofi_Prevalidacion,Acreedor,Division_Personal,Clave_Organizacion) " & _
           "VALUES ('" & sUsuario & "','" & sClave & "'," & iCodPais & ",'JD'," & iCodigo & ",'" & sNombre & "','" & sDNI & "'," & iCodDepartamento & ",'Jefe Departamento'," & _
           "    '" & seMail & "',null,'N',null,null,null,null,null,'',null,null,null,null)"
    bInsertarJefeDepartamento = gBdGALILEO.bExecute(sSql)
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    sMensaje = Err.Number & " - " & Err.Description
    Resume Exit_Rutina
    Resume Next
End Function

Public Function bCambiar_DepartamentoUsuarioPersonal(iCodPais As Integer, iCodigo As Integer, iCodDepartamentoOrigen As Integer, iCodDepartamentoDestino As Integer, ByRef sMensaje As String) As Boolean

    Dim iCodigoAux As Long
    
    On Error GoTo Error_Rutina
    
    bCambiar_DepartamentoUsuarioPersonal = False
    
    sMensaje = ""
    
    'Comprobamos sí el código que tiene el usuario en el departamento destino está usado, si lo está debemos asignarle uno nuevo
    sSql = "SELECT Codigo FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " " & _
           "AND Departamento = " & iCodDepartamentoDestino & " " & _
           "AND Codigo = " & iCodigo
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
    'No existe el código en el departamento destino, pudemos seguir usando el mismo
        iCodigoAux = iCodigo
    Else
    'El código ya existe, por lo que tenemos que darle otro para el departamento destino
    
        sSql = "SELECT NVL(MAX(Codigo),0)+1 FROM Usuarios_Personal " & _
               "WHERE CodPais = " & iCodPais & " " & _
               "AND Departamento = " & iCodDepartamentoDestino
        If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
            sMensaje = "No se ha podido obtener código para el usuario en el departamento destino."
            bCambiar_DepartamentoUsuarioPersonal = False
            Exit Function
        End If
        iCodigoAux = afoRegistro(0)
        Set afoRegistro = Nothing
    
    End If
    
    'Modificar datos usuario
    gBdGALILEO.BeginTrans True
    sSql = "UPDATE Usuarios_Personal SET Departamento = " & iCodDepartamentoDestino & ",Codigo = " & iCodigoAux & " " & _
           "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamentoOrigen & " AND Codigo = " & iCodigo
    If Not gBdGALILEO.bExecute(sSql) Then
        gBdGALILEO.RollbackTrans
        bCambiar_DepartamentoUsuarioPersonal = False
        Exit Function
    End If
    'Cambimos sus gatos asociados
    sSql = "UPDATE Gastos_Internos SET Departamento = " & iCodDepartamentoDestino & ",Codigo = " & iCodigoAux & " " & _
           "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamentoOrigen & " AND Codigo = " & iCodigo
    If Not gBdGALILEO.bExecute(sSql) Then
        gBdGALILEO.RollbackTrans
        bCambiar_DepartamentoUsuarioPersonal = False
        Exit Function
    End If
    
    sSql = "UPDATE Usuarios_Personal SET Departamento = " & iCodDepartamentoDestino & ",Codigo = " & iCodigoAux & " " & _
           "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamentoOrigen & " AND Codigo = " & iCodigo
    If Not gBDGASTOS.bExecute(sSql) Then
        gBdGALILEO.RollbackTrans
        bCambiar_DepartamentoUsuarioPersonal = False
        Exit Function
    End If
    gBdGALILEO.CommitTrans
    bCambiar_DepartamentoUsuarioPersonal = True
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    sMensaje = Err.Number & " - " & Err.Description
    Resume Exit_Rutina
    Resume Next
End Function

Public Function bGuardar_UsuarioPersonal(iCodPais As Integer, iCodigo As Integer, iCodDepartamento As Integer, sUsuario As String, sJerarquia As String, sNombre As String, sDNI As String, sCargo As String, seMail As String, sCuenta_Bancaria As String, sid_SAP As String, sid_Epsilon As String, sCECO As String, ByRef sMensaje As String) As Boolean

    Dim iCodigoAux As Long
    Dim sPass As String
    Dim sDNI_E As String
    Dim sCuenta_Bancaria_E As String
    
    On Error GoTo Error_Rutina
    
    bGuardar_UsuarioPersonal = False
    
    sMensaje = ""
    
    sDNI_E = Encrypt_v2(sDNI)
    sPass = Encrypt(sDNI)
    sCuenta_Bancaria_E = Encrypt_v2(sCuenta_Bancaria)
    
    If iCodigo = 0 Then
    'Nuevo usuario
    
        sSql = "SELECT Codigo FROM Usuarios_Personal " & _
               "WHERE CodPais = " & iCodPais & " " & _
               "AND Usuario = '" & sUsuario & "'"
        If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
            sMensaje = "El usuario ya existe en el sistema."
            bGuardar_UsuarioPersonal = False
            Exit Function
        End If
        Set afoRegistro = Nothing
        
        sSql = "SELECT NVL(MAX(Codigo),0)+1 FROM Usuarios_Personal " & _
               "WHERE CodPais = " & iCodPais & " " & _
               "AND Departamento = " & iCodDepartamento
        If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
            sMensaje = "No se ha podido obtener código para el nuevo usuario."
            bGuardar_UsuarioPersonal = False
            Exit Function
        End If
        iCodigoAux = afoRegistro(0)
        Set afoRegistro = Nothing
        
        gBdGALILEO.BeginTrans True
        sSql = "INSERT INTO Usuarios_Personal (Usuario,Clave,CodPais,Jerarquia,Codigo,Nombre,Departamento,Cargo,Dir_eMail,FechaBaja,DependedeMarketing,ConCal,idPerfil,Evaluador," & _
                    "Usu_Vali_Vaca,Manager_GES,Bloqueado_Por,id_Ofi_Prevalidacion,Acreedor,Division_Personal,Clave_Organizacion,Pass,Primer_Acceso,DNI,id_SAP,id_Epsilon,CECO) " & _
               "VALUES ('" & sUsuario & "',''," & iCodPais & ",'" & sJerarquia & "'," & iCodigoAux & ",'" & sNombre & "'," & iCodDepartamento & ",'" & sCargo & "','" & _
               seMail & "',null,'N',null,null,null,null,null,'',null,null,null,null,'" & sPass & "','S','','" & sid_SAP & "','" & sid_Epsilon & "','" & sCECO & "')"
        If Not gBdGALILEO.bExecute(sSql) Then
            gBdGALILEO.RollbackTrans
            bGuardar_UsuarioPersonal = False
            Exit Function
        End If
        sSql = "INSERT INTO Usuarios_Personal (CodPais,Departamento,Codigo,DNI,Cuenta_Bancaria) " & _
               "VALUES (" & iCodPais & "," & iCodDepartamento & "," & iCodigoAux & ",'" & sDNI_E & "','" & sCuenta_Bancaria_E & "')"
        If Not gBDGASTOS.bExecute(sSql) Then
            gBdGALILEO.RollbackTrans
            bGuardar_UsuarioPersonal = False
            Exit Function
        End If
        gBdGALILEO.CommitTrans
        bGuardar_UsuarioPersonal = True
    
    Else
    'Modificar datos usuario
    
        gBdGALILEO.BeginTrans True
        sSql = "UPDATE Usuarios_Personal SET Jerarquia = '" & sJerarquia & "',Nombre = '" & sNombre & "',Cargo = '" & sCargo & "',Dir_eMail = '" & seMail & "'," & _
                                            "id_SAP = '" & sid_SAP & "',id_Epsilon = '" & sid_Epsilon & "',CECO = '" & sCECO & "' " & _
               "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " AND Codigo = " & iCodigo
        If Not gBdGALILEO.bExecute(sSql) Then
            gBdGALILEO.RollbackTrans
            bGuardar_UsuarioPersonal = False
            Exit Function
        End If
        'Por si cambia la jerarquia y que no se queden los gastos ya introducidos colgados, updateamos la tabla de gastos
        sSql = "UPDATE Gastos_Internos SET Jerarquia = '" & sJerarquia & "' " & _
               "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " AND Codigo = " & iCodigo
        If Not gBdGALILEO.bExecute(sSql) Then
            gBdGALILEO.RollbackTrans
            bGuardar_UsuarioPersonal = False
            Exit Function
        End If
        
        sSql = "UPDATE Usuarios_Personal SET DNI = '" & sDNI_E & "',Cuenta_Bancaria = '" & sCuenta_Bancaria_E & "' " & _
               "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " AND Codigo = " & iCodigo
        If Not gBDGASTOS.bExecute(sSql) Then
            gBdGALILEO.RollbackTrans
            bGuardar_UsuarioPersonal = False
            Exit Function
        End If
        gBdGALILEO.CommitTrans
        bGuardar_UsuarioPersonal = True
        
    End If
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    sMensaje = Err.Number & " - " & Err.Description
    Resume Exit_Rutina
    Resume Next
End Function

Public Function bRestablecerClave_UsuarioPersonal(iCodPais As Integer, iCodigo As Integer, iCodDepartamento As Integer, sDNI As String, ByRef sMensaje As String) As Boolean

    Dim sPass As String
    
    On Error GoTo Error_Rutina
    
    bRestablecerClave_UsuarioPersonal = False
    sMensaje = "El país no es válido"
    If iCodPais = 0 Then Exit Function
    sMensaje = "El usuario no es válido"
    If iCodigo = 0 Then Exit Function
    sMensaje = "El departamento no es válido"
    If iCodDepartamento = 0 Then Exit Function
    sMensaje = "El DNI no es válido"
    If Len(sDNI) < 8 Then Exit Function
    sMensaje = ""
    
    sPass = Encrypt(sDNI)
    
    gBdGALILEO.BeginTrans True
    sSql = "UPDATE Usuarios_Personal SET Pass='" & sPass & "',Primer_Acceso = 'S' " & _
           "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " AND Codigo = " & iCodigo
    If Not gBdGALILEO.bExecute(sSql) Then
        gBdGALILEO.RollbackTrans
        sMensaje = "Error al intentar realizar cambios en la BD"
        bRestablecerClave_UsuarioPersonal = False
        Exit Function
    End If
    gBdGALILEO.CommitTrans
    bRestablecerClave_UsuarioPersonal = True
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    sMensaje = Err.Number & " - " & Err.Description
    Resume Exit_Rutina
    Resume Next
End Function

Public Function bAlta_Baja_UsuariosPersonal(iCodPais As Integer, iCodDepartamento As Integer, iCodigo As Integer, iAccion As Integer) As Boolean

    On Error GoTo Error_Rutina
    
    sSql = "UPDATE Usuarios_Personal SET FechaBaja = "
    If iAccion = 0 Then
        sSql = sSql & " TODAY "
    Else
        sSql = sSql & " NULL "
    End If
    sSql = sSql & "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " AND Codigo = " & iCodigo
    bAlta_Baja_UsuariosPersonal = gBdGALILEO.bExecute(sSql)
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    Resume Exit_Rutina
    Resume Next
End Function


Public Function bInsertarTarjeta(iCodPais As Integer, iCodDepartamento As Integer, iCodigoUsuario As Long, sNombre As String, sJerarquia As String, sTarjeta As String, sMensaje As String) As Boolean

    Dim sEmpresa As String
    Dim sCIF As String

    On Error GoTo Error_Rutina

    sMensaje = ""
    
    sSql = "SELECT NumTarjeta FROM Tarjeta_Solred " & _
           "WHERE NumTarjeta = '" & sTarjeta & "'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        sMensaje = "La tarjeta ya existe en el sistema."
        bInsertarTarjeta = False
        Exit Function
    End If
    Set afoRegistro = Nothing
    
    sEmpresa = ""
    sCIF = ""
    sSql = "SELECT FIRST 1 Empresa,CIF FROM Tarjeta_Solred " & _
           "WHERE CodPais = " & iCodPais
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
       sEmpresa = Trim(afoRegistro("Empresa"))
       sCIF = Trim(afoRegistro("CIF"))
    End If
    Set afoRegistro = Nothing
    
    sSql = "INSERT INTO Tarjeta_Solred (CodPais,NumTarjeta,Titular,Empresa,CIF,Departamento,Codigo,Jerarquia,Ins_Gasto_Int) " & _
           "VALUES (" & iCodPais & ",'" & sTarjeta & "','" & sNombre & "','" & sEmpresa & "','" & sCIF & "'," & iCodDepartamento & "," & iCodigoUsuario & ",'" & sJerarquia & "','S')"
    bInsertarTarjeta = gBdGALILEO.bExecute(sSql)
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    sMensaje = Err.Number & " - " & Err.Description
    Resume Exit_Rutina
    Resume Next
End Function

Public Function bInsertarNuevoDepartamento(iCodPais As Integer, iCodDepartamento As Integer, sNombreDepartamento As String, iCECO As Long, sDivisa As String) As Boolean

    sSql = "INSERT INTO DepartamentosCL (Codigo,CodPais,Descripcion,CECO,Divisa,Activo) VALUES (" & iCodDepartamento & "," & iCodPais & ",'" & sNombreDepartamento & "'," & iCECO & ",'" & sDivisa & "','S')"
    bInsertarNuevoDepartamento = gBdGALILEO.bExecute(sSql)
    
End Function

Public Function bModificarDepartamento(iCodPais As Integer, iCodDepartamento As Integer, sNombreDepartamento As String, iCECO As Long) As Boolean
    sSql = "UPDATE DepartamentosCL SET Descripcion = '" & sNombreDepartamento & "',CECO = " & iCECO & " WHERE CodPais = " & iCodPais & " AND Codigo = " & iCodDepartamento
    bModificarDepartamento = gBdGALILEO.bExecute(sSql)
End Function

Public Function bActivarDesactivarDepartamento(iCodPais As Integer, iCodDepartamento As Integer, sActivo As String) As Boolean
    sSql = "UPDATE DepartamentosCL SET Activo = '" & sActivo & "' WHERE CodPais = " & iCodPais & " AND Codigo = " & iCodDepartamento
    bActivarDesactivarDepartamento = gBdGALILEO.bExecute(sSql)
End Function


Public Function bDameDepartamentosCL(iCodPais As Integer, afoRegDatos As ADODB.Recordset, id_Ofi_PreValidacion As Integer) As Boolean
'Con iCodPais = 0 el procedimiento devuelve los departamentos de todos los países
    
    'sSql = "EXECUTE PROCEDURE DepartamentosCL(" & iCodPais & "," & id_Ofi_PreValidacion & ")"
    sSql = "EXECUTE PROCEDURE DepartamentosCL_v2(" & iCodPais & "," & id_Ofi_PreValidacion & ")"
    bDameDepartamentosCL = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function

Public Function bDameEmpleadosDepartamento(iCodPais As Integer, iCodDepartamento As Integer, afoRegDatos As ADODB.Recordset, _
    Optional sListarBajas As String = "N", Optional id_Ofi_PreValidacion As Integer = 0, Optional sSoloPendientePreValidar As String = "N") As Boolean

    'sSql = "EXECUTE PROCEDURE UsuariosInternosDepartamento(" & iCodPais & "," & iCodDepartamento & ",'" & sListarBajas & "'," & id_Ofi_PreValidacion & ",'" & sSoloPendientePreValidar & "')"
    sSql = "EXECUTE PROCEDURE UsuariosInternosDepartamento_v2(" & iCodPais & "," & iCodDepartamento & ",'" & sListarBajas & "'," & id_Ofi_PreValidacion & ",'" & sSoloPendientePreValidar & "')"
    bDameEmpleadosDepartamento = gBdGALILEO.bRecordset(sSql, afoRegDatos)
    
End Function

Public Function bDameGastosPersonalInterno(iCodPais As Integer, iCodDepartamento As Integer, iCodEmpleado As Long, sJerEmpleado As String, _
    sFechaDesde As String, sFechaHasta As String, sValidado As String, sPagado As String, sContaSAP As String, id_Ofi_PreValidacion As Integer, _
    afoRegDatos As ADODB.Recordset) As Boolean

    '18-16-18 Piden desde GES poder sacar información filtrando por los diferentes estados de Validado, Pagado y Cantabilizado SAP
    Dim sFechaDesdeAux As String
    Dim sFechaHastaAux As String
    
    sFechaDesdeAux = Mid(sFechaDesde, 5) & Mid(sFechaDesde, 3, 2) & Mid(sFechaDesde, 1, 2)
    sFechaHastaAux = Mid(sFechaHasta, 5) & Mid(sFechaHasta, 3, 2) & Mid(sFechaHasta, 1, 2)
    
    'sSql = "EXECUTE PROCEDURE gastos_GastosUsuarioInternoPorFecha (" & iCodPais & "," & iCodDepartamento & "," & iCodEmpleado & ",'" & sJerEmpleado & "','" & sFechaDesde & "','" & sFechaHasta & "')"
    
    sSql = "EXECUTE PROCEDURE gastos_GastosUsuarioInternoPorFecha11 (" & iCodPais & "," & iCodDepartamento & "," & iCodEmpleado & ",'" & sJerEmpleado & "','" & _
           sFechaDesdeAux & "','" & sFechaHastaAux & "','" & sValidado & "','" & sPagado & "','" & sContaSAP & "'," & id_Ofi_PreValidacion & ")"
    bDameGastosPersonalInterno = gBdGALILEO.bRecordset(sSql, afoRegDatos)
    
End Function

Public Function bDameGastosPersonalInterno_v2(iCodPais As Integer, iCodDepartamento As Integer, iCodEmpleado As Long, sJerEmpleado As String, _
    sFechaDesde As String, sFechaHasta As String, sValidado As String, sPagado As String, sContaSAP As String, id_Ofi_PreValidacion As Integer, _
    afoRegDatos As ADODB.Recordset) As Boolean

    Dim sSqlValidado As String
    Dim sSqlPagado As String
    Dim sSqlContaSAP As String
    
    sSqlValidado = ""
    sSqlPagado = ""
    sSqlContaSAP = ""
        
    If sValidado <> "X" Then sSqlValidado = " AND NVL(a.Validado,'N') = '" & sValidado & "'"
    If sPagado <> "X" Then sSqlPagado = " AND NVL(a.Pagado,'N') = '" & sPagado & "'"
    If sContaSAP <> "X" Then sSqlContaSAP = " AND NVL(a.Conta_SAP,'N') = '" & sContaSAP & "'"
    
    sSql = "DROP TABLE IF EXISTS tGasUsu1"
    gBdGALILEO.bExecute sSql
    
    sSql = "SELECT a.Contador,a.TipoGasto,a.Tipo_Com_Inv,a.Fecha,a.Importe,a.Numero,a.Numero2,a.Donde,a.Visa,a.Validado,a.Pagado,a.Factura," & _
           "NVL(a.FamiliaGasto,'O') FamiliaGasto,a.Proveedor,a.Personas,a.Ruta_Ticket,NVL(b.Empresa,'') Empresa,NVL(Moneda,'') Moneda," & _
           "NVL(a.Cambio,1) Cambio,NVL(a.Importe_Ori,Importe) Importe_Ori,NVL(a.Pep_Ceco,'') Pep_Ceco,NVL(a.Archivado_En,'') Archivado_En," & _
           "NVL(a.Conta_SAP,'N') Conta_SAP,NVL(a.ModificadoAdm,'N') ModificadoAdm,NVL(a.CIF_Proveedor,'') CIF_Proveedor," & _
           "EXTEND(a.Fecha_Val_Rec,YEAR TO DAY) Fecha_Val_Rec,NVL(a.Mot_Rechazo,'') Mot_Rechazo," & _
           "NVL(a.Cuenta_Contable,'') Cuenta_Contable,NVL(c.Nombre,'') DescCuenta_Contable,EXTEND(Fecha_Insert,YEAR TO SECOND) Fecha_Insert," & _
           "NVL(a.Documento_SAP,'') Documento_SAP,NVL(a.Cambio_Guardado_Por_Usuario,'') Cambio_Guardado_Por_Usuario," & _
           "NVL(a.CPD_Nombre,'') CPD_Nombre,NVL(a.CPD_CP,'') CPD_CP,NVL(a.CPD_Ciudad,'') CPD_Ciudad,NVL(a.CPD_Pais,'') CPD_Pais,NVL(a.CPD_CIF,'') CPD_CIF," & _
           "a.Fecha_Conta_SAP Fecha_Conta_SAP,NVL(a.Desglose_IVA,'N') Desglose_IVA,NVL(a.Importe_IVA,0) Importe_IVA,NVL(a.Base_IVA,0) Base_IVA,NVL(a.Tipo_IVA,'') Tipo_IVA, " & _
           "NVL(a.Importe_IVA2,0) Importe_IVA2,NVL(a.Base_IVA2,0) Base_IVA2,NVL(a.Tipo_IVA2,'') Tipo_IVA2,NVL(a.PreValidado,'') PreValidado " & _
           "FROM Gastos_Internos a,OUTER Gastos_Empresas b,OUTER GES_CuentasContables c " & _
           "WHERE a.CodPais = " & iCodPais & " AND a.Departamento = " & iCodDepartamento & " AND a.Codigo = " & iCodEmpleado & " " & _
           "AND a.Jerarquia = '" & sJerEmpleado & "' AND a.Fecha BETWEEN '" & sFechaDesde & "' AND '" & sFechaHasta & "' " & _
           "AND NVL(a.id_Empresa,0) = b.id" & sSqlValidado & sSqlPagado & sSqlContaSAP & " " & _
           "AND NVL(a.Cuenta_Contable,'') = c.id " & _
           "INTO TEMP tGasUsu1 WITH NO LOG"
    If iCodPais = giPAIS_GES And id_Ofi_PreValidacion > 0 Then
        sSql = "SELECT a.Contador,a.TipoGasto,a.Tipo_Com_Inv,a.Fecha,a.Importe,a.Numero,a.Numero2,a.Donde,a.Visa,a.Validado,a.Pagado,a.Factura," & _
               "NVL(a.FamiliaGasto,'O') FamiliaGasto,a.Proveedor,a.Personas,a.Ruta_Ticket,NVL(b.Empresa,'') Empresa,NVL(Moneda,'') Moneda," & _
               "NVL(a.Cambio,1) Cambio,NVL(a.Importe_Ori,Importe) Importe_Ori,NVL(a.Pep_Ceco,'') Pep_Ceco,NVL(a.Archivado_En,'') Archivado_En," & _
               "NVL(a.Conta_SAP,'N') Conta_SAP,NVL(a.ModificadoAdm,'N') ModificadoAdm,NVL(a.CIF_Proveedor,'') CIF_Proveedor," & _
               "EXTEND(a.Fecha_Val_Rec,YEAR TO DAY) Fecha_Val_Rec,NVL(a.Mot_Rechazo,'') Mot_Rechazo," & _
               "NVL(a.Cuenta_Contable,'') Cuenta_Contable,NVL(c.Nombre,'') DescCuenta_Contable,EXTEND(Fecha_Insert,YEAR TO SECOND) Fecha_Insert," & _
               "NVL(a.Documento_SAP,'') Documento_SAP,NVL(a.Cambio_Guardado_Por_Usuario,'') Cambio_Guardado_Por_Usuario," & _
               "NVL(a.CPD_Nombre,'') CPD_Nombre,NVL(a.CPD_CP,'') CPD_CP,NVL(a.CPD_Ciudad,'') CPD_Ciudad,NVL(a.CPD_Pais,'') CPD_Pais,NVL(a.CPD_CIF,'') CPD_CIF," & _
               "a.Fecha_Conta_SAP Fecha_Conta_SAP,NVL(a.Desglose_IVA,'N') Desglose_IVA,NVL(a.Importe_IVA,0) Importe_IVA,NVL(a.Base_IVA,0) Base_IVA,NVL(a.Tipo_IVA,'') Tipo_IVA, " & _
               "NVL(a.Importe_IVA2,0) Importe_IVA2,NVL(a.Base_IVA2,0) Base_IVA2,NVL(a.Tipo_IVA2,'') Tipo_IVA2,NVL(a.PreValidado,'') PreValidado " & _
               "FROM Gastos_Internos a,Usuarios_Personal d,OUTER Gastos_Empresas b,OUTER GES_CuentasContables c " & _
               "WHERE a.CodPais = " & iCodPais & " AND a.Departamento = " & iCodDepartamento & " AND a.Codigo = " & iCodEmpleado & " " & _
               "AND a.Jerarquia = '" & sJerEmpleado & "' AND a.Fecha BETWEEN '" & sFechaDesde & "' AND '" & sFechaHasta & "' " & _
               "AND NVL(a.id_Empresa,0) = b.id" & sSqlValidado & sSqlPagado & sSqlContaSAP & " " & _
               "AND NVL(a.Cuenta_Contable,'') = c.id " & _
               "AND a.CodPais = d.CodPais AND a.Departamento = d.Departamento AND a.Codigo = d.Codigo AND a.Jerarquia = d.Jerarquia " & _
               "AND (NVL(a.Pep_Ceco,'') IN (SELECT DISTINCT PEPCECO " & _
                                           "FROM GES_PEPCECO " & _
                                           "WHERE Responsable IN (" & _
                                                "SELECT DISTINCT Codigo FROM Usuarios_Personal WHERE CodPais = " & iCodPais & " AND Usuario IN (" & _
                                                    "SELECT Usu_Responsable FROM Gastos_Relacion_PreValidacion WHERE CodPais = " & iCodPais & " AND id_Ofi_PreValidacion = " & id_Ofi_PreValidacion & "))) " & _
                                 "OR " & _
                                 "NVL(d.Manager_GES,0) IN (SELECT DISTINCT Codigo FROM Usuarios_Personal WHERE CodPais = " & iCodPais & " AND Usuario IN (" & _
                                                                  "SELECT Usu_Responsable FROM Gastos_Relacion_PreValidacion WHERE CodPais = " & iCodPais & " AND id_Ofi_PreValidacion = " & id_Ofi_PreValidacion & "))) " & _
               "INTO TEMP tGasUsu1 WITH NO LOG"
    End If
    If Not gBdGALILEO.bExecute(sSql) Then
        bDameGastosPersonalInterno_v2 = False
        Exit Function
    End If
    
    sSql = "SELECT a.Contador,a.TipoGasto,a.Tipo_Com_Inv,a.Fecha,a.Importe,a.Numero,a.Numero2,a.Donde,a.Visa,a.Validado,a.Pagado,a.Factura," & _
           "a.FamiliaGasto,a.Proveedor,a.Personas,a.Ruta_Ticket,a.Empresa,a.Moneda,a.Cambio,a.Importe_Ori,a.Pep_Ceco,a.Archivado_En," & _
           "a.Conta_SAP,a.ModificadoAdm,a.CIF_Proveedor,a.Fecha_Val_Rec,a.Mot_Rechazo Motivo_Rechazo,a.Cuenta_Contable,a.DescCuenta_Contable Desc_CuentaContable," & _
           "a.Fecha_Insert,a.Documento_SAP,a.Cambio_Guardado_Por_Usuario,a.CPD_Nombre,a.CPD_CP,a.CPD_Ciudad,a.CPD_Pais,a.CPD_CIF,a.Fecha_Conta_SAP," & _
           "a.Desglose_IVA , a.Importe_IVA, a.Base_IVA, a.Tipo_IVA, a.Importe_IVA2, a.Base_IVA2, a.Tipo_IVA2, a.PreValidado " & _
           "FROM tGasUsu1 a " & _
           "ORDER BY a.Fecha,a.Contador"
    bDameGastosPersonalInterno_v2 = gBdGALILEO.bRecordset(sSql, afoRegDatos)
    
End Function

Public Function bDameGastosPersonalInterno_v3(iCodPais As Integer, iCodDepartamento As Integer, iCodEmpleado As Long, sJerEmpleado As String, _
    sFechaDesde As String, sFechaHasta As String, sValidado As String, sPagado As String, sContaSAP As String, id_Ofi_PreValidacion As Integer, _
    afoRegDatos As ADODB.Recordset) As Boolean

    Dim sSqlValidado As String
    Dim sSqlPagado As String
    Dim sSqlContaSAP As String
    
    sSqlValidado = ""
    sSqlPagado = ""
    sSqlContaSAP = ""
        
    If sValidado <> "X" Then sSqlValidado = " AND NVL(a.Validado,'N') = '" & sValidado & "'"
    If sPagado <> "X" Then sSqlPagado = " AND NVL(a.Pagado,'N') = '" & sPagado & "'"
    If sContaSAP <> "X" Then sSqlContaSAP = " AND NVL(a.Conta_SAP,'N') = '" & sContaSAP & "'"
    
    sSql = "DROP TABLE IF EXISTS tGasUsu1"
    gBdGALILEO.bExecute sSql
    
    sSql = "SELECT a.Contador,a.TipoGasto,a.Tipo_Com_Inv,a.Fecha,a.Importe,a.Numero,a.Numero2,a.Donde,a.Visa,a.Validado,a.Pagado,a.Factura," & _
           "NVL(a.FamiliaGasto,'O') FamiliaGasto,a.Proveedor,a.Personas,a.Ruta_Ticket,NVL(b.Empresa,'') Empresa,NVL(Moneda,'') Moneda," & _
           "NVL(a.Cambio,1) Cambio,NVL(a.Importe_Ori,Importe) Importe_Ori,NVL(a.Pep_Ceco,'') Pep_Ceco,NVL(a.Archivado_En,'') Archivado_En," & _
           "NVL(a.Conta_SAP,'N') Conta_SAP,NVL(a.ModificadoAdm,'N') ModificadoAdm,NVL(a.CIF_Proveedor,'') CIF_Proveedor," & _
           "EXTEND(a.Fecha_Val_Rec,YEAR TO DAY) Fecha_Val_Rec,NVL(a.Mot_Rechazo,'') Mot_Rechazo," & _
           "NVL(a.Cuenta_Contable,'') Cuenta_Contable,NVL(c.Nombre,'') DescCuenta_Contable,EXTEND(Fecha_Insert,YEAR TO SECOND) Fecha_Insert," & _
           "NVL(a.Documento_SAP,'') Documento_SAP,NVL(a.Cambio_Guardado_Por_Usuario,'') Cambio_Guardado_Por_Usuario," & _
           "NVL(a.CPD_Nombre,'') CPD_Nombre,NVL(a.CPD_CP,'') CPD_CP,NVL(a.CPD_Ciudad,'') CPD_Ciudad,NVL(a.CPD_Pais,'') CPD_Pais,NVL(a.CPD_CIF,'') CPD_CIF," & _
           "a.Fecha_Conta_SAP Fecha_Conta_SAP,NVL(a.Desglose_IVA,'N') Desglose_IVA,NVL(a.Importe_IVA,0) Importe_IVA,NVL(a.Base_IVA,0) Base_IVA,NVL(a.Tipo_IVA,'') Tipo_IVA," & _
           "NVL(a.Importe_IVA2,0) Importe_IVA2,NVL(a.Base_IVA2,0) Base_IVA2,NVL(a.Tipo_IVA2,'') Tipo_IVA2,NVL(a.PreValidado,'') PreValidado,NVL(a.Validador,'') Validador," & _
           "NVL(a.Contador_Asociado,0) Contador_Asociado,NVL(a.PreValidador,'') PreValidador " & _
           "FROM Gastos_Internos a,OUTER Gastos_Empresas b,OUTER GES_CuentasContables c " & _
           "WHERE a.CodPais = " & iCodPais & " AND a.Departamento = " & iCodDepartamento & " AND a.Codigo = " & iCodEmpleado & " " & _
           "AND a.Jerarquia = '" & sJerEmpleado & "' AND a.Fecha BETWEEN '" & sFechaDesde & "' AND '" & sFechaHasta & "' " & _
           "AND NVL(a.id_Empresa,0) = b.id" & sSqlValidado & sSqlPagado & sSqlContaSAP & " " & _
           "AND NVL(a.Cuenta_Contable,'') = c.id "
    If iCodPais = giPAIS_GES And id_Ofi_PreValidacion > 0 Then sSql = sSql & " AND NVL(a.Ofi_PreValidacion,0) = " & id_Ofi_PreValidacion & " "
    sSql = sSql & "INTO TEMP tGasUsu1 WITH NO LOG"
    If Not gBdGALILEO.bExecute(sSql) Then
        bDameGastosPersonalInterno_v3 = False
        Exit Function
    End If
    
    sSql = "SELECT a.Contador,a.TipoGasto,a.Tipo_Com_Inv,a.Fecha,a.Importe,a.Numero,a.Numero2,a.Donde,a.Visa,a.Validado,a.Pagado,a.Factura," & _
           "a.FamiliaGasto,a.Proveedor,a.Personas,a.Ruta_Ticket,a.Empresa,a.Moneda,a.Cambio,a.Importe_Ori,a.Pep_Ceco,a.Archivado_En," & _
           "a.Conta_SAP,a.ModificadoAdm,a.CIF_Proveedor,a.Fecha_Val_Rec,a.Mot_Rechazo Motivo_Rechazo,a.Cuenta_Contable,a.DescCuenta_Contable Desc_CuentaContable," & _
           "a.Fecha_Insert,a.Documento_SAP,a.Cambio_Guardado_Por_Usuario,a.CPD_Nombre,a.CPD_CP,a.CPD_Ciudad,a.CPD_Pais,a.CPD_CIF,a.Fecha_Conta_SAP," & _
           "a.Desglose_IVA , a.Importe_IVA, a.Base_IVA, a.Tipo_IVA, a.Importe_IVA2, a.Base_IVA2, a.Tipo_IVA2, a.PreValidado,a.Validador,a.Contador_Asociado,a.PreValidador " & _
           "FROM tGasUsu1 a " & _
           "ORDER BY a.Fecha,a.Contador"
    bDameGastosPersonalInterno_v3 = gBdGALILEO.bRecordset(sSql, afoRegDatos)
    
End Function

Public Function sNombreGasto(sCategoriaGasto As String, sTipoGasto As String) As String
    
    sSql = "EXECUTE PROCEDURE gastos_NombreGasto('" & sCategoriaGasto & "','" & sTipoGasto & "')"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then sNombreGasto = Trim(NoNull(afoRegistro("Nombre"), "T"))
    
End Function

Public Function sidGasto(sCategoriaGasto As String, sTipoGasto As String) As String

    sSql = "EXECUTE PROCEDURE gastos_idGasto('" & sCategoriaGasto & "','" & sTipoGasto & "')"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then sidGasto = Trim(NoNull(afoRegistro("idGasto"), "T"))

End Function

Public Function bEliminarGastoInterno(sContador As String, sJerarquia As String) As Boolean

    bEliminarGastoInterno = False

    sSql = "EXECUTE PROCEDURE gv_borrargasto (" & sContador & ",'" & sJerarquia & "')"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then If afoRegistro(0) = 0 Then bEliminarGastoInterno = True
    Set afoRegistro = Nothing
    
End Function

Public Function bEliminarGastoInterno_GES(sContador As String) As Boolean

    bEliminarGastoInterno_GES = False
    
    sSql = "UPDATE Gastos_Internos SET CodPais = CodPais * -1 WHERE Contador = " & sContador
    bEliminarGastoInterno_GES = gBdGALILEO.bExecute(sSql)
    
End Function

Public Function sDameJerarquiaEmpleado(iCodPais As Integer, iCodDepartamento As Integer, iCodEmpleado As Long) As String

    sSql = "EXECUTE PROCEDURE JerarquiaUsuarioInterno(" & iCodPais & "," & iCodDepartamento & "," & iCodEmpleado & ")"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then sDameJerarquiaEmpleado = Trim(NoNull(afoRegistro(0), "T"))
    Set afoRegistro = Nothing
    
End Function

Public Function sDameAcreedorEmpleado(iCodPais As Integer, iCodDepartamento As Integer, lCodigo As Long, sJerarquia As String) As String

    sDameAcreedorEmpleado = ""
    sSql = "SELECT NVL(Acreedor,'') FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " AND Codigo = " & lCodigo & " AND Jerarquia = '" & sJerarquia & "' AND NVL(Acreedor,'') <> ''"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then sDameAcreedorEmpleado = Trim(afoRegistro(0))
    Set afoRegistro = Nothing
    
End Function

Public Sub DameDatosAuxEmpleado(iCodPais As Integer, iCodDepartamento As Integer, lCodigo As Long, sJerarquia As String, _
                                      sAcreedor As String, sDivision_Personal As String, sClave_Organizacion As String, _
                                      iManager As Long, iOfi_PreValidacion As Integer)
    
    sAcreedor = ""
    sDivision_Personal = ""
    sClave_Organizacion = ""
    iManager = -1
    iOfi_PreValidacion = -1
    
    sSql = "SELECT NVL(Acreedor,'') Acreedor,NVL(Division_Personal,'') Division_Personal,NVL(Clave_Organizacion,'') Clave_Organizacion,NVL(Manager_GES,-1) Manager " & _
           "FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " AND Codigo = " & lCodigo & " AND Jerarquia = '" & sJerarquia & "'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        sAcreedor = Trim(afoRegistro("Acreedor"))
        sDivision_Personal = Trim(afoRegistro("Division_Personal"))
        sClave_Organizacion = Trim(afoRegistro("Clave_Organizacion"))
        iManager = afoRegistro("Manager")
    End If
    Set afoRegistro = Nothing
    
    sSql = "SELECT id_Ofi_PreValidacion FROM Gastos_Relacion_PreValidacionV2 " & _
           "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " " & _
           "AND Division_Personal = '" & sDivision_Personal & "' AND Clave_Organizacion = '" & sClave_Organizacion & "'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then iOfi_PreValidacion = afoRegistro("id_Ofi_PreValidacion")
    Set afoRegistro = Nothing
    
End Sub


Public Function bMarcarGastosInternosValidados(sContadores As String, sTextoValidacion As String) As Boolean

    sSql = "EXECUTE PROCEDURE gastos_MarcarGastosInternosValidados(" & sContadores & ",'" & sTextoValidacion & "')"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then If afoRegistro(0) = 0 Then bMarcarGastosInternosValidados = True
    Set afoRegistro = Nothing
    
End Function

Public Function bMarcarGastosInternosPreValidados(sContadores As String, sUsuarioPreValidacion As String) As Boolean

    sSql = "EXECUTE PROCEDURE gastos_MarcarGastosInternosPreValidados(" & sContadores & ",'" & sUsuarioPreValidacion & "')"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then If afoRegistro(0) = 0 Then bMarcarGastosInternosPreValidados = True
    Set afoRegistro = Nothing

End Function

Public Function bMarcarGastosInternosPagados(sContadores As String) As Boolean

    sSql = "EXECUTE PROCEDURE gastos_MarcarGastosInternosPagados(" & sContadores & ")"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then If afoRegistro(0) = 0 Then bMarcarGastosInternosPagados = True
    Set afoRegistro = Nothing
    
End Function

Public Sub VaciarBox(Control As Control)
    
    Dim iCont As Integer
    
    For iCont = 0 To Control.ListCount - 1
        Control.RemoveItem 0
    Next iCont
    
    Control.Clear

End Sub

Public Function bEsCodigoDS(iCodPais As Integer, lCodigo As Long) As Boolean

    sSql = "SELECT Codigo FROM Directores WHERE CodPais = " & iCodPais & " AND Codigo = " & lCodigo & " AND Jerarquia = 'DS'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then bEsCodigoDS = True
    Set afoRegistro = Nothing
    
End Function

Public Function bDameEmpresasGasto(afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT id id_Empresa,Empresa FROM Gastos_Empresas ORDER BY Orden"
    bDameEmpresasGasto = gBdGALILEO.bRecordset(sSql, afoRegDatos)
    
End Function

Public Function sDameTipoGastoPais(iCodPais As Integer) As String

    sDameTipoGastoPais = "I"
    sSql = "SELECT id FROM Gastos_Empresas " & _
           "WHERE CodPais = " & iCodPais
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then sDameTipoGastoPais = afoRegistro("id")
    Set afoRegistro = Nothing

End Function

Public Function bDameTiposGasto(sTipo As String, afoRegDatos As ADODB.Recordset)

    sSql = "SELECT * FROM TiposGasto " & _
           "WHERE Tipo = '" & sTipo & "' " & _
           "ORDER BY idGasto"
    bDameTiposGasto = gBdGALILEO.bRecordset(sSql, afoRegDatos)
           
End Function

Public Function bDameDivisas(afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT Valor FROM TiposGasto_Combos WHERE Tipo = '7' AND idDetalleGasto = 34 ORDER BY id_DG_Combo"
    bDameDivisas = gBdGALILEO.bRecordset(sSql, afoRegDatos)
    
End Function
    
Public Function dDameCambioDivisa(iCodPais As Integer, sDivisa As String, sFecha As String, sDivisa_Departamento As String) As Double

    Dim iCodPais_Divisa As Integer
    Dim sDivisaAux As String
    
    dDameCambioDivisa = 1
    
    If iCodPais = giPAIS_GES Then
    
        sSql = "SELECT Cambio " & _
               "FROM Gastos_Cambios_Divisa_SAP " & _
               "WHERE Fecha_Validez = '" & sFecha & "' " & _
               "AND Divisa_De = '" & sDivisa & "' AND Divisa_A = '" & sDivisa_Departamento & "'"
    
    Else
    
        sDivisaAux = ""
        Select Case sDivisa
        Case "USD"
            iCodPais_Divisa = 7
        Case "GBP"
            iCodPais_Divisa = 44
        Case "MXN"
            iCodPais_Divisa = 52
        Case "CLP"
            iCodPais_Divisa = 56
        Case "PAB"
            iCodPais_Divisa = 7
        Case "CRC"
            iCodPais_Divisa = 6
        Case "MAD"
            iCodPais_Divisa = 12
        Case "TND"
            iCodPais_Divisa = 16
        Case Else
            sDivisaAux = sDivisa
        End Select
        
        sSql = Replace(sFecha, "/", "")
        If sDivisaAux = "" Then
            sSql = "SELECT a.Cambio " & _
                   "FROM contabilidad@clay_dw:he_Cambios a,Calendario b " & _
                   "WHERE a.Pais_id = " & iCodPais_Divisa & " AND b.Fecha = '" & sSql & "' " & _
                   "AND a.Anyo_id = YEAR(b.Fecha) AND a.Mes_id = (YEAR(b.Fecha)*100)+MONTH(b.Fecha)"
        Else
            sSql = "SELECT a.Cambio " & _
                   "FROM Gastos_Cambios_Divisa a,Calendario b " & _
                   "WHERE a.Divisa = '" & sDivisa & "' AND b.Fecha = '" & sSql & "' " & _
                   "AND YEAR(a.Fecha) = YEAR(b.Fecha) AND MONTH(a.Fecha) = MONTH(b.Fecha)"
        End If
        
    End If
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then dDameCambioDivisa = afoRegistro(0)
    Set afoRegistro = Nothing
        
End Function

Public Function bDameCuentasContablesGES(afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT id Cuenta_Contable,Nombre DescCuenta_Contable FROM GES_CuentasContables " & _
           "ORDER BY id"
    bDameCuentasContablesGES = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function

Public Function bComprobarPEPCECO(sPEPCECO As String, sDepartamento As String, sFecha As String, lResponsable As Long, lDirectorProyecto As Long, sPreValidador As String) As Boolean
    
    bComprobarPEPCECO = False
    lResponsable = 0
    lDirectorProyecto = 0
    sPreValidador = ""
    
    sSql = Replace(sFecha, "/", "")
    sSql = "SELECT PEPCECO,NVL(Responsable,0) Responsable,NVL(Director_Proyecto,0) Director_Proyecto,NVL(PreValidador,'') PreValidador " & _
           "FROM GES_PEPCECO " & _
           "WHERE Sociedad = " & sDepartamento & " AND '" & sSql & "' BETWEEN Fecha_Aceptacion AND Fecha_Cierre " & _
           "AND LOWER(PEPCECO) = '" & LCase(sPEPCECO) & "'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        bComprobarPEPCECO = True
        lResponsable = afoRegistro("Responsable")
        lDirectorProyecto = afoRegistro("Director_Proyecto")
        sPreValidador = afoRegistro("PreValidador")
        If sPreValidador = "0" Then sPreValidador = ""
    End If
    Set afoRegistro = Nothing
    
End Function

Public Function bDameDatosHistoricosGastoInterno(lContador As Long, afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT a.*,b.Empresa " & _
           "FROM Log_Gastos_Internos a,Gastos_Empresas b " & _
           "WHERE a.Contador = " & lContador & " AND a.id_Empresa = b.id " & _
           "ORDER BY a.Fecha_Modi"
    bDameDatosHistoricosGastoInterno = gBdGALILEO.bRecordset(sSql, afoRegDatos)
    
End Function

Public Function sDameRutaEquipo(iCodPais As Integer) As String

    sDameRutaEquipo = ""
    
    If iCodPais = giPAIS_ESPAÑA Then
        sDameRutaEquipo = "C:\"
    Else
        sDameRutaEquipo = "V:\"
    End If

End Function

Public Function bDameEmpleados_GastosValidados_NoPagados(iCodPais As Integer, afoRegDatos As ADODB.Recordset, sTipoBusqueda As String) As Boolean

    If sTipoBusqueda = "gastos" Then sSql = "AND a.TipoGasto NOT IN (99) "
    If sTipoBusqueda = "anticipos" Then sSql = "AND a.TipoGasto IN (99) "

    sSql = "SELECT a.Codigo,c.Nombre Nom_Empleado,a.Departamento,b.Descripcion Nom_Departamento,CASE WHEN c.FechaBaja IS NULL THEN 'N' ELSE 'S' END EsBaja,MIN(Fecha) Fecha_Ini,MIN(Fecha_Insert) MinFechaInsert " & _
           "FROM Gastos_Internos a,DepartamentosCL b,Usuarios_Personal c " & _
           "WHERE a.CodPais = " & iCodPais & " AND a.CodPais = b.CodPais AND a.Departamento = b.Codigo " & _
           "AND a.CodPais = c.CodPais AND a.Codigo = c.Codigo AND a.Jerarquia = c.Jerarquia AND a.Departamento = c.Departamento " & _
           "AND NVL(a.Validado,'N') = 'S' AND NVL(a.Pagado,'N') = 'N' " & sSql & _
           "GROUP BY a.Codigo,c.Nombre,a.Departamento,b.Descripcion,c.FechaBaja " & _
           "ORDER BY b.Descripcion,c.Nombre"
    bDameEmpleados_GastosValidados_NoPagados = gBdGALILEO.bRecordset(sSql, afoRegDatos)
           
End Function

Public Function bDameEmpleados_GastosValidados_NoContabilizados(iCodPais As Integer, afoRegDatos As ADODB.Recordset, sTipoBusqueda As String) As Boolean

    If sTipoBusqueda = "gastos" Then sSql = "AND a.TipoGasto NOT IN (99) "
    If sTipoBusqueda = "anticipos" Then sSql = "AND a.TipoGasto IN (99) "

    sSql = "SELECT a.Codigo,c.Nombre Nom_Empleado,a.Departamento,b.Descripcion Nom_Departamento,CASE WHEN c.FechaBaja IS NULL THEN 'N' ELSE 'S' END EsBaja,MIN(Fecha) Fecha_Ini,MIN(Fecha_Insert) MinFechaInsert " & _
           "FROM Gastos_Internos a,DepartamentosCL b,Usuarios_Personal c " & _
           "WHERE a.CodPais = " & iCodPais & " AND a.CodPais = b.CodPais AND a.Departamento = b.Codigo " & _
           "AND a.CodPais = c.CodPais AND a.Codigo = c.Codigo AND a.Jerarquia = c.Jerarquia AND a.Departamento = c.Departamento " & _
           "AND NVL(a.Validado,'N') = 'S' AND NVL(a.Conta_SAP,'N') = 'N' " & sSql & _
           "GROUP BY a.Codigo,c.Nombre,a.Departamento,b.Descripcion,c.FechaBaja " & _
           "ORDER BY b.Descripcion,c.Nombre"
    bDameEmpleados_GastosValidados_NoContabilizados = gBdGALILEO.bRecordset(sSql, afoRegDatos)
           
End Function

Public Function bDameEmpleados_GastosNoValidados(iCodPais As Integer, afoRegDatos As ADODB.Recordset, sTipoBusqueda As String) As Boolean

    If sTipoBusqueda = "gastos" Then sSql = "AND a.TipoGasto NOT IN (99) "
    If sTipoBusqueda = "anticipos" Then sSql = "AND a.TipoGasto IN (99) "

    sSql = "SELECT a.Codigo,c.Nombre Nom_Empleado,a.Departamento,b.Descripcion Nom_Departamento,CASE WHEN c.FechaBaja IS NULL THEN 'N' ELSE 'S' END EsBaja,MIN(Fecha) Fecha_Ini,MIN(Fecha_Insert) MinFechaInsert " & _
           "FROM Gastos_Internos a,DepartamentosCL b,Usuarios_Personal c " & _
           "WHERE a.CodPais = " & iCodPais & " AND a.CodPais = b.CodPais AND a.Departamento = b.Codigo " & _
           "AND a.CodPais = c.CodPais AND a.Codigo = c.Codigo AND a.Jerarquia = c.Jerarquia AND a.Departamento = c.Departamento " & _
           "AND NVL(a.Validado,'N') = 'N' " & sSql & _
           "GROUP BY a.Codigo,c.Nombre,a.Departamento,b.Descripcion,c.FechaBaja " & _
           "ORDER BY b.Descripcion,c.Nombre"
    bDameEmpleados_GastosNoValidados = gBdGALILEO.bRecordset(sSql, afoRegDatos)
           
End Function

Public Function bCIF_Proveedores_SinConciliar(iCodPais As Integer) As Boolean

    Dim afoRegAux As ADODB.Recordset

    bCIF_Proveedores_SinConciliar = False

    sSql = "SELECT Contador,Fecha,Donde,CIF_Proveedor,Ruta_Ticket,Importe_Ori,Moneda " & _
           "FROM Gastos_Internos " & _
           "WHERE CodPais = " & iCodPais & " AND TipoGasto = 11 AND NVL(Proveedor_OK,'N') = 'N' " & _
           "ORDER BY Fecha,Contador"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Function
    End If
    
    'Recorro todos los gastos sin el proveedor conciliado e intento conciliarlo automáticamente
    While Not afoRegistro.EOF
    
        sSql = "SELECT id FROM GES_Vendor WHERE NVL(idFiscal,'') = '" & NoNull(afoRegistro("CIF_Proveedor"), "A") & "'"
        If gBdGALILEO.bRecordset(sSql, afoRegAux) Then
            sSql = "UPDATE Gastos_Internos SET Proveedor_OK = 'S',id_Proveedor = '" & afoRegAux("id") & "' WHERE Contador = " & afoRegistro("Contador")
            gBdGALILEO.bExecute sSql
        Else
            bCIF_Proveedores_SinConciliar = True
        End If
        Set afoRegAux = Nothing
    
        afoRegistro.MoveNext
        
    Wend
    Set afoRegistro = Nothing
           
End Function

Public Function bGastosProveedores_SinConciliar(iCodPais As Integer, afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT a.Contador,a.Fecha,a.Donde,a.CIF_Proveedor,a.Ruta_Ticket,a.Importe_Ori,a.Moneda,a.Codigo,b.Nombre " & _
           "FROM Gastos_Internos a,Usuarios_Personal b " & _
           "WHERE a.CodPais = " & iCodPais & " AND a.TipoGasto = 11 AND NVL(a.Proveedor_OK,'N') = 'N' " & _
           "AND a.CodPais = b.CodPais AND a.Codigo = b.Codigo AND a.Jerarquia = b.Jerarquia " & _
           "ORDER BY a.Fecha,a.Contador"
    bGastosProveedores_SinConciliar = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function

Public Function bProveedoresGES(afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT * FROM GES_Vendor ORDER BY Nombre"
    bProveedoresGES = gBdGALILEO.bRecordset(sSql, afoRegDatos)
    
End Function

Public Function Logon(sServidor As String) As Boolean

    Dim sIP_Servidor As String
    Dim sSystem As String
    Dim sMand As String
    Dim sUSU As String
    Dim sPass As String
    
    Set oBapiControl = CreateObject("SAP.BAPI.1")
    Set oConnection = oBapiControl.Connection
    oConnection.MessageServer = ""
    Select Case LCase(sServidor)
    Case "ges_produccion"
    
        sIP_Servidor = "172.25.205.12"
        sSystem = "00"
        sMand = "300"
        sUSU = "ZWS_USER"
        sPass = "01interface"
    
    Case "ges_calidad"
    
        sIP_Servidor = "172.25.205.41"
        sSystem = "00"
        sMand = "300"
        sUSU = "ZWS_USER"
        sPass = "01interface"
        
    Case "ges_desarrollo"
    
        sIP_Servidor = "172.25.205.16"
        sSystem = "00"
        sMand = "300"
        sUSU = "ZWS_USER"
        sPass = "01interface"
        
    Case Else
    
        Set oConnection = Nothing
        Logon = False
        MsgBox "Logon failed.", , ""
        Exit Function
        
    End Select
    
    oConnection.ApplicationServer = sIP_Servidor
    oConnection.client = sMand
    oConnection.User = sUSU
    oConnection.Password = sPass
    oConnection.Language = "ES"
    
    If Not oConnection.Logon(0, True) Then
        Set oConnection = Nothing
        Logon = False
        MsgBox "Logon failed.", , ""
    Else
        Logon = True
    End If
    
End Function

Public Function DesLogon()
    On Error Resume Next
    oConnection.Logoff
End Function

Public Function bDameDatosGastoInterno_Contador(iCodPais As Integer, lContador As Long, afoRegDatos As ADODB.Recordset) As Boolean

    Select Case iCodPais
    Case giPAIS_GES
    
        sSql = "SELECT a.*,d.NombreGasto,b.Tipo Tipo_PEPCECO,c.id idProveedor,NVL(c.Nombre,'') Nombre_Proveedor " & _
               "FROM Gastos_Internos a,TiposGasto d,OUTER GES_PepCeco b,OUTER GES_Vendor c " & _
               "WHERE a.CodPais = 92 AND a.Contador = " & lContador & _
               "AND a.Departamento = b.Sociedad AND LOWER(a.PEP_CECO) = LOWER(b.PEPCECO) " & _
               "AND LOWER(NVL(a.CIF_Proveedor,'')) = LOWER(NVL(c.idFiscal,'')) " & _
               "AND d.Tipo = '7' AND a.TipoGasto = d.TipoGasto"
    
    Case Else
        sSql = "SELECT * FROM Gastos_Internos WHERE Contador = " & lContador
    End Select
    bDameDatosGastoInterno_Contador = gBdGALILEO.bRecordset(sSql, afoRegDatos)
    
End Function

Public Function bAsignarProveedorGasto(iCodPais As Integer, lContadorGasto As Long, idProveedor As String, sIDFiscal As String) As Boolean

    sSql = "UPDATE Gastos_Internos SET id_Proveedor = '" & idProveedor & "',Proveedor_OK = 'S',CIF_Proveedor = '" & sIDFiscal & "' " & _
           "WHERE CodPais = " & iCodPais & " AND Contador = " & lContadorGasto
    bAsignarProveedorGasto = gBdGALILEO.bExecute(sSql)

End Function

Public Function bComprobarCIF_Proveedor_GES(sCIF_Proveedor As String, afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT * FROM GES_Vendor WHERE LOWER(idFiscal) = '" & LCase(sCIF_Proveedor) & "'"
    bComprobarCIF_Proveedor_GES = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function

Public Function bMarcarGasto_ContabilizadoSAP(lContador As Long, sDocumento_SAP As String, sFechaContaSAP As String) As Boolean

    sSql = "UPDATE Gastos_Internos SET Conta_SAP = 'S',Documento_SAP = '" & sDocumento_SAP & "',Fecha_Conta_SAP = '" & sFechaContaSAP & "' WHERE Contador = " & lContador
    bMarcarGasto_ContabilizadoSAP = gBdGALILEO.bExecute(sSql)
    
End Function

Public Function bRechazarGasto(lContador As Long, sMotivoRechazo As String) As Boolean

    sSql = "UPDATE Gastos_Internos SET Validado = 'R',Mot_Rechazo = 'Rechazo ADMIN: " & sMotivoRechazo & "',Env_Mail_Rechazo = 'N',Fecha_Val_Rec = CURRENT YEAR TO MINUTE " & _
           "WHERE Contador = " & lContador
    bRechazarGasto = gBdGALILEO.bExecute(sSql)
    
End Function

Public Function bEmpleadoBloqueado(iCodPais As Integer, iCodDepartamento As Integer, iCodEmpleado As Long, sJerEmpleado As String) As Boolean

    bEmpleadoBloqueado = False
    
    sSql = "SELECT NVL(Bloqueado_Por,'') FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " AND Jerarquia = '" & sJerEmpleado & "' AND Codigo = " & iCodEmpleado
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then If Trim(afoRegistro(0)) <> "" Then bEmpleadoBloqueado = True
    Set afoRegistro = Nothing

End Function

Public Function sEmpleadoBloqueado_Por(iCodPais As Integer, iCodDepartamento As Integer, iCodEmpleado As Long, sJerEmpleado As String) As String

    sEmpleadoBloqueado_Por = ""
    
    sSql = "SELECT NVL(a.Bloqueado_Por,''),b.Nombre " & _
           "FROM Usuarios_Personal a,OUTER prf_Usuarios b " & _
           "WHERE a.CodPais = " & iCodPais & " AND a.Departamento = " & iCodDepartamento & " AND a.Jerarquia = '" & sJerEmpleado & "' AND a.Codigo = " & iCodEmpleado & " " & _
           "AND NVL(a.Bloqueado_Por,'') = b.id_Usuario"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then If Trim(afoRegistro(0)) <> "" Then sEmpleadoBloqueado_Por = Trim(afoRegistro("Nombre"))
    Set afoRegistro = Nothing

End Function

Public Sub EmpleadoBloquear(iCodPais As Integer, iCodDepartamento As Integer, iCodEmpleado As Long, sJerEmpleado As String, sUsuarioBloqueo As String)

    sSql = "UPDATE Usuarios_Personal SET Bloqueado_Por = '" & sUsuarioBloqueo & "' " & _
           "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " AND Jerarquia = '" & sJerEmpleado & "' AND Codigo = " & iCodEmpleado
    gBdGALILEO.bExecute sSql

End Sub

Public Sub EmpleadoDesBloquear(iCodPais As Integer, iCodDepartamento As Integer, iCodEmpleado As Long, sJerEmpleado As String)

    sSql = "UPDATE Usuarios_Personal SET Bloqueado_Por = '' " & _
           "WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " AND Jerarquia = '" & sJerEmpleado & "' AND Codigo = " & iCodEmpleado
    gBdGALILEO.bExecute sSql

End Sub

Public Sub EmpleadoDesBloquear_UsuGasto()

    sSql = "UPDATE Usuarios_Personal SET Bloqueado_Por = '' " & _
           "WHERE Bloqueado_Por = '" & goUsuario.IdUsuario & "'"
    gBdGALILEO.bExecute sSql

End Sub

Public Function bDamePassUsuInterno(iCodPais As Integer, iCodDepartamento As Integer, iCodigo As Long, sPass As String) As Boolean

    bDamePassUsuInterno = False
    sPass = ""
    
    sSql = "SELECT Clave FROM Usuarios_Personal WHERE CodPais = " & iCodPais & " AND Departamento = " & iCodDepartamento & " AND Codigo = " & iCodigo
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        sPass = Trim(afoRegistro("Clave"))
        bDamePassUsuInterno = True
    End If
    Set afoRegistro = Nothing
    
End Function


Public Function NoNulo(vValor As Variant, Optional sTipo As String = "T") As Variant

    'Los diferentes que podrá tomar sTipo serán los siguientes:
    'T=Texto; N=Numérico;

    On Error GoTo Error_Rutina
    
    If IsNull(vValor) Then
    
        Select Case sTipo
        Case "T"
            NoNulo = ""
        Case "N"
            NoNulo = 0
        Case Else
            NoNulo = ""
        End Select
        
    Else
    
        If vValor = "" Then
        
            Select Case sTipo
            Case "T"
                NoNulo = ""
            Case "N"
                NoNulo = 0
            Case Else
                NoNulo = ""
            End Select
            
        Else
            NoNulo = vValor
        End If
        
    End If
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Function

Public Function bCuentaContableValida(sCuentaContable As String) As Boolean

    bCuentaContableValida = False
    
    If CodPaisTrabajo <> giPAIS_GES Then Exit Function
    
    sSql = "SELECT id FROM GES_CuentasContables WHERE id = '" & sCuentaContable & "'"
    bCuentaContableValida = gBdGALILEO.bRecordset(sSql, afoRegistro)
    Set afoRegistro = Nothing

End Function

Public Function bPreValidacion_DameOficinas(iCodPais As Integer, afoRegDatos As ADODB.Recordset) As Boolean

    bPreValidacion_DameOficinas = False

    sSql = "SELECT * FROM Gastos_Oficinas_PreValidacion WHERE CodPais = " & iCodPais & " ORDER BY Nombre"
    bPreValidacion_DameOficinas = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function

Public Sub PreValidacion_CrearOficina(iCodPais As Integer, sOficina As String)

    sSql = "SELECT * FROM Gastos_Oficinas_PreValidacion WHERE CodPais = " & iCodPais & " AND Nombre = '" & sOficina & "'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Sub
    End If
    Set afoRegistro = Nothing
    
    sSql = "INSERT INTO Gastos_Oficinas_PreValidacion (CodPais,Nombre) VALUES (" & iCodPais & ",'" & sOficina & "')"
    gBdGALILEO.bExecute sSql
    
End Sub

Public Function bPreValidacion_DameResponsablesAsignadosOficina(iCodPais As Integer, id_Oficina As Integer, afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT a.Usu_Responsable,b.Departamento,b.Nombre " & _
           "FROM Gastos_Relacion_PreValidacion a,Usuarios_Personal b " & _
           "WHERE a.CodPais = " & iCodPais & " AND a.id_Ofi_PreValidacion = " & id_Oficina & " " & _
           "AND a.CodPais = b.CodPais AND a.Usu_Responsable = b.Usuario " & _
           "ORDER BY b.Nombre"
    bPreValidacion_DameResponsablesAsignadosOficina = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function

Public Function bPreValidacion_DameEstructurasOrganizativasAsignadosOficina(iCodPais As Integer, id_Oficina As Integer, afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT TRIM(Departamento) Departamento,TRIM(Division_Personal) Division_Personal,TRIM(Clave_Organizacion) Clave_Organizacion " & _
           "FROM Gastos_Relacion_PreValidacionV2 " & _
           "WHERE CodPais = " & iCodPais & " AND id_Ofi_PreValidacion = " & id_Oficina & " " & _
           "ORDER BY Departamento,Division_Personal,Clave_Organizacion"
    bPreValidacion_DameEstructurasOrganizativasAsignadosOficina = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function


Public Function bPreValidacion_DameResponsables(iCodPais As Integer, id_Oficina As Integer, lDepartamento As Long, afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT Codigo,Nombre " & _
           "FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " AND Codigo IN (SELECT DISTINCT Responsable FROM GES_PEPCECO WHERE 1 = 1"
    If lDepartamento > 0 Then sSql = sSql & " AND Sociedad = '" & Format(lDepartamento, "0000") & "'"
    sSql = sSql & ") AND Usuario NOT IN (SELECT Usu_Responsable FROM Gastos_Relacion_PreValidacion WHERE CodPais = " & iCodPais & " AND id_Ofi_PreValidacion = " & id_Oficina & ") " & _
           "ORDER BY Nombre"
    bPreValidacion_DameResponsables = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function

Public Function bPreValidacion_DameEstructurasOrganizativas(iCodPais As Integer, afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT DISTINCT Departamento,TRIM(NVL(Division_Personal,'')) Division_Personal,TRIM(NVL(Clave_Organizacion,'')) Clave_Organizacion " & _
           "FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais
    sSql = sSql & " AND Departamento||TRIM(NVL(Division_Personal,''))||TRIM(NVL(Clave_Organizacion,'')) NOT IN " & _
                    "(SELECT TRIM(Departamento)||TRIM(Division_Personal)||TRIM(Clave_Organizacion) FROM Gastos_Relacion_PreValidacionV2 WHERE CodPais = " & iCodPais & ") " & _
           "ORDER BY Departamento,Division_Personal,Clave_Organizacion"
    bPreValidacion_DameEstructurasOrganizativas = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function

Public Function bPreValidacion_DamePreValidadoresAsignadosOficina(iCodPais As Integer, id_Oficina As Integer, afoRegDatos As ADODB.Recordset) As Boolean

    sSql = "SELECT Usuario,Codigo,Nombre " & _
           "FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " AND NVL(id_Ofi_PreValidacion,0) = " & id_Oficina & " " & _
           "ORDER BY Nombre"
    bPreValidacion_DamePreValidadoresAsignadosOficina = gBdGALILEO.bRecordset(sSql, afoRegDatos)
    
End Function

Public Function bPreValidacion_DamePreValidadores(iCodPais As Integer, lDepartamento As Long, afoRegDatos As ADODB.Recordset) As Boolean

    '12-11-19 No vamos a tener en cuenta el departamento
'    sSql = "SELECT Codigo,Nombre " & _
'           "FROM Usuarios_Personal " & _
'           "WHERE CodPais = " & iCodPais & " AND Codigo NOT IN (SELECT DISTINCT Responsable FROM GES_PEPCECO WHERE Sociedad = '" & Format(lDepartamento, "0000") & "') " & _
'           "AND NVL(id_Ofi_PreValidacion,0) = 0 AND Departamento = " & lDepartamento & " " & _
'           "ORDER BY Nombre"

    sSql = "SELECT Codigo,Nombre " & _
           "FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " AND Codigo NOT IN (SELECT DISTINCT Responsable FROM GES_PEPCECO) AND Codigo NOT IN (SELECT DISTINCT Director_Proyecto FROM GES_PEPCECO) " & _
           "AND NVL(id_Ofi_PreValidacion,0) = 0 " & _
           "ORDER BY Nombre"
    bPreValidacion_DamePreValidadores = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function

Public Function bPreValidacion_DamePreValidadoresV2(iCodPais As Integer, afoRegDatos As ADODB.Recordset) As Boolean

    '11-04-22 Por indicación de Mikel, quitamos el control de que los responsables o directores de proyecto de centros de coste no puedan ser validadores
'    sSql = "SELECT Codigo,Nombre " & _
'           "FROM Usuarios_Personal " & _
'           "WHERE CodPais = " & iCodPais & " AND Codigo NOT IN (SELECT DISTINCT Responsable FROM GES_PEPCECO) AND Codigo NOT IN (SELECT DISTINCT Director_Proyecto FROM GES_PEPCECO) " & _
'           "AND NVL(id_Ofi_PreValidacion,0) = 0 " & _
'           "ORDER BY Nombre"
    sSql = "SELECT Codigo,Nombre " & _
           "FROM Usuarios_Personal " & _
           "WHERE CodPais = " & iCodPais & " " & _
           "AND NVL(id_Ofi_PreValidacion,0) = 0 " & _
           "ORDER BY Nombre"
    '11-04-22 Fin
    bPreValidacion_DamePreValidadoresV2 = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function


Public Function bPreValidacion_AsignarResponsable(iCodPais As Integer, lCodigo_Responsable As Long, id_Oficina As Integer) As Boolean

    bPreValidacion_AsignarResponsable = False

    sSql = "SELECT Usuario FROM Usuarios_Personal WHERE CodPais = " & iCodPais & " AND Codigo = " & lCodigo_Responsable
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Function
    End If
    sSql = "INSERT INTO Gastos_Relacion_PreValidacion (id_Ofi_PreValidacion,Usu_Responsable,CodPais) " & _
           "VALUES (" & id_Oficina & ",'" & Trim(afoRegistro("Usuario")) & "'," & iCodPais & ")"
    bPreValidacion_AsignarResponsable = gBdGALILEO.bExecute(sSql)
    Set afoRegistro = Nothing

End Function

Public Function bPreValidacion_AsignarPreValidador(iCodPais As Integer, lCodigo_PreValidador As Long, id_Oficina As Integer) As Boolean

    bPreValidacion_AsignarPreValidador = False
    
    sSql = "SELECT Usuario FROM Usuarios_Personal WHERE CodPais = " & iCodPais & " AND Codigo = " & lCodigo_PreValidador & " AND NVL(id_Ofi_Prevalidacion,0) = 0"
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Function
    End If
    sSql = "UPDATE Usuarios_Personal SET id_Ofi_PreValidacion = " & id_Oficina & " " & _
           "WHERE CodPais = " & iCodPais & " AND Usuario = '" & Trim(afoRegistro("Usuario")) & "'"
    bPreValidacion_AsignarPreValidador = gBdGALILEO.bExecute(sSql)
    Set afoRegistro = Nothing

End Function

Public Function bPreValidacion_EliminarAsignacionResponsableOficina(iCodPais As Integer, id_Oficina As Integer, sUsu_Responsable As String) As Boolean

    bPreValidacion_EliminarAsignacionResponsableOficina = False
    
    gBdGALILEO.BeginTrans True
    
    sSql = "DELETE FROM Gastos_Relacion_PreValidacion " & _
           "WHERE CodPais = " & iCodPais & " AND id_Ofi_PreValidacion = " & id_Oficina & " AND Usu_Responsable = '" & sUsu_Responsable & "'"
    If Not gBdGALILEO.bExecute(sSql) Then
        gBdGALILEO.RollbackTrans
        Exit Function
    End If
    
    sSql = "SELECT * FROM Gastos_Relacion_PreValidacion WHERE id_Ofi_PreValidacion = " & id_Oficina
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
    'No hay ningún responsable en la oficina, por lo que eliminamos los prevalidadores asignados
        sSql = "UPDATE Usuarios_Personal SET id_Ofi_PreValidacion = 0 WHERE CodPais = " & iCodPais & " AND NVL(id_Ofi_PreValidacion,0) = " & id_Oficina
        If Not gBdGALILEO.bExecute(sSql) Then
            gBdGALILEO.RollbackTrans
            Set afoRegistro = Nothing
            Exit Function
        End If
    End If
    Set afoRegistro = Nothing
    
    gBdGALILEO.CommitTrans
    bPreValidacion_EliminarAsignacionResponsableOficina = True

End Function

Public Function bPreValidacion_EliminarPrevalidador(iCodPais As Integer, sUsuario As String) As Boolean

    sSql = "UPDATE Usuarios_Personal SET id_Ofi_PreValidacion = 0 WHERE CodPais = " & iCodPais & " AND Usuario = '" & Trim(sUsuario) & "'"
    bPreValidacion_EliminarPrevalidador = gBdGALILEO.bExecute(sSql)
    
End Function

Public Function bEsUsuarioPreValidacion(iCodPais As Integer, sUsuario As String, id_Ofi_PreValidacion As Integer) As Boolean

    bEsUsuarioPreValidacion = False
    id_Ofi_PreValidacion = 0
    
    sSql = "SELECT NVL(id_Ofi_PreValidacion,0) FROM Usuarios_Personal WHERE CodPais = " & iCodPais & " AND Usuario = '" & sUsuario & "'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then If afoRegistro(0) > 0 Then id_Ofi_PreValidacion = afoRegistro(0)
    Set afoRegistro = Nothing
    
End Function

Public Function LimpiarCadena(sCadena As String) As String

    On Error GoTo Error_Rutina
    
    LimpiarCadena = Trim(sCadena)
    
    If Trim(sCadena = "") Then Exit Function
    
    While InStr(1, sCadena, "'")
        sCadena = Mid(sCadena, 1, InStr(1, sCadena, "'") - 1) & "`" & Mid(sCadena, InStr(1, sCadena, "'") + 1)
    Wend
    
    While InStr(1, sCadena, """")
        sCadena = Mid(sCadena, 1, InStr(1, sCadena, """") - 1) & "`" & Mid(sCadena, InStr(1, sCadena, """") + 1)
    Wend
    
    While InStr(1, sCadena, ";")
        sCadena = Mid(sCadena, 1, InStr(1, sCadena, ";") - 1) & "," & Mid(sCadena, InStr(1, sCadena, ";") + 1)
    Wend
    
    LimpiarCadena = Trim(sCadena)
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Function

Public Function sDameNombreUsuario(iCodPais As Integer, sUsuario As String) As String
    
    sDameNombreUsuario = ""
    If iCodPais <> giPAIS_GES Then Exit Function
    
    sSql = "SELECT Codigo,Nombre FROM Usuarios_Personal WHERE CodPais = " & iCodPais & " AND Usuario = '" & Trim(sUsuario) & "'"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then sDameNombreUsuario = "(" & afoRegistro("Codigo") & ") " & Trim(afoRegistro("Nombre"))
    Set afoRegistro = Nothing

End Function

Public Function sDameValidador(lContador As Long, iCodPais As Integer, iCodDepartamento As Integer, lCodigo As Long, sPEPCECO As String, sValidado As String) As String

    Dim sAux As String

    sDameValidador = ""
    
    On Error GoTo Exit_Rutina
    
    If iCodPais <> giPAIS_GES Then Exit Function
    
    If sValidado = "S" Then
    
        sSql = "SELECT NVL(Validacion,'') FROM Gastos_Internos WHERE CodPais = " & iCodPais & " AND Contador = " & lContador
        If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
            Set afoRegistro = Nothing
            Exit Function
        End If
        sAux = Trim(afoRegistro(0))
        sDameValidador = sAux
        If sAux = "" Then
            Set afoRegistro = Nothing
            Exit Function
        End If
        sSql = "SELECT Codigo,Nombre FROM Usuarios_Personal WHERE CodPais = " & iCodPais & " AND Codigo = " & Mid(sAux, 1, InStr(1, sAux, "_") - 1)
        sAux = Mid(sAux, InStr(1, sAux, "_") + 1)
        sSql = sSql & " AND Jerarquia = '" & Mid(sAux, 1, InStr(1, sAux, "_") - 1) & "'"
        sAux = Mid(sAux, InStr(1, sAux, "_") + 1)
        sSql = sSql & " AND Departamento = " & sAux
    
    Else
    
        sSql = "SELECT Responsable FROM GES_PEPCECO WHERE Sociedad = '" & Format(iCodDepartamento, "0000") & "' AND PEPCECO = '" & sPEPCECO & "'"
        If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
            Set afoRegistro = Nothing
            Exit Function
        End If
        If afoRegistro("Responsable") = lCodigo Then
            sSql = "SELECT b.Codigo,b.Nombre FROM Usuarios_Personal a,Usuarios_Personal b " & _
                   "WHERE a.CodPais = " & iCodPais & " AND a.Codigo = " & afoRegistro("Responsable") & " " & _
                   "AND a.CodPais = b.CodPais AND a.Manager_GES = b.Codigo"
        Else
            sSql = "SELECT Codigo,Nombre FROM Usuarios_Personal WHERE CodPais = " & iCodPais & " AND Codigo = " & afoRegistro("Responsable")
        End If
    
    End If
    
    If Not gBdGALILEO.bRecordset(sSql, afoRegistro) Then
        Set afoRegistro = Nothing
        Exit Function
    End If
    sDameValidador = "(" & afoRegistro("Codigo") & ") " & Trim(afoRegistro("Nombre"))
    
    Set afoRegistro = Nothing
    
Exit_Rutina:
    Exit Function
End Function

Public Function bDameGastosInternos_Informe(iCodPais As Integer, sFechaDesde As String, sFechaHasta As String, iCodDepartamento As Integer, _
    lCodigo As Long, sJerarquia As String, afoRegDatos As ADODB.Recordset, sTipoBusqueda As String) As Boolean

    If sTipoBusqueda = "gastos" Then sSql = "AND a.TipoGasto NOT IN (99) "
    If sTipoBusqueda = "anticipos" Then sSql = "AND a.TipoGasto IN (99) "

    sSql = "SELECT a.Contador,a.Codigo,b.Nombre,a.Departamento,a.PEP_CECO,e.Descripcion TipoGasto,a.Importe,a.Importe_Ori,a.Cambio,a.Moneda,a.Fecha," & _
           "a.Donde,NVL(a.Validado,'N') Validado,NVL(a.Conta_SAP,'N') Conta_SAP,NVL(a.Pagado,'N') Pagado," & _
           "CASE WHEN NVL(a.Ofi_PreValidacion,0) > 0 THEN NVL(a.Prevalidado,'N') ELSE '-' END PreValidado," & _
           "CASE WHEN NVL(a.Ofi_PreValidacion,0) > 0 THEN (SELECT Nombre FROM Gastos_Oficinas_PreValidacion WHERE id = NVL(a.Ofi_PreValidacion,0)) ELSE '-----' END OficinaPrevalidacion," & _
           "NVL(a.Validador,0) Validador,NVL(f.Nombre,'-----') NombreValidador " & _
           "FROM Gastos_Internos a,Usuarios_Personal b,Gastos_Empresas c ,TiposGasto d,TiposGasto_Idm e,OUTER Usuarios_Personal f " & _
           "WHERE a.CodPais = " & iCodPais & " AND a.CodPais = b.CodPais AND a.CodPais = c.CodPais " & _
           "AND a.Codigo = b.Codigo AND a.Jerarquia = b.Jerarquia AND a.Departamento = b.Departamento " & _
           "AND a.TipoGasto||'' = d.TipoGasto AND c.id||'' = d.Tipo AND d.idGasto = e.idGasto AND e.idIdioma = 1 " & _
           "AND a.Fecha BETWEEN '" & sFechaDesde & "' AND '" & sFechaHasta & "' AND a.CodPais = f.CodPais AND NVL(a.Validador,0) = f.Codigo " & sSql
    If iCodDepartamento > 0 Then sSql = sSql & "AND a.Departamento = " & iCodDepartamento & " "
    If lCodigo > 0 Then sSql = sSql & "AND a.Codigo = " & lCodigo & " AND a.Jerarquia = '" & sJerarquia & "' "
    sSql = sSql & "ORDER BY a.Contador"
    bDameGastosInternos_Informe = gBdGALILEO.bRecordset(sSql, afoRegDatos)

End Function

Public Function sDameMonedaDepartamento(iCodPais As Integer, iCodDepartamento As Integer) As String

    sDameMonedaDepartamento = "EUR"
    If iCodPais <> giPAIS_GES Then Exit Function

    sSql = "SELECT NVL(Divisa,'EUR') " & _
           "FROM DepartamentosCL " & _
           "WHERE CodPais = " & iCodPais & " AND Codigo = " & iCodDepartamento
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then sDameMonedaDepartamento = afoRegistro(0)
    Set afoRegistro = Nothing
    
End Function

Public Function bPreValidacion_AsignarEstructuraOrganizativa(iCodPais As Integer, sEstructuraOrganizativa As String, id_Oficina As Integer) As Boolean

    Dim sAux As String
    Dim sDepartamento As String
    Dim sDivision_Personal As String
    Dim sClave_Organizacion As String

    bPreValidacion_AsignarEstructuraOrganizativa = False
    
    sAux = sEstructuraOrganizativa
    sDepartamento = Trim(Mid(sAux, 1, InStr(1, sAux, "-") - 1))
    sAux = Trim(Mid(sAux, InStr(1, sAux, "-") + 1))
    sDivision_Personal = Trim(Mid(sAux, 1, InStr(1, sAux, "-") - 1))
    sAux = Trim(Mid(sAux, InStr(1, sAux, "-") + 1))
    sClave_Organizacion = sAux
    
    sSql = "INSERT INTO Gastos_Relacion_PreValidacionV2 (id_Ofi_PreValidacion,Departamento,Division_Personal,Clave_Organizacion,CodPais) " & _
           "VALUES (" & id_Oficina & ",'" & sDepartamento & "','" & sDivision_Personal & "','" & sClave_Organizacion & "'," & iCodPais & ")"
    bPreValidacion_AsignarEstructuraOrganizativa = gBdGALILEO.bExecute(sSql)
    Set afoRegistro = Nothing

End Function

Public Function bPreValidacion_EliminarAsignacionEstructuraOrganizativaOficina(iCodPais As Integer, id_Oficina As Integer, _
    sDepartamento As String, sDivision_Personal As String, sClave_Organizacion As String) As Boolean
    
    sSql = "DELETE FROM Gastos_Relacion_PreValidacionV2 " & _
           "WHERE CodPais = " & iCodPais & " AND id_Ofi_PreValidacion = " & id_Oficina & " " & _
           "AND Departamento = '" & sDepartamento & "' AND Division_Personal = '" & sDivision_Personal & "' AND Clave_Organizacion = '" & sClave_Organizacion & "'"
    bPreValidacion_EliminarAsignacionEstructuraOrganizativaOficina = gBdGALILEO.bExecute(sSql)
    
End Function

Public Sub InicializarRuta()

    Dim Fso As New FileSystemObject
    Dim Drv As Drive
    
    On Error GoTo Error_Rutina
    
    If goUsuario.IdUsuario = "fsanchezc" Then
        gsRutaEquipo = "\\tsclient\C\"
        Exit Sub
    End If
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set Drv = Fso.GetDrive(Fso.GetDriveName("V:"))
    If Err.Number = 68 Then
        
        On Error GoTo Error_Rutina
        Err.Clear
        gsRutaEquipo = "C:\"
        
    Else
        
        On Error GoTo Error_Rutina
        If Drv.DriveType = Remote And Drv.ShareName = "\\Client\C$" Then
            gsRutaEquipo = "V:\"
        Else
            gsRutaEquipo = "C:\"
        End If
        
    End If
    Set Drv = Nothing
    Set Fso = Nothing

Exit_Rutina:
    Exit Sub
Error_Rutina:
    If gsRutaEquipo = "" Then
        Select Case goUsuario.PaisTrabajo
        Case 34
            gsRutaEquipo = "C:\"
        Case Else
            gsRutaEquipo = "V:\"
        End Select
    End If
    'MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Public Function Encrypt(InPutStr As String) As String
        
    Dim sCryptKey As String
    Dim sCadena As String
    Dim sCaracter As String
    Dim i As Integer
    Dim X As Integer
    Dim sOutPutStr As String
    Dim cc As Integer
    
    On Error GoTo Exit_Rutina
    
    sCadena = InPutStr
    If Len(sCadena) < 20 Then
        sCadena = sCadena + Mid("24F1j1PcWeJf35uhdp+(t8SnU+[vrF1jVeu/n8oFBurFKk6NesH34hi34z/Kv", Len(InPutStr), 20 - Len(InPutStr))
    End If
    sCryptKey = "GNQ}4i0-[CldnU+[vrF1j1PcWeJfVv4QGBurFK6}[l[H1S:oY\v@U}i,oD]f/n8oFk6NesH--PJeCLdp+(t8SVe:ewY(wR9p-CzG<,Q/(U[.pXDiz/KvnXPBXnk"
    
    sOutPutStr = ""
    cc = 0

    For i = 1 To Len(sCadena)

        X = Asc(Mid(sCadena, i, 1))
        X = X - 40
        If X < 0 Then X = X + 125
        While X < 0
            X = X + 7
        Wend

        X = X + Asc(Mid(sCryptKey, cc + 1, 1))
        If X > 125 Then X = X - 125
        While X > 125
            X = X - 3
        Wend
        If X < 40 Then X = X + 40
        While X < 40
            X = X + 4
        Wend

        sCaracter = Chr(X)
        Select Case sCaracter
            Case "<"
                sCaracter = "U"
            Case "*"
                sCaracter = "0"
            Case "?"
                sCaracter = ">"
            Case "€"
                sCaracter = "E"
            Case Else
                sCaracter = sCaracter
        End Select
        sOutPutStr = sOutPutStr & sCaracter

        cc = (cc + 1) Mod Len(sCryptKey)

    Next

    Encrypt = sOutPutStr
    
Exit_Rutina:
    Exit Function
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Function

Public Sub ListarEncriptacion()

    Dim sFichero As String
    Dim iFichero As Integer
    Dim sAux As String
    Dim value As String
    
    On Error GoTo Exit_Rutina
    
    sSql = "SELECT Usuario,CodPais,Codigo,Departamento,Clave,Pass " & _
           "FROM Usuarios_Personal " & _
           "ORDER BY CodPais,Departamento,Codigo"
    If gBdGALILEO.bRecordset(sSql, afoRegistro) Then
    
        sFichero = "C:\GeneracionCSV\Encrypt_UsuGastos_vb6.csv"
        iFichero = FreeFile
        Open sFichero For Output As #iFichero
        
        Print #iFichero, "CodPais;Usuario;Codigo;Departamento;Clave;Pass;Encriptacion"
        
        While Not afoRegistro.EOF
            
            sAux = Encrypt(Trim(afoRegistro("Clave")))
            value = Trim(afoRegistro("CodPais")) & ";" & Trim(afoRegistro("Usuario")) & ";" & Trim(afoRegistro("Codigo")) & ";" & Trim(afoRegistro("Departamento")) & ";""" & Trim(afoRegistro("Clave")) & """;""" & Trim(afoRegistro("Pass")) & """;""" & sAux & """"
            Print #iFichero, value
            
            afoRegistro.MoveNext
            
        Wend

    Else
        MsgBox "No se han encontrado datos", vbApplicationModal + vbExclamation, "Información"
    End If
    MsgBox "Acción finalizada", vbApplicationModal + vbInformation, "Información"
    
Exit_Rutina:
    Close
    Exit Sub
Error_Rutina:
    MsgBox Err.Number & " - " & Err.Description, vbApplicationModal + vbExclamation, "Error"
    Resume Exit_Rutina
    Resume Next
End Sub

Public Function Encrypt_v2(ByVal Word As String, Optional ByVal Key As String = "", Optional ByVal Mode As Boolean = False) As String
    
    Dim miKey As String
    Dim w As Long, k As Long, p As Long, j As Long, NuChr As Long
    Dim Cd As String, Kd As String, Rd As String
    
    miKey = "GNQ}4i0-[CldnU+[vrF1j1PcWeJfVv4QGBurFK6}[l[H1S:oY\v@U}i,oD]f/n8oFk6NesH--PJeCLdp+(t8SVe:ewY(wR9p-CzG<,Q/(U[.pXDiz/KvnXPBXnk"
    If Trim(Key) <> "" Then miKey = Trim(Key)

    w = Len(Word)
    k = Len(miKey)
    
    ' Modalidad de Encripción...
    If Mode = False Then
        For j = 1 To w
            Cd = Mid(Word, j, 1)
            If p = k Then p = 0
            p = p + 1
            Kd = Mid(miKey, p, 1)
            NuChr = Asc(Cd) + Asc(Kd)
            If NuChr > 255 Then
                NuChr = NuChr - 255
            End If
            Rd = Rd & Chr(NuChr)
        Next
        Encrypt_v2 = Rd
        Exit Function
    End If
    
    ' Modalidad de Dencripción...
    If Mode = True Then
        For j = 1 To w
            Cd = Mid(Word, j, 1)
            If p = k Then p = 0
            p = p + 1
            Kd = Mid(miKey, p, 1)
            NuChr = Asc(Cd) - Asc(Kd)
            If NuChr < 0 Then
                NuChr = NuChr + 255
            End If
            Rd = Rd & Chr(NuChr)
        Next
        Encrypt_v2 = Rd
        Exit Function
    End If
    
End Function

Public Function ValidarMail(seMail As String) As Boolean
    Dim re As New RegExp
    re.Pattern = "\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+"
    If re.Test(seMail) Then
        ValidarMail = True
    Else
        ValidarMail = False
    End If
    Set re = Nothing
End Function
