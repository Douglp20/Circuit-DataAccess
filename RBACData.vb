Public Class RBACData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#Region "Get User Control"
    Public Function getRBACUserLoginPrivilege(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[UserLogin_get_role_privilege_by_UserID]"
        Dim Parameter As String = "@UserID"
        Dim Type As String = SqlDbType.Int



        getRBACUserLoginPrivilege = ViperCon.getSqlDataAdapterWithParameter(connection.getConnection, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getRBACUserLoginPrivilegePersmission(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[UserLogin_get_role_privilege_permission_by_UserID]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@UserID")
        Parameter.Add("@privilegeID")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)



        getRBACUserLoginPrivilegePersmission = ViperCon.getSqlDataAdapterWithParameters(connection.getConnection, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
#Region "Update Data"
    Public Sub UpdatePrivilegePersmission(value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[Update_RBAC_privilege_persmission]"

        Dim Parameter As New ArrayList
        Parameter.Add("@roleID")
        Parameter.Add("@privilegeID")
        Parameter.Add("@permissionID")
        Parameter.Add("@Checked")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub UpdatePrivilege(value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[Update_RBAC_privilege]"

        Dim Parameter As New ArrayList
        Parameter.Add("@roleID")
        Parameter.Add("@privilegeID")
        Parameter.Add("@Checked")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
#Region "Get Data"
    Public Function getRBACUsers(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[RBAC_Users_by_roleID]"

        Dim Parameter As String = "@RoleID"
        Dim Type As String = SqlDbType.Int



        getRBACUsers = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getRBACData() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[RBAC_get_rbac_role]"


        getRBACData = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function







Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)




    End Function

    Public Function getRBACPrivilegeByRoleID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[RBAC_get_rbac_role_privilege_by_roleID]"

        Dim Parameter As String = "@RoleID"
        Dim Type As String = SqlDbType.Int


        getRBACPrivilegeByRoleID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Function







Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)




    End Function
    Public Function getRBACPrivilegePermissionByRoleID(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[RBAC_get_rbac_role_privilege_Persmission_by_roleID]"

        Dim Parameter As New ArrayList
        Parameter.Add("@roleID")
        Parameter.Add("@privilegeID")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        getRBACPrivilegePermissionByRoleID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)



        Exit Function







Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)




    End Function

#End Region

#Region "Save Data"

    Public Sub SaveRBACRole(value As String)

        On Error GoTo Err

        Dim sp As String = "[Insert_RBAC_role]"


        Dim Parameter As String = "@Role"
        Dim Type As String = SqlDbType.VarChar



        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub UpdateRBACRole(value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[Update_RBAC_role]"

        Dim Parameter As New ArrayList
        Parameter.Add("@ID")
        Parameter.Add("@Role")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region

End Class
