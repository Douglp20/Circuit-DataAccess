Public Class UserLoginData
    Public Sub New()
    End Sub
    Public Event errorMessage(ByVal errDes As String, ByVal errNo As Integer, ByVal errTrace As String)

    Private WithEvents ViperCon As New douglas.Viper.Connection.Connection()
    Private Connection As New Connection

#Region "Error Control"
    Private Sub ErrorMessage_ViperCon(ByVal errDes As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim errMessage As String = ">> Called by the module : " + Me.ToString()
        RaiseEvent errorMessage(errDes, errNo, errTrace)
    End Sub


#End Region

#Region "The Login Process "
    Public Function getLoginID(LoginID As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[UserLogin_get_UserLogin_by_LoginID]"
        Dim strParameter As String = "@LoginID"
        Dim strType As String = SqlDbType.Char
        Dim strQueryString As String = LoginID

        ' Dim d As String = Connection.getConnection

        getLoginID = ViperCon.getSqlDataAdapterWithParameter(Connection.getConnection, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getForgotPasswordEmail(value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[UserLogin_get_forgot_password_email]"
        Dim Parameter As String = "@value"
        Dim Type As String = SqlDbType.Char

        getForgotPasswordEmail = ViperCon.getSqlDataAdapterWithParameter(Connection.getConnection, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region

#Region "Security Permission "
    Public Function getLoginIDPersmission(LoginID As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[UserLogin_get_UserLoginPermission_by_LoginID]"
        Dim strParameter As String = "@LoginID"
        Dim strType As String = SqlDbType.Char
        Dim strQueryString As String = LoginID


        getLoginIDPersmission = ViperCon.getSqlDataAdapterWithParameter(Connection.getConnection, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Function getRBACUserRoleUserID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[RBAC_role_by_UserID]"
        Dim Parameter As String = "@UserID"
        Dim Type As String = SqlDbType.Char


        getRBACUserRoleUserID = ViperCon.getSqlDataAdapterWithParameter(Connection.getConnection, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)


    End Function

#End Region
#Region "Save Data"
    Public Function UpdateRBACUserRole(ByRef Values As ArrayList)

        On Error GoTo Err




        Dim sp As String = "[Update_RBAC_user_role]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@RoleID")
        Parameter.Add("@UserID")
        Parameter.Add("@Checked")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)


        ViperCon.ExecuteProcessWithParameters(Connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function SaveUserLoginPassword(ByRef arrValues As ArrayList)

        On Error GoTo Err

       
        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        sp = "[update_User_Password]"

        arrParameter.Add("@UserName")
        arrParameter.Add("@Password")

        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)

        For i As Integer = 0 To arrValues.Count - 1
            arrValuesPass.Add(arrValues(i))
        Next


        ViperCon.ExecuteProcessWithParameters(Connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateUserLogin(ByRef Values As ArrayList)

        On Error GoTo Err




        Dim sp As String = "[update_User_Login]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@id")
        Parameter.Add("@StaffID")
        Parameter.Add("@LoginID")
        Parameter.Add("@per_Admin")
        Parameter.Add("@per_Payroll")
        Parameter.Add("@per_Staff")
        Parameter.Add("@per_invoice_admin")
        Parameter.Add("@per_invoice")
        Parameter.Add("@per_order_admin")
        Parameter.Add("@per_order")
        Parameter.Add("@per_company")
        Parameter.Add("@per_wholesale")
        Parameter.Add("@per_disable_login")
        Parameter.Add("@per_change_password")
        Parameter.Add("@per_maintenance")
        Parameter.Add("@Newpassword")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(Connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateUserLoginLocked(ByRef Values As ArrayList)

        On Error GoTo Err




        Dim sp As String = "[update_User_Login_Locked]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@LoginID")
        Parameter.Add("@locked")



        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)


        ViperCon.ExecuteProcessWithParameters(Connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Sub InsertUserLogin(ByRef Values As ArrayList)

        On Error GoTo Err


        Dim sp As String = "[insert_User_Login]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@StaffID")
        Parameter.Add("@LoginID")
        Parameter.Add("@per_Admin")
        Parameter.Add("@per_Payroll")
        Parameter.Add("@per_Staff")
        Parameter.Add("@per_invoice_admin")
        Parameter.Add("@per_invoice")
        Parameter.Add("@per_order_admin")
        Parameter.Add("@per_order")
        Parameter.Add("@per_company")
        Parameter.Add("@per_wholesale")
        Parameter.Add("@per_disable_login")
        Parameter.Add("@per_change_password")
        Parameter.Add("@per_maintenance")
        Parameter.Add("@Newpassword")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(Connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
#Region "Get Data"

    Public Function getLoginOldPassword(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[UserLogin_get_changepassword_by_LoginID]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@loginID")
        arrParameter.Add("@password")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)

        getLoginOldPassword = ViperCon.getSqlDataAdapterWithParameters(Connection.ConnectionString, sp, arrParameter, arrType, arrQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getStaffUserLoginPermsision(StaffID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[UserLogin_get_Staff_UserLoginPermission_by_StaffID]"
        Dim strParameter As String = "@StaffID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = StaffID


        getStaffUserLoginPermsision = ViperCon.getSqlDataAdapterWithParameter(Connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
End Class
