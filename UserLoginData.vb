Public Class UserLoginData
    Public Sub New()
    End Sub
    Public Event errorMessage(ByVal errDes As String, ByVal errNo As Integer, ByVal errTrace As String)

    Private WithEvents ViperCon As New Viper.Connection.Connection()
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

        Dim d As String = Connection.getConnection

        getLoginID = ViperCon.getSqlDataAdapterWithParameter(Connection.getConnection, sp, strParameter, strType, strQueryString)


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
#End Region
#Region "Save Data"

    Public Function SaveUserLoginPassword(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim ViperCon As New Viper.Connection.Connection
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
    Public Function SaveUserLogin(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim ViperCon As New Viper.Connection.Connection
        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList
        ''Insert a snew record

        If arrValues(0) = 0 Then
            sp = "[insert_User_Login]"
            For i As Integer = 1 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        Else
            arrParameter.Add("@id")
            arrType.Add(SqlDbType.Int)
            sp = "[update_User_Login]"
            For i As Integer = 0 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        End If

        arrParameter.Add("@StaffID")
        arrParameter.Add("@LoginID")
        arrParameter.Add("@per_Admin")
        arrParameter.Add("@per_Payroll")
        arrParameter.Add("@per_invoice_admin")
        arrParameter.Add("@per_invoice")
        arrParameter.Add("@per_order_admin")
        arrParameter.Add("@per_order")
        arrParameter.Add("@per_company")
        arrParameter.Add("@per_wholesale")
        arrParameter.Add("@per_disable_login")
        arrParameter.Add("@per_change_password")
        arrParameter.Add("@Newpassword")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(Connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)
    End Function
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
