Public Class TaskAssignmentData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New douglas.Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub


#Region "Error Control"
    Private Sub ErrorMessage_ViperCon(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#End Region
#Region "Service Data"
    Public Function GetTaskAssignmentHistory(value As String) As SqlClient.SqlDataAdapter
        On Error GoTo Err

        Dim sp As String = "[SERVICE_get_service_history]"
        Dim Parameter As String = "@source"
        Dim Type As String = SqlDbType.VarChar



        GetTaskAssignmentHistory = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region

#Region "Get  Data"


    Public Function getAllTastAssignment() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Task_get_all_taskAssignment]"

        getAllTastAssignment = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAllTastAssignmentByLoginID(loginID As String, Index As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Task_get_taskAssignment_by_LoginID]"
        Dim arrParameters As New ArrayList
        arrParameters.Add("@LoginID")
        arrParameters.Add("@Index")

        Dim arrTypes As New ArrayList
        arrTypes.Add(SqlDbType.Char)
        arrTypes.Add(SqlDbType.Char)

        Dim arrQueryString As New ArrayList
        arrQueryString.Add(loginID)
        arrQueryString.Add(Index)


        getAllTastAssignmentByLoginID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameters, arrTypes, arrQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getTastAssignmentCountByLoginID(loginID As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Task_get_taskAssignment_count_by_LoginID]"
        Dim strParameter As String = "@LoginID"
        Dim strType As String = SqlDbType.Char
        Dim strQueryString As String = loginID


        getTastAssignmentCountByLoginID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getTastAssignmentByID(ID As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Task_get_taskAssignment_by_ID]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID


        getTastAssignmentByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAllStaff() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Task_get_all_Users]"

        getAllStaff = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Mail"
    Public Function GetMailTaskAssignment(Value As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err


        Dim sp As String = "[Mail_get_sub_taskAssignment]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int

        GetMailTaskAssignment = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub UpdateMailTaskAssignment(value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_mail_sub_taskAssignment]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Parameter.Add("@OrderID")
        Parameter.Add("@subject")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)




Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
#End Region
#Region "Save "
    Public Sub InsertTaskAssignment(ByRef arrValue As ArrayList)
        On Error GoTo Err



        On Error GoTo Err

        Dim sp As String = "[insert_taskAssignment]"


        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@subject")
        arrParameter.Add("@task")
        arrParameter.Add("@startDate")
        arrParameter.Add("@assignedToStaffID")
        arrParameter.Add("@priority")
        arrParameter.Add("@active")
        arrParameter.Add("@status")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Date)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValue)


        Exit Sub




Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
    Public Sub UpdateTaskAssignment(ByRef arrValue As ArrayList)
        On Error GoTo Err



        On Error GoTo Err

        Dim sp As String = "[update_taskAssignment]"


        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        arrParameter.Add("@ID")
        arrParameter.Add("@subject")
        arrParameter.Add("@task")
        arrParameter.Add("@startDate")
        arrParameter.Add("@assignedToStaffID")
        arrParameter.Add("@priority")
        arrParameter.Add("@active")
        arrParameter.Add("@status")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Date)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValue)


        Exit Sub




Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
#End Region

End Class
