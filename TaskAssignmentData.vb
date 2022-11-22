Public Class TaskAssignmentData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
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
    Public Function getTaskInfo(Value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_certificate_assignment_by_OrderID]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int

        getTaskInfo = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, Value)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getTaskAdminSubContractor(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Task_get_assignee_order_certificate]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int



        getTaskAdminSubContractor = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getAllTaskAssignment(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Task_get_all_taskAssignment]"
        Dim Parameter As String = "@Index"
        Dim Type As String = SqlDbType.Int
        getAllTaskAssignment = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getAllTaskAssignmentByLoginID(loginID As String, Index As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Task_get_taskAssignment_by_LoginID]"
        Dim Parameter As New ArrayList
        Parameter.Add("@LoginID")
        Parameter.Add("@Index")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.Char)
        Type.Add(SqlDbType.Char)

        Dim value As New ArrayList
        value.Add(loginID)
        value.Add(Index)


        getAllTaskAssignmentByLoginID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getTaskAssignmentCountByLoginID(loginID As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Task_get_taskAssignment_count_by_LoginID]"
        Dim strParameter As String = "@LoginID"
        Dim strType As String = SqlDbType.Char
        Dim strQueryString As String = loginID


        getTaskAssignmentCountByLoginID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getTastAssignmentByID(value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Task_get_taskAssignment_by_ID]"
        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int



        getTastAssignmentByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



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
    '#Region "Mail"
    '    Public Function GetMailTaskAssignment(Value As Integer) As SqlClient.SqlDataAdapter
    '        On Error GoTo Err


    '        Dim sp As String = "[task_get_mail_sub_taskAssignment]"
    '        Dim Parameter As String = "@OrderID"
    '        Dim Type As String = SqlDbType.Int

    '        GetMailTaskAssignment = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, Value)


    '        Exit Function

    'Err:
    '        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
    '        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    '    End Function
    '    Public Sub UpdateMailTaskAssignment(value As ArrayList)
    '        On Error GoTo Err


    '        Dim sp As String = "[update_task_mail_sub_taskAssignment]"
    '        Dim Parameter As New ArrayList
    '        Dim Type As New ArrayList
    '        Parameter.Add("@OrderID")
    '        Parameter.Add("@subject")

    '        Type.Add(SqlDbType.Int)
    '        Type.Add(SqlDbType.VarChar)



    '        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)




    'Err:
    '        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
    '        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    '    End Sub
    '#End Region
#Region "Save "
    Public Sub InsertTaskAssignment(ByRef Value As ArrayList)
        On Error GoTo Err



        On Error GoTo Err

        Dim sp As String = "[insert_taskAssignment]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@subject")
        Parameter.Add("@task")
        Parameter.Add("@startDate")
        Parameter.Add("@assignedToStaffID")
        Parameter.Add("@priority")
        Parameter.Add("@active")
        Parameter.Add("@tasktype")
        Parameter.Add("@status")
        Parameter.Add("@OrderID")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub




Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
    Public Sub UpdateTaskAssignment(ByRef Value As ArrayList)
        On Error GoTo Err



        On Error GoTo Err

        Dim sp As String = "[update_taskAssignment]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@subject")
        Parameter.Add("@task")
        Parameter.Add("@startDate")
        Parameter.Add("@assignedToStaffID")
        Parameter.Add("@priority")
        Parameter.Add("@active")
        Parameter.Add("@status")
        Parameter.Add("@OrderID")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub




Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
#End Region

End Class
