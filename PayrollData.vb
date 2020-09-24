

Public Class PayrollData

    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub



#Region "Get Data Payroll"
    Public Function getPayrollYearMonth() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim ViperCon As New Viper.Connection.Connection
        Dim sp As String = "[Payroll_get_payroll_year_month]"


        getPayrollYearMonth = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Function getPayrollBySearch(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim ViperCon As New Viper.Connection.Connection
        Dim sp As String = "[Payroll_get_payroll_by_search]"


        Dim arrParameter As New ArrayList
        arrParameter.Add("@year")
        arrParameter.Add("@Month")
        arrParameter.Add("@StaffID")

        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Int)
        getPayrollBySearch = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getPayrollByID(ByRef ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim ViperCon As New Viper.Connection.Connection
        Dim sp As String = "[Payroll_get_payroll_by_ID]"

        Dim QueryString As String = ID
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int

        getPayrollByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, QueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getBACSBalanceByStaffID(ByRef ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim ViperCon As New Viper.Connection.Connection
        Dim sp As String = "[Payroll_get_payroll_BACS_balance]"

        Dim QueryString As String = ID
        Dim strParameter As String = "@StaffID"
        Dim strType As String = SqlDbType.Int

        getBACSBalanceByStaffID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, QueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

#Region "Get Data Staff"
    Public Function getPayrollStaff() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim ViperCon As New Viper.Connection.Connection
        Dim sp As String = "[Payroll_get_staff_payroll]"


        getPayrollStaff = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getAllStaffList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Staff_get_all_staff_list]"

        getAllStaffList = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getStaffbyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_get_staff_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getStaffbyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderItemAssignedByStaffID(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim ViperCon As New Viper.Connection.Connection
        Dim sp As String = "[Payroll_get_orderItem_assigned_by_StaffID]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@StaffID")
        arrParameter.Add("@year")
        arrParameter.Add("@Month")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)

        getOrderItemAssignedByStaffID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrQueryString)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getOrderItemAvailableByStaffID(ByRef ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim ViperCon As New Viper.Connection.Connection
        Dim sp As String = "[Payroll_get_orderItem_available_by_StaffID]"

        Dim QueryString As String = ID
        Dim strParameter As String = "@StaffID"
        Dim strType As String = SqlDbType.Int

        getOrderItemAvailableByStaffID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, QueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#Region "material Coding"

    Public Function getMaterialAssignedByStaffID(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim ViperCon As New Viper.Connection.Connection
        Dim sp As String = "[Payroll_get_material_assigned_by_StaffID]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@StaffID")
        arrParameter.Add("@year")
        arrParameter.Add("@Month")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)

        getMaterialAssignedByStaffID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrQueryString)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getMaterialAvailableByStaffID(ByRef ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim ViperCon As New Viper.Connection.Connection
        Dim sp As String = "[Payroll_get_Material_available_by_StaffID]"

        Dim QueryString As String = ID
        Dim strParameter As String = "@StaffID"
        Dim strType As String = SqlDbType.Int

        getMaterialAvailableByStaffID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, QueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#End Region
#Region "Save Data"

#End Region
End Class
