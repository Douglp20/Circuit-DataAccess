

Public Class PayrollData

    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub



#Region "Get Data Payroll"
    Public Function getPayrollYear() As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Payroll_get_payroll_year]"


        getPayrollYear = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getPayrollYearMonth(value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Payroll_get_payroll_month_by_year]"

        Dim Parameter As String = "@year"
        Dim Type As String = SqlDbType.VarChar

        getPayrollYearMonth = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Function getPayrollBySearch(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


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
    Public Function getPayrollByID(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Payroll_get_payroll_by_ID]"

        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int

        getPayrollByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getBACSBalanceByStaffID(ByRef ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


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


        Dim sp As String = "[Payroll_get_open_payroll_staff]"


        getPayrollStaff = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getAllStaffList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Payroll_get_all_staff]"

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
    Public Function zzzgetOrderItemAssignedByStaffID(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Payroll_get_orderItem_assigned_by_StaffID]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@StaffID")
        arrParameter.Add("@year")
        arrParameter.Add("@Month")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)

        '   getOrderItemAssignedByStaffID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrQueryString)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getOrderItemByStaffID(ByRef value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Payroll_get_orderItem_by_StaffID]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@StaffID")
        Parameter.Add("@PayrollID")
        Parameter.Add("@Index")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)

        getOrderItemByStaffID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#Region "material Coding"

    Public Function zzzgetMaterialAssignedByStaffID(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Payroll_get_material_assigned_by_StaffID]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@StaffID")
        arrParameter.Add("@year")
        arrParameter.Add("@Month")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)

        'getMaterialAssignedByStaffID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrQueryString)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getMaterialByStaffID(ByRef value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Payroll_get_material_by_StaffID]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@StaffID")
        Parameter.Add("@PayrollID")
        Parameter.Add("@Index")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)

        getMaterialByStaffID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getHourByStaffID(ByRef value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Payroll_get_hour_by_StaffID]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@StaffID")
        Parameter.Add("@PayrollID")
        Parameter.Add("@Index")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)

        getHourByStaffID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#End Region
#Region "Save Data"


    Public Function InsertPayroll(ByRef arrQueryString As ArrayList) As Integer
        On Error GoTo Err


        Dim sp As String = "[insert_Payroll]"
        Dim strParameterOutput As String = "@ID"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@StaffID")
        arrParameter.Add("@Month")
        arrParameter.Add("@Year")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)




        InsertPayroll = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString(), sp, arrParameter, strParameterOutput, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function UpdatePayroll(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_Payroll]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@id")
        arrParameter.Add("@Staffid")
        arrParameter.Add("@Month")
        arrParameter.Add("@Year")
        arrParameter.Add("@notes")
        arrParameter.Add("@Hour_worked")
        arrParameter.Add("@Hour_Added")
        arrParameter.Add("@Hourly_Rate")
        arrParameter.Add("@Hourly_total")
        arrParameter.Add("@material_les")
        arrParameter.Add("@material_om")
        arrParameter.Add("@BACS")
        arrParameter.Add("@OrderTotal")
        arrParameter.Add("@TAX")
        arrParameter.Add("@NI")
        arrParameter.Add("@BACS_Outstanding")
        arrParameter.Add("@BACS_Paid")
        arrParameter.Add("@deduction")
        arrParameter.Add("@addition")
        arrParameter.Add("@closed")
        arrParameter.Add("@UserName")


        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function deleteOrderAssignment(ByRef id As Integer)

        On Error GoTo Err

        Dim sp As String = "[Delete_payroll_orderItem_Assignment]"
        Dim strParameter As String = "PayrollID"
        Dim strType As String = SqlDbType.Int


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString(), sp, strParameter, strType, id)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateOrderAssignment(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[Insert_payroll_orderItem_Assignment]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList




        arrParameter.Add("@PayrollID")
        arrParameter.Add("@Staffid")
        arrParameter.Add("@OrderItemID")
        arrParameter.Add("@UserName")



        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function deleteMaterialAssignment(ByRef id As Integer)

        On Error GoTo Err

        Dim sp As String = "[Delete_payroll_material_assignment]"
        Dim strParameter As String = "PayrollID"
        Dim strType As String = SqlDbType.Int


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString(), sp, strParameter, strType, id)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateMaterialAssignment(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[Insert_payroll_material_Assignment]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@PayrollID")
        arrParameter.Add("@Staffid")
        arrParameter.Add("@MaterialID")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function deleteHourAssignment(ByRef id As Integer)

        On Error GoTo Err

        Dim sp As String = "[Delete_payroll_timesheet_assignment]"
        Dim strParameter As String = "PayrollID"
        Dim strType As String = SqlDbType.Int


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString(), sp, strParameter, strType, id)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateHourAssignment(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[Insert_payroll_timesheet_assignment]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@PayrollID")
        arrParameter.Add("@Staffid")
        arrParameter.Add("@timesheetID")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
End Class
