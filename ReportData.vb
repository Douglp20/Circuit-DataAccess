Public Class ReportData
    Public Sub New()
    End Sub

    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New douglas.Viper.Connection.Connection()
    Private connection As New Connection

#Region "Error Control"
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#End Region
#Region "Order"
    Public Function getWorkTicketbyOrderID(id As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_work_ticket_by_order_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getWorkTicketbyOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getPicturebyOrderID(id As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_order_picture_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getPicturebyOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderSearch(value As String) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_get_order_by_search]"
        Dim Parameter As String = "@PassWhere"
        Dim Type As String = SqlDbType.VarChar



        getOrderSearch = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
    Public Function getInvoiceSheetbyInvoiceID(id As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_invoice_By_ID]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getInvoiceSheetbyInvoiceID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanyInvoiceStatement(value As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_company_invoice_statement]"
        Dim Parameter As String = "@companyID"
        Dim Type As String = SqlDbType.Int







        getCompanyInvoiceStatement = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanyInvoiceBatchStatement(value As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_company_invoice_batch_statement]"
        Dim Parameter As String = "@companyID"
        Dim Type As String = SqlDbType.Int







        getCompanyInvoiceBatchStatement = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getInvoicePaymentSheetbyInvoiceID(id As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_invoice_Payment_By_ID]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getInvoicePaymentSheetbyInvoiceID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getInvoiceBatchSheetbyInvoiceID(id As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_batch_invoice_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getInvoiceBatchSheetbyInvoiceID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getPayrollByPayrollID(id As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_Payroll_get_payroll]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getPayrollByPayrollID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAppointmentByDate(appDate As String) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_get_appointment_by_date]"
        Dim strParameter As String = "@date"
        Dim strType As String = SqlDbType.VarChar
        Dim strQueryString As String = appDate


        getAppointmentByDate = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#Region "Daily Report"
    Public Function getDailyReportbyIndex(value As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_daily_statement_by_index]"
        Dim Parameter As String = "@Index"
        Dim Type As String = SqlDbType.Int



        getDailyReportbyIndex = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

#End Region
#Region "Get Timesheet Data"
    Public Function getTimesheetUser(value As ArrayList) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[REPORT_get_timesheet_by_year]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Parameter.Add("@StaffID")
        Parameter.Add("@Year")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)

        getTimesheetUser = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region

#Region "Wholesaler"
    Public Function getSubContractortMaterial(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[REPORT_get_wholesaler_by_subContractor]"
        Dim Parameter As New ArrayList
        Parameter.Add("@StaffID")
        Parameter.Add("@Year")
        Parameter.Add("@Month")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)

        getSubContractortMaterial = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
End Class
