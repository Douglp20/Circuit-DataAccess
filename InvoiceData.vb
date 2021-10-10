Public Class InvoiceData
    Public Sub New()

    End Sub
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
    Private connection As New Connection
#Region "Error Control"
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#End Region


#Region "Get Data  Invoice"
    Public Function getInvoiceByOrderID(OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_invoice_by_OrderID]"
        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = OrderID



        getInvoiceByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getInvoiceOrderDataByOrderID(OrderID As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err

        Dim sp As String = "[Invoice_get_order_data_by_id]"
        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = OrderID



        getInvoiceOrderDataByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getOrderItemWaitingInvocingByOrder(OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_OrderItem_waiting_invocing_by_OrderID]"
        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = OrderID



        getOrderItemWaitingInvocingByOrder = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderItemByInvoiceIDList(InvoiceID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_OrderItem_by_InvoiceID]"
        Dim strParameter As String = "@InvoiceID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = InvoiceID



        getOrderItemByInvoiceIDList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getInvoicedCompanyByYear(arrValues As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_Invoiced_Company_by_year]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        arrParameter.Add("@Year")
        arrParameter.Add("@Month")

        arrType.Add(SqlDbType.Char)
        arrType.Add(SqlDbType.Char)




        getInvoicedCompanyByYear = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getInvoicedByCompanyID(Values As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_Invoiced_by_companyID]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@CompanyID")
        Parameter.Add("@Year")
        Parameter.Add("@Month")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Char)
        Type.Add(SqlDbType.Char)

        getInvoicedByCompanyID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function GetInvoiceByInvoiceNo(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        sp = "[Invoice_get_invoice_by_InvoiceNo]"

        Parameter.Add("@Index")
        Parameter.Add("@InvoiceNo")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)


        GetInvoiceByInvoiceNo = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Get Data BulkInvoice"
    Public Function getCompanyCompletedOrderByYear(year As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_uninvoice_company_order_by_year]"

        Dim strParameter As String = "@Year"
        Dim strType As String = SqlDbType.Char
        Dim strQueryString As String = year



        getCompanyCompletedOrderByYear = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getUnInvoiceCompletedOrderByCompanyIDandYear(arrValues As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_uninvoice_order_By_CompanyIDandYear]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        arrParameter.Add("@CompanyID")
        arrParameter.Add("@Year")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Char)

        getUnInvoiceCompletedOrderByCompanyIDandYear = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValues)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanyCompletedOrderByOrderNo(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        sp = "[Order_get_order_company_by_OrderNo]"

        arrParameter.Add("@Index")
        arrParameter.Add("@OrderNo")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)


        getCompanyCompletedOrderByOrderNo = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanyCompletedOrderByJobNo(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        sp = "[Order_get_order_company_by_JobNo]"

        arrParameter.Add("@Index")
        arrParameter.Add("@JObNo")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)


        getCompanyCompletedOrderByJobNo = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

#End Region

#Region "Orders waiting to be Invoiced"
    Public Function getCompanyOrderToBeInvoiced(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_company_waiting_Invoicing]"

        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID



        getCompanyOrderToBeInvoiced = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderToBeInvoicedByCompanyID(arrValues As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_order_waiting_invoicing_by_companyID]"

        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@companyID")
        arrParameter.Add("@OrderID")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)


        getOrderToBeInvoicedByCompanyID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save Data Invoice"
    Public Function CreateNewInvoice(ByRef arrValues As ArrayList) As Integer

        On Error GoTo Err

        Dim sp As String = "[insert_Invoice_CreateANewSingleInvoice]"
        Dim strParameterOutput As String = "@ID"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@OrderID")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)



        CreateNewInvoice = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString, sp, arrParameter, strParameterOutput, arrType, arrValues)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function AssignInvoiceOrderItem(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        sp = "[Update_Invoice_OrderItem_Assign]"
        arrParameter.Add("@Index")
        arrParameter.Add("@OrderID")
        arrParameter.Add("@InvoiceID")
        arrParameter.Add("@OrderItemID")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save Invoice Bulk"
    Public Function GenerateNewInvoice(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        sp = "[insert_Invoice_Bulk_Invoice]"

        arrParameter.Add("@OrderID")
        arrParameter.Add("@UserName")

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
