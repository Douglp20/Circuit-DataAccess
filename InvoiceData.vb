Public Class InvoiceData
    Public Sub New()

    End Sub
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Viper.Connection.Connection()
    Private connection As New Connection
#Region "Error Control"
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#End Region


#Region "Get Data  Invoice"
    Public Function getInvoiceByOrderIDList(OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_invoice_by_OrderID_List]"
        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = OrderID



        getInvoiceByOrderIDList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




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

    Public Function getUnInvoiceItemByOrderIDList(OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_UnInvoiced_OrderItem_by_OrderID_list]"
        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = OrderID



        getUnInvoiceItemByOrderIDList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderItemByInvoiceIDList(InvoiceID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_OrderItem_by_InvoiceID_List]"
        Dim strParameter As String = "@InvoiceID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = InvoiceID



        getOrderItemByInvoiceIDList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getInvoiceOrderCompanyByOrderYearList(arrValues As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_invoice_orders_Company_by_year_list]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        arrParameter.Add("@Year")
        arrParameter.Add("@Month")

        arrType.Add(SqlDbType.Char)
        arrType.Add(SqlDbType.Char)




        getInvoiceOrderCompanyByOrderYearList = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getInvoiceOrderByCompanyIDandYear(arrValues As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_invoice_orders_By_CompanyIDandYear]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        arrParameter.Add("@CompanyID")
        arrParameter.Add("@Year")
        arrParameter.Add("@Month")
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Char)
        arrType.Add(SqlDbType.Char)

        getInvoiceOrderByCompanyIDandYear = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function GetInvoiceByInvoiceNoList(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        sp = "[Invoice_get_invoice_InvoiceNo_List]"

        arrParameter.Add("@Index")
        arrParameter.Add("@InvoiceNo")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)


        GetInvoiceByInvoiceNoList = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)


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
