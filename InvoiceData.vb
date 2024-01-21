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
    Public Function getInvoiceEmailMessage(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_email_status_by_id]"

        Dim Parameter As String = "@ID"

        Dim Type As String = SqlDbType.Int

        getInvoiceEmailMessage = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

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

    Public Function getInvoiceOrderItemToBeInvoicedByOrderID(OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_OrderItem_ToBe_Invoice_by_OrderID]"
        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = OrderID



        getInvoiceOrderItemToBeInvoicedByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




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
    Public Function GetInvoiceBySearchInvoiceNo(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        sp = "[Invoice_get_invoice_Search_by_InvoiceNo]"

        Parameter.Add("@Index")
        Parameter.Add("@InvoiceNo")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)


        GetInvoiceBySearchInvoiceNo = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


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
    Public Function getInvoiceCompany(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_Invoice_company]"

        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID



        getInvoiceCompany = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderByCompanyID(arrValues As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_Orders_by_companyID]"

        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@companyID")
        arrParameter.Add("@OrderID")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)


        getOrderByCompanyID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save Data Invoice"
    Public Sub UpdateInvoiceNotes(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_Invoice_Notes]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@InvoiceID")
        Parameter.Add("@Notes")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Function CreateNewInvoice(ByRef arrValues As ArrayList) As Integer

        On Error GoTo Err

        Dim sp As String = "[insert_Invoice_Single_Invoice]"
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
    Public Function UpdateInvoiceEmaiStatus(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_invoice_email_action_by_ID]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@ID")
        Parameter.Add("@InvoiceEmailMessage")
        Parameter.Add("@InvoiceEmailStatus")
        Parameter.Add("@InvoiceEmailStatusID")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


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

    Public Sub UpdateBatchInvoiceNotes(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_BatchInvoice_Notes]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@AppID")
        Parameter.Add("@Notes")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub

    Public Function getInvoiceDataByInvoiceID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_invoice_By_ID]"

        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int



        getInvoiceDataByInvoiceID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
End Class
