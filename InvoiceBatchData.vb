Public Class InvoiceBatchData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub
#Region "Error"
    Private Sub ErrorMessage_ViperCon(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#End Region




#Region "Get Data"
    Public Function getBatchedCompanyByYear(arrValues As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Batched_Company_by_year]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        arrParameter.Add("@Year")
        arrParameter.Add("@Month")

        arrType.Add(SqlDbType.Char)
        arrType.Add(SqlDbType.Char)




        getBatchedCompanyByYear = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function zzzgetAllBatchedOrders() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_batched_companyID]"


        zzzgetAllBatchedOrders = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAvailableBatchCompany() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_available_batch_company]"


        getAvailableBatchCompany = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getBatchesByCompanyID(Values As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_batches_by_companyID]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@CompanyID")
        Parameter.Add("@Year")
        Parameter.Add("@Month")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Char)
        Type.Add(SqlDbType.Char)


        getBatchesByCompanyID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)





        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderByBatchID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Order_by_batchID]"
        Dim strParameter As String = "@BatchID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        getOrderByBatchID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAvailableBatchByCompanyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_available_batch_by_companyID]"



        Dim strParameter As String = "@id"

        Dim strType As String = SqlDbType.VarChar


        getAvailableBatchByCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, id)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getInvoiceBatchByBatchID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_InvoiceBatch_by_batchID]"
        Dim strParameter As String = "@InvoiceBatchID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        getInvoiceBatchByBatchID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getApplicationByBatchID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_application_By_ID]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        getApplicationByBatchID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getApplicationOrderByBatchID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Order_application_by_batchID]"
        Dim strParameter As String = "@batchID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        getApplicationOrderByBatchID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getContractNoOrders(arrValue As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Order_by_contractNo]"
        Dim Parameter As New ArrayList

        Parameter.Add("@contractNo")
        Parameter.Add("@OrderRun")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)




        getContractNoOrders = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, Parameter, Type, arrValue)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getSelectedContractNoOrders(Value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Order_select_by_contractNo]"
        Dim Parameter As String = "@WHERE"
        Dim Type As String = SqlDbType.VarChar





        getSelectedContractNoOrders = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderItemByOrderID(OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_orderitem_by_OrderID]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int



        getOrderItemByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, OrderID)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function GetNewBatchInfo(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_new_batch_invoice]"
        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int



        GetNewBatchInfo = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, ID)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

#End Region

#Region "Save Data"
    Public Function InsertNewInvoiceBatch(ByRef arrQueryString As ArrayList) As Integer
        On Error GoTo Err

        Dim sp As String = "[insert_Invoice_Batch_Invoice]"
        Dim strParameterOutput As String = "@ID"


        Dim arrParameter As New ArrayList

        arrParameter.Add("@InvoiceNotes")
        arrParameter.Add("@UserName")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)


        InsertNewInvoiceBatch = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString(), sp, arrParameter, strParameterOutput, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub UpdateBatchByBatchID(ByRef arrQueryString As ArrayList)
        On Error GoTo Err

        Dim sp As String = "[Update_batchInvoice]"


        Dim arrParameter As New ArrayList

        arrParameter.Add("@ID")
        arrParameter.Add("@Notes")
        arrParameter.Add("@UserName")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrQueryString)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Function ZZZSaveInvoiceByBatchID(ByRef arrQueryString As ArrayList) As Integer
        On Error GoTo Err

        Dim sp As String = "[insert_Invoice_Batch_Invoice]"
        Dim strParameterOutput As String = "@ID"


        Dim arrParameter As New ArrayList

        arrParameter.Add("@invoicebatchID")
        arrParameter.Add("@UserName")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)


        '' SaveInvoiceByBatchID = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString(), sp, arrParameter, strParameterOutput, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub AssignOrderToInvoiceBatch(ByRef value As ArrayList)
        On Error GoTo Err

        Dim sp As String = "[insert_Invoice_Batch_Invoice_Order]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@invoiceID")
        Parameter.Add("@orderID")
        Parameter.Add("@UserName")



        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
#End Region
#Region "Search"
    Public Function GetInvoiceByInvoiceNo(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        sp = "[Batch_get_invoice_by_InvoiceNo]"

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
    Public Function GetInvoiceByBatchNo(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        sp = "[Batch_get_invoice_by_BatchNo]"

        Parameter.Add("@Index")
        Parameter.Add("@BatchNo")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)


        GetInvoiceByBatchNo = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
End Class
