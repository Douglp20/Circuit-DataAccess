Public Class InvoiceBatchData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub

#Region "Get Data"
    Public Function getAllBatchCompany() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_all_batch_company]"


        getAllBatchCompany = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAllBatchByCompanyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_all_batch_by_companyID]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int

        getAllBatchByCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, id)



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
    Public Function getOrderContractList(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_company_Order_contract_list]"



        Dim strParameter As String = "@id"

        Dim strType As String = SqlDbType.VarChar


        getOrderContractList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, id)


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
    Public Function getContractNoOrders(contractNo As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Order_by_contractNo]"



        Dim strParameter As String = "@contractNo"

        Dim strType As String = SqlDbType.VarChar


        getContractNoOrders = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, contractNo)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

#End Region

#Region "Save Data"
    Public Function SaveBatchGetBatchID(ByRef arrQueryString As ArrayList) As Integer
        On Error GoTo Err

        Dim sp As String = "[insert_batchInvoice]"
        Dim strParameterOutput As String = "@ID"


        Dim arrParameter As New ArrayList

        arrParameter.Add("@Notes")
        arrParameter.Add("@UserName")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)


        SaveBatchGetBatchID = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString(), sp, arrParameter, strParameterOutput, arrType, arrQueryString)


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
    Public Function SaveInvoiceByBatchID(ByRef arrQueryString As ArrayList) As Integer
        On Error GoTo Err

        Dim sp As String = "[insert_Invoice_Batch_Invoice]"
        Dim strParameterOutput As String = "@ID"


        Dim arrParameter As New ArrayList

        arrParameter.Add("@invoicebatchID")
        arrParameter.Add("@UserName")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)


        SaveInvoiceByBatchID = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString(), sp, arrParameter, strParameterOutput, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub AssignOrderToInvoiceBatch(ByRef arrQueryString As ArrayList)
        On Error GoTo Err

        Dim sp As String = "[insert_Invoice_Batch_Invoice_Order]"


        Dim arrParameter As New ArrayList

        arrParameter.Add("@invoicebatchID")
        arrParameter.Add("@invoiceID")
        arrParameter.Add("@orderID")
        arrParameter.Add("@UserName")

        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrQueryString)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
#End Region
End Class
