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
#Region "Application Stage 1"
    Public Function getBatchedApplicationStage1Company() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_Stage1_company]"



        getBatchedApplicationStage1Company = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Sub DeleteApplicationImportStage1(value As Integer)

        On Error GoTo Err

        Dim sp As String = "[Delete_Batch_Application_Stage1_by_Company]"

        Dim Parameter As String = "@companyID"
        Dim Type As String = SqlDbType.Int



        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub ProcessApplicationStage1ImportByCompanyID(value As Integer)

        On Error GoTo Err

        Dim sp As String = "[Update_batch_application_stage1_imported_data_by_companyID]"

        Dim Parameter As String = "@companyID"
        Dim Type As String = SqlDbType.Int



        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub InsertApplicationImportStage1(value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[Insert_Batch_Application_Import_Stage1]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@companyID")
        Parameter.Add("@OrderNumber")
        Parameter.Add("@refNumber")
        Parameter.Add("@appTotal")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Money)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Function GetAppBatchStage1Info(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_Stage1_newNumber]"
        Dim Parameter As String = "@companyID"
        Dim Type As String = "SqlDbType.Int"


        GetAppBatchStage1Info = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getBatchedApplicationStage1OrdersByOrderID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_stage1_order_data_by_OrderID]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int



        getBatchedApplicationStage1OrdersByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getBatchedApplicationStage1OrdersByCompany(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_stage1_order_data_by_companyID]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@companyID")
        Parameter.Add("@Index")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)

        getBatchedApplicationStage1OrdersByCompany = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getApplicationStage1OrderImportedDataByCompanyID(Value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_application_stage1_Order_imported_data_by_companyID]"
        Dim Parameter As String = "@companyID"

        Dim Type As String = "SqlDbType.int"


        getApplicationStage1OrderImportedDataByCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Function InsertBatchAppStage1NewNumber(ByRef value As ArrayList) As Integer
        On Error GoTo Err

        Dim sp As String = "[insert_Batch_Invoice_App_Stage1_NewNumber]"
        Dim strParameterOutput As String = "@ID"


        Dim Parameter As New ArrayList

        Parameter.Add("@InvoiceNumber")
        Parameter.Add("@batchNumber")
        Parameter.Add("@InvoiceNotes")
        Parameter.Add("@UserName")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)

        InsertBatchAppStage1NewNumber = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString(), sp, Parameter, strParameterOutput, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub UpdatebatchAppStage1OrderItemAssignment(ByRef value As ArrayList)
        On Error GoTo Err

        Dim sp As String = "[Update_Batch_Application_Stage1_OrderItem_assignment]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@orderID")
        Parameter.Add("@invoiceID")
        Parameter.Add("@appID")
        Parameter.Add("@UserName")



        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub UpdateBatchApplicationStage1OrderItemFeedback(ByRef value As ArrayList)
        On Error GoTo Err

        Dim sp As String = "[Update_Batch_Application_Stage1_OrderItem_feedback]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@action")
        Parameter.Add("@reject")
        Parameter.Add("@reason")
        Parameter.Add("@UserName")



        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Function InsertBatchApplicationStage1NewNumber(ByRef value As ArrayList) As Integer
        On Error GoTo Err

        Dim sp As String = "[insert_Batch_Application_Stage1_NewNumber]"
        Dim strParameterOutput As String = "@ID"


        Dim Parameter As New ArrayList


        Parameter.Add("@AppNumber")
        Parameter.Add("@AppNumberSub")
        Parameter.Add("@InvoiceNotes")
        Parameter.Add("@UserName")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        InsertBatchApplicationStage1NewNumber = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString(), sp, Parameter, strParameterOutput, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region

#Region "Application Stage 2"
    Public Function getBatchedApplicationStage2Company() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_Stage2_company]"



        getBatchedApplicationStage2Company = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function GetAppBatchStage2Info(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_Stage2_newNumber]"
        Dim Parameter As String = "@companyID"
        Dim Type As String = "SqlDbType.Int"


        GetAppBatchStage2Info = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Function getBatchedApplicationStage2OrdersByOrderID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_stage2_order_data_by_OrderID]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int



        getBatchedApplicationStage2OrdersByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function InsertBatchAppStage2NewNumber(ByRef value As ArrayList) As Integer
        On Error GoTo Err

        Dim sp As String = "[insert_Batch_Invoice_App_Stage2_NewNumber]"
        Dim strParameterOutput As String = "@ID"

        Dim Parameter As New ArrayList

        Parameter.Add("@InvoiceNumber")
        Parameter.Add("@batchNumber")
        Parameter.Add("@InvoiceNotes")
        Parameter.Add("@UserName")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        InsertBatchAppStage2NewNumber = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString(), sp, Parameter, strParameterOutput, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function



    Public Sub UpdatebatchAppStage2OrderItemAssignment(ByRef value As ArrayList)
        On Error GoTo Err

        Dim sp As String = "[Update_Batch_Application_Stage2_OrderItem_assignment]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@orderID")
        Parameter.Add("@invoiceID")
        Parameter.Add("@appID")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Function InsertBatchApplicationStage2NewNumber(ByRef value As ArrayList) As Integer
        On Error GoTo Err

        Dim sp As String = "[insert_Batch_Application_Stage2_NewNumber]"
        Dim strParameterOutput As String = "@ID"


        Dim Parameter As New ArrayList


        Parameter.Add("@batchAppNumber")
        Parameter.Add("@InvoiceNotes")
        Parameter.Add("@UserName")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        InsertBatchApplicationStage2NewNumber = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString(), sp, Parameter, strParameterOutput, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getApplicationStage2OrderContractByCompanyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_stage2_Order_contract_by_companyID]"



        Dim strParameter As String = "@id"

        Dim strType As String = SqlDbType.VarChar


        getApplicationStage2OrderContractByCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, id)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getApplicationStage2OrderByContractNo(Value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_application_stage2_Order_by_contractNo]"
        Dim Parameter As String = "@contractNo"

        Dim Type As String = "SqlDbType.VarChar"


        getApplicationStage2OrderByContractNo = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region


#Region "Application Stage 3"
    Public Function getBatchedApplicationStage3Company() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_Stage3_company]"



        getBatchedApplicationStage3Company = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getBatchedApplicationStage3POByCompanyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_application_stage3_PO_by_companyID]"
        Dim strParameter As String = "@CompanyID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        getBatchedApplicationStage3POByCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getBatchedApplicationStage3OrderByAppBatchID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_Stage3_Order_by_appBatchID]"
        Dim strParameter As String = "@appBatchID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        getBatchedApplicationStage3OrderByAppBatchID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getBatchedApplicationStage3OrdersByOrderID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_stage3_order_data_by_OrderID]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int



        getBatchedApplicationStage3OrdersByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getStage3OrderItemByOrderID(OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_Stage3_orderitem_by_OrderID]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int



        getStage3OrderItemByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, OrderID)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function GetAppBatchStage3Info(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_Stage3_newNumber]"




        Dim Parameter As String = "@appBatchID"
        Dim Type As String = SqlDbType.Int


        GetAppBatchStage3Info = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function InsertBatchAppStage3NewNumber(ByRef value As ArrayList) As Integer
        On Error GoTo Err

        Dim sp As String = "[insert_Batch_Invoice_App_Stage3_NewNumber]"
        Dim strParameterOutput As String = "@ID"

        Dim Parameter As New ArrayList

        Parameter.Add("@InvoiceNumber")
        Parameter.Add("@batchNumber")
        Parameter.Add("@InvoiceNotes")
        Parameter.Add("@UserName")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        InsertBatchAppStage3NewNumber = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString(), sp, Parameter, strParameterOutput, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function InsertBatchApplicationStage3NewNumber(ByRef value As ArrayList) As Integer
        On Error GoTo Err

        Dim sp As String = "[insert_Batch_Application_Stage3_NewNumber]"
        Dim strParameterOutput As String = "@ID"


        Dim Parameter As New ArrayList


        Parameter.Add("@AppNumber")
        Parameter.Add("@AppNumberSub")
        Parameter.Add("@InvoiceNotes")
        Parameter.Add("@UserName")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        InsertBatchApplicationStage3NewNumber = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString(), sp, Parameter, strParameterOutput, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub UpdatebatchAppStage3OrderItemAssignment(ByRef value As ArrayList)
        On Error GoTo Err

        Dim sp As String = "[Update_Batch_Application_Stage3_OrderItem_assignment]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@orderID")
        Parameter.Add("@invoiceID")
        Parameter.Add("@appID")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.Int)
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



#Region "Get Data"
    Public Function getInvoiceIDbyAppID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_invoiceID_By_appID]"
        Dim strParameter As String = "@appID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        getInvoiceIDbyAppID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

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


    Public Function getBatchedApplicationByCompanyIDYearMonth(Values As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_application_by_companyIDYearMonth]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@CompanyID")
        Parameter.Add("@Year")
        Parameter.Add("@Month")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Char)
        Type.Add(SqlDbType.Char)


        getBatchedApplicationByCompanyIDYearMonth = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)





        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderByAppBatchID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Order_by_appBatchID]"
        Dim strParameter As String = "@appBatchID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        getOrderByAppBatchID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function


    Public Function getInvoiceBatchByBatchAppID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_InvoiceBatch_Info_by_AppID]"
        Dim strParameter As String = "@appBatchID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        getInvoiceBatchByBatchAppID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)




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
    Public Function getOrderExcelByAppBatchID(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_order_excel_by_appID]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@AppBatchID")
        Parameter.Add("@ExcelRef")



        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)

        getOrderExcelByAppBatchID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getApplicationHistricalPictureByAppBatchID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[BATCH_get_Applicatopn_histrical_Picture_by_ID]"
        Dim Parameter As String = "@AppBatchID"
        Dim Type As String = SqlDbType.Int


        getApplicationHistricalPictureByAppBatchID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderExcelHeaderByAppBatchID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_order_excel_header_by_appID]"
        Dim strParameter As String = "@AppBatchID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        getOrderExcelHeaderByAppBatchID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getContractNo_NoPOOrders(Value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Order_by_contractNo_WithNoPONo]"
        Dim Parameter As String = "@contractNo"

        Dim Type As String = "SqlDbType.VarChar"


        getContractNo_NoPOOrders = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, Value)


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
    Public Function GetApplicationSearchByInvoiceNo(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        sp = "[Batch_get_invoice_Search_by_InvoiceNo]"

        Parameter.Add("@Index")
        Parameter.Add("@InvoiceNo")
        Parameter.Add("@AppNo")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)

        GetApplicationSearchByInvoiceNo = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    '    Public Function GetInvoiceByBatchNo(ByRef Values As ArrayList)

    '        On Error GoTo Err

    '        Dim sp As String
    '        Dim arrValuesPass As New ArrayList

    '        Dim Parameter As New ArrayList
    '        Dim Type As New ArrayList

    '        sp = "[Batch_get_invoice_Search_by_batchNo]"

    '        Parameter.Add("@Index")
    '        Parameter.Add("@BatchNo")

    '        Type.Add(SqlDbType.Int)
    '        Type.Add(SqlDbType.Int)


    '        GetInvoiceByBatchNo = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


    '        Exit Function

    'Err:
    '        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
    '        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    '    End Function
#End Region

#Region "Old Access Batch Process"

    Public Function getAccessBatchCompany() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Access_batch_company]"


        getAccessBatchCompany = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAccessBatchPOByCompanyID(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Access_batch__PO_by_companyID]"

        Dim Parameter As New ArrayList
        Parameter.Add("@companyID")
        Parameter.Add("@Index")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)




        getAccessBatchPOByCompanyID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAccessContractAndPOOrders(arrValue As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Access_Order_by_contractNo_PONo]"
        Dim Parameter As New ArrayList

        Parameter.Add("@contractNo")
        Parameter.Add("@PONumber")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)




        getAccessContractAndPOOrders = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, Parameter, Type, arrValue)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getBatchAccessOrdersByOrderID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Access_order_data_by_OrderID]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int



        getBatchAccessOrdersByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
End Class
