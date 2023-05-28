﻿
Public Class OrderItemData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#Region "Error Control"

#End Region


#Region "SubContractor Assignments"

    Public Function getOrderSubContractorAssignments(value As ArrayList) As SqlClient.SqlDataAdapter
        On Error GoTo Err

        Dim sp As String = "[Order_get_SubContractor_assignments]"
        Dim Parameter As New ArrayList

        Parameter.Add("@Year")
        Parameter.Add("@Month")
        Parameter.Add("@StaffID")
        Parameter.Add("@JobNumber")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)

        getOrderSubContractorAssignments = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function
#End Region

#Region "Get Data"
    Public Function getBatchAppOrderItemDataOrderID(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Batch_get_Application_Stage_correction_orderitem_by_OrderID]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@OrderID")
        Parameter.Add("@Index")

        Type.Add(SqlDbType.Int)
        Type.add(SqlDbType.Int)




        getBatchAppOrderItemDataOrderID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function
    Public Function getOrderItemDataOrderID(OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderitem_by_OrderID]"

        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = OrderID


        getOrderItemDataOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function
    Public Function getOrderQuotationItemDataOrderIDList(OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderQuotation_item_by_OrderID_List]"

        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = OrderID


        getOrderQuotationItemDataOrderIDList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function
    Public Function getOrderItemDataByOrderID(OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_Order_orderitem_by_OrderID]"

        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = OrderID


        getOrderItemDataByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function
    Public Function getOrderQuotationItemDataOrderID(OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderitem_quotation_by_OrderID]"

        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = OrderID


        getOrderQuotationItemDataOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function
    Public Function getAllSubContractorList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Staff_get_all_subContractor_list]"

        getAllSubContractorList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getCompanyProjectPriceList(arrValue As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[ORDER_get_company_project_pricelist]"
        Dim arrParameter As New ArrayList

        arrParameter.Add("@companyID")
        arrParameter.Add("@projectID")

        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)

        getCompanyProjectPriceList = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValue)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanyProjectPricelistbySearch(ByRef arrValue As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_company_project_pricelist_search]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@companyID")
        arrParameter.Add("@ProjectID")
        arrParameter.Add("@SearchText")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)

        getCompanyProjectPricelistbySearch = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValue)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save Data"
    Public Function InsertOrderItems(ByRef arrValues As ArrayList)

        On Error GoTo Err


        Dim sp As String = "[insert_order_items]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@code")
        arrParameter.Add("@description")
        arrParameter.Add("@staffID")
        arrParameter.Add("@quantity")
        arrParameter.Add("@cost")
        arrParameter.Add("@discount")
        arrParameter.Add("@OrderID")
        arrParameter.Add("@ItemSplitID")
        arrParameter.Add("@ItemSplit")
        arrParameter.Add("@UserName")



        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Float)
        arrType.Add(SqlDbType.Money)
        arrType.Add(SqlDbType.Float)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function


    Public Sub UpdateOrderItemBatchApp(ByRef Values As ArrayList)

        On Error GoTo Err


        Dim sp As String = "[update_order_items_batch_app]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@id")
        Parameter.Add("@BatchAppID")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub

    Public Function SaveOrderItem(ByRef Values As ArrayList)

        On Error GoTo Err


        Dim sp As String = "[Save_order_items]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@id")
        Parameter.Add("@OrderID")
        Parameter.Add("@code")
        Parameter.Add("@codes")
        Parameter.Add("@description")
        Parameter.Add("@staffID")
        Parameter.Add("@quantity")
        Parameter.Add("@cost")
        Parameter.Add("@discount")
        Parameter.Add("@quoteDesc")
        Parameter.Add("@quoteLocation")
        Parameter.Add("@quoteReason")
        Parameter.Add("@ItemSplitID")
        Parameter.Add("@ItemSplit")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function InsertOrderQuickItems(ByRef Values As ArrayList)

        On Error GoTo Err


        Dim sp As String = "[insert_order_quick_items]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@OrderID")
        Parameter.Add("@code")
        Parameter.Add("@codes")
        Parameter.Add("@description")
        Parameter.Add("@quantity")
        Parameter.Add("@cost")
        Parameter.Add("@discount")
        Parameter.Add("@subTotal")
        Parameter.Add("@total")
        Parameter.Add("@quoteDesc")
        Parameter.Add("@quoteLocation")
        Parameter.Add("@quoteReason")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function UpdateOrderQuickItems(ByRef Values As ArrayList)

        On Error GoTo Err


        Dim sp As String = "[update_order_quick_items]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@OrderID")
        Parameter.Add("@code")
        Parameter.Add("@codes")
        Parameter.Add("@description")
        Parameter.Add("@quantity")
        Parameter.Add("@cost")
        Parameter.Add("@discount")
        Parameter.Add("@subTotal")
        Parameter.Add("@total")
        Parameter.Add("@quoteDesc")
        Parameter.Add("@quoteLocation")
        Parameter.Add("@quoteReason")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function SaveSubContractor(ByRef arrValues As ArrayList)

        On Error GoTo Err


        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList




        ''Update an existong record

        sp = "[update_order_items_subContractor]"

        arrParameter.Add("@id")
        arrParameter.Add("@staffID")
        arrParameter.Add("@UserName")


        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)

        For i As Integer = 0 To arrValues.Count - 1
            arrValuesPass.Add(arrValues(i))
        Next


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function SaveImportedItems(ByRef arrValues As ArrayList)

        On Error GoTo Err


        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        ''Insert a new record
        sp = "[insert_order_imported_items]"
        arrParameter.Add("@OrderID")
        arrParameter.Add("@code")
        arrParameter.Add("@description")
        arrParameter.Add("@quantity")
        arrParameter.Add("@cost")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Float)
        arrType.Add(SqlDbType.Money)
        arrType.Add(SqlDbType.VarChar)


        For i As Integer = 0 To arrValues.Count - 1
            arrValuesPass.Add(arrValues(i))
        Next




        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#Region "Delete Order Item"
    Public Function DeleteOrderItem(ByRef ItemID As Integer)

        On Error GoTo Err


        Dim sp As String = "[delete_order_item]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ItemID


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function DeleteImportedItems(ByRef OrderId As Integer)

        On Error GoTo Err


        Dim sp As String = "[delete_order_imported_items]"
        Dim strValuesPass As String = OrderId

        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, strParameter, strType, strValuesPass)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
    Public Function UpdateBatchApplicationOrderQuickItems(ByRef Values As ArrayList)

        On Error GoTo Err


        Dim sp As String = "[update_Batch_Application_order_items]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@OrderID")
        Parameter.Add("@code")
        Parameter.Add("@codes")
        Parameter.Add("@description")
        Parameter.Add("@quantity")
        Parameter.Add("@cost")
        Parameter.Add("@discount")
        Parameter.Add("@subTotal")
        Parameter.Add("@total")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
End Class
