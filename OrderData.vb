Public Class OrderData
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
#Region "Order Locked by Users"
    Public Function CheckgetOrderDataID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_locked_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        CheckgetOrderDataID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function SaveOrderUnlockUser(id As Integer)

        On Error GoTo Err

        Dim sp As String = "[UserLogin_order_unlocked_users]"
        Dim strParameter As String = "@ID"
        Dim strType As String = "SqlDbType.Int"
        Dim QueryString As String = id


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString(), sp, strParameter, strType, QueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
#Region "Get Order list Data"
    Public Function getOrderDataSearchReturnCount(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_by_search_ReturnCount]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@index")
        arrParameter.Add("@company")
        arrParameter.Add("@JobNo")
        arrParameter.Add("@Status")
        arrParameter.Add("@OtherNo")
        arrParameter.Add("@Contract")
        arrParameter.Add("@Address")
        arrParameter.Add("@InvoiceNo")
        arrParameter.Add("@PostCode")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)

        getOrderDataSearchReturnCount = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderDataSearch(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_by_search]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@index")
        arrParameter.Add("@company")
        arrParameter.Add("@JobNo")
        arrParameter.Add("@Status")
        arrParameter.Add("@OtherNo")
        arrParameter.Add("@Contract")
        arrParameter.Add("@Address")
        arrParameter.Add("@InvoiceNo")
        arrParameter.Add("@PostCode")
        arrParameter.Add("@PageIndex")
        arrParameter.Add("@PageSize")




        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)

        getOrderDataSearch = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderSaveForLater(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_saveforlater_by_loginID]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@admin")
        arrParameter.Add("@UserName")




        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)

        getOrderSaveForLater = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
#Region "Get Order Data"
    Public Function getOrderByID(ByRef UserName As String, id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_data_by_id]"
        Dim arrParameter As New ArrayList
        arrParameter.Add("@id")
        arrParameter.Add("@Username")

        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)

        Dim arrQueryString As New ArrayList
        arrQueryString.Add(id)
        arrQueryString.Add(UserName)

        getOrderByID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
#Region "Get Order Notes"
    Public Function getOrderNotesID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_notes_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getOrderNotesID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Sub InsertNotes(ByRef arrValue As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[insert_OrderNotes]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@OrderID")
        arrParameter.Add("@Notes")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)




        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValue)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

    Public Function UpdateNotes(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_OrderNotes]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@ID")
        arrParameter.Add("@OrderID")
        arrParameter.Add("@Notes")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Order Material Coding"


    Public Function getOrderMaterialByID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_ordermaterial_by_ID]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getOrderMaterialByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderMaterialByOrderIDList(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_ordermaterial_by_OrderID_list]"
        Dim strParameter As String = "@Orderid"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getOrderMaterialByOrderIDList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Sub DeleteMaterial(ByRef arrValue As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[delete_OrderMaterial]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@ID")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValue)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
#End Region
#Region "Save Data"

    Public Function InsertOrders(ByRef arrValue As ArrayList) As Integer
        On Error GoTo Err

        Dim arrQueryString As New ArrayList
        Dim sp As String = "[insert_orders]"
        Dim strParameterOutput As String = "@ID"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@CompanyID")
        arrParameter.Add("@OrderNo")
        arrParameter.Add("@Address")
        arrParameter.Add("@Postcode")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)

        InsertOrders = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString, sp, arrParameter, strParameterOutput, arrType, arrValue)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function UpdateOrders(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_Orders]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@id")
        arrParameter.Add("@companyID")
        arrParameter.Add("@job_type_id")
        arrParameter.Add("@live_update")
        arrParameter.Add("@NOC")
        arrParameter.Add("@coded")
        arrParameter.Add("@waiting_for_wholesalers")
        arrParameter.Add("@certificate_requested")
        arrParameter.Add("@certificate_done")
        arrParameter.Add("@NoAccess")
        arrParameter.Add("@Cancelled")
        arrParameter.Add("@PO_sent")
        arrParameter.Add("@PO_received")
        arrParameter.Add("@target_date")
        arrParameter.Add("@Completed_date")
        arrParameter.Add("@materials_ordered")
        arrParameter.Add("@materials_instock")
        arrParameter.Add("@vo_sent")
        arrParameter.Add("@vo_agreed")
        arrParameter.Add("@Appointment_date")
        arrParameter.Add("@cancelled_date")
        arrParameter.Add("@Start_time")
        arrParameter.Add("@End_time")
        arrParameter.Add("@order_no")
        arrParameter.Add("@ref_no")
        arrParameter.Add("@PO_number")
        arrParameter.Add("@post_code")
        arrParameter.Add("@Contract_no")
        arrParameter.Add("@Tenant")
        arrParameter.Add("@Priority")
        arrParameter.Add("@Address")
        arrParameter.Add("@companynotes")
        arrParameter.Add("@personalnotes")
        arrParameter.Add("@vo_details")
        arrParameter.Add("@po_notes")
        arrParameter.Add("@Status")
        arrParameter.Add("@OrderRun_No")
        arrParameter.Add("@MaterialNotes")
        arrParameter.Add("@UserName")


        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)


        arrType.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Sub UpdateOrderSaveForLater(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_orders_saveforlater]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@id")
        arrParameter.Add("@UserName")


        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub

#End Region
#Region "Delete Order Notes"
    Public Function DeleteNotes(ByRef ID As Integer)

        On Error GoTo Err

        Dim sp As String = "[delete_Order_Notes_by_ID]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region

#Region "Save Material"
    Public Sub InsertMaterial(ByRef arrValue As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[insert_OrderMaterial]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@OrderID")
        arrParameter.Add("@WholesalerID")
        arrParameter.Add("@StaffID")
        arrParameter.Add("@InvoiceNo")
        arrParameter.Add("@Amount")
        arrParameter.Add("@Notes")
        arrParameter.Add("@Invoicedate")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.VarChar)




        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValue)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Function UpdateMaterial(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_OrderMaterial]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@id")
        arrParameter.Add("@OrderID")
        arrParameter.Add("@WholesalerID")
        arrParameter.Add("@StaffID")
        arrParameter.Add("@InvoiceNo")
        arrParameter.Add("@Amount")
        arrParameter.Add("@Notes")
        arrParameter.Add("@Invoicedate")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region " SubContractor and Diary"
    Public Function getDiaryOrderInfoByOrderID(ByRef id As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err

        Dim sp As String = "[Order_get_diary_orderInfo_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getDiaryOrderInfoByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getOrderAssignmentByOrderID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_assignment_data_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getOrderAssignmentByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderSubContractorByOrder(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim ViperCon As New Viper.Connection.Connection
        Dim sp As String = "[Order_get_order_subcontractor_by_OrderID]"
        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id

        getOrderSubContractorByOrder = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub SaveOrderSubcontractor(ByRef arrValue As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_order_subcontractor]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@OrderID")
        arrParameter.Add("@StaffID")
        arrParameter.Add("@Notes")
        arrParameter.Add("@Date")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Text)
        arrType.Add(SqlDbType.Date)
        arrType.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValue)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
#End Region
End Class
