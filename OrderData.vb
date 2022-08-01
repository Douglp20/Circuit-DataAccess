﻿Public Class OrderData
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

#Region "Get Order list Data"
    Public Function getOrderCertEngineer() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[ORDER_get_certificateEngineer]"


        getOrderCertEngineer = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderLinkList(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[ORDER_get_link_order_by_ID]"
        Dim parameter As String = "@OrderID"
        Dim type As String = SqlDbType.Int

        getOrderLinkList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, parameter, type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderSearchSubContractor() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[order_search_subcontractor_list]"


        getOrderSearchSubContractor = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderDataSearchDefault(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_by_search_default]"

        Dim arrParameter As New ArrayList

        arrParameter.Add("@SearchNotStarted")
        arrParameter.Add("@SearchInProgress")





        Dim arrType As New ArrayList

        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)


        getOrderDataSearchDefault = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderDashboad() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_dashboard]"



        getOrderDashboad = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderDataSearchReturnCount(ByRef QueryString As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_by_search_ReturnCount]"

        Dim strParameter As String = "@PassWhere"
        Dim strType As String = SqlDbType.VarChar

        getOrderDataSearchReturnCount = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, QueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderDataQuickSearchbyIndex(value As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[Order_get_quick_search_by_index]"
        Dim Parameter As String = "@Index"
        Dim Type As String = SqlDbType.Int



        getOrderDataQuickSearchbyIndex = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderDataSearch(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_by_search]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@PassWhere")
        arrParameter.Add("@PageIndex")
        arrParameter.Add("@PageSize")




        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)

        getOrderDataSearch = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderDataSearchReturnCountOld(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_by_search_ReturnCount]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@index")
        arrParameter.Add("@JobNo")
        arrParameter.Add("@Status")
        arrParameter.Add("@OtherNo")
        arrParameter.Add("@Contract")
        arrParameter.Add("@Address")
        arrParameter.Add("@InvoiceNo")
        arrParameter.Add("@PostCode")
        arrParameter.Add("@SearchNotStarted")
        arrParameter.Add("@SearchInProgress")

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
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        getOrderDataSearchReturnCountOld = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderDataSearchOld(ByRef arrQueryString As ArrayList) As SqlClient.SqlDataAdapter

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
        arrParameter.Add("@SearchNotStarted")
        arrParameter.Add("@SearchInProgress")
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
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)

        getOrderDataSearchOld = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderSaveForLater(ByRef value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_saveforlater_by_loginID]"

        Dim Parameter As New ArrayList
        Parameter.Add("@admin")
        Parameter.Add("@UserName")




        Dim Type As New ArrayList
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)

        getOrderSaveForLater = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
#Region "Get Order Data"
    Public Function getOrderStatusByOrderID(ByRef OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_status_by_Orderid]"

        Dim strParameter As String = "@OrderID"

        Dim strType As String = SqlDbType.Int

        getOrderStatusByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, OrderID)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderStatusCount() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[orders_get_status_count]"


        getOrderStatusCount = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
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
    Public Function getOrderAddressDuplicate(ByRef value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_address_duplicate]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@Postcode")
        Parameter.Add("@Address")
        Parameter.Add("@OrderNo")

        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)

        getOrderAddressDuplicate = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
#Region "Get Order Notes"

    Public Function getOrderNotesByOrderID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_notes_by_OrderID]"
        Dim strParameter As String = "@Orderid"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getOrderNotesByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Sub InsertNotes(ByRef arrValue As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[insert_Order_Notes]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@OrderID")
        arrParameter.Add("@Notes")
        arrParameter.Add("@AppointmentCheck")
        arrParameter.Add("@WorksheetCheck")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)




        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValue)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

    Public Function UpdateNotes(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_Order_Notes]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@ID")
        arrParameter.Add("@OrderID")
        arrParameter.Add("@Notes")
        arrParameter.Add("@AppointmentCheck")
        arrParameter.Add("@WorksheetCheck")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderNotesHistoryByOrderID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderNotes_history_by_OrderID]"
        Dim Parameter As String = "@Orderid"
        Dim Type As String = SqlDbType.Int



        getOrderNotesHistoryByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region

#Region "Save Order Link"
    Public Sub UpdateOrderLink(ByRef Value As ArrayList)
        On Error GoTo Err

        Dim sp As String = "[UPDATE_link_order_managed_data]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@ParentID")
        Parameter.Add("@ChildID")
        Parameter.Add("@Index")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

#End Region
#Region "Save Data"

    Public Function InsertOrders(ByRef Value As ArrayList) As Integer
        On Error GoTo Err

        Dim arrQueryString As New ArrayList
        Dim sp As String = "[insert_orders]"
        Dim ParameterOutput As String = "@ID"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@CompanyID")
        Parameter.Add("@OrderNo")
        Parameter.Add("@Address")
        Parameter.Add("@Postcode")
        Parameter.Add("@CalloutNo")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)

        InsertOrders = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString, sp, Parameter, ParameterOutput, Type, Value)
        Return InsertOrders

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function UpdateOrders(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_Orders]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@id")
        Parameter.Add("@companyID")
        Parameter.Add("@job_type_id")
        Parameter.Add("@project_type_id")
        Parameter.Add("@coded")
        Parameter.Add("@certificateRequested")
        Parameter.Add("@Cancelled")
        Parameter.Add("@PO_sent")
        Parameter.Add("@PO_received")
        Parameter.Add("@target_date")
        Parameter.Add("@Completed_date")
        Parameter.Add("@vo_sent")
        Parameter.Add("@vo_agreed")
        Parameter.Add("@Appointment_date")
        Parameter.Add("@cancelled_date")
        Parameter.Add("@Start_time")
        Parameter.Add("@End_time")
        Parameter.Add("@order_no")
        Parameter.Add("@ref_no")
        Parameter.Add("@PO_number")
        Parameter.Add("@post_code")
        Parameter.Add("@Contract_no")
        Parameter.Add("@Tenant")
        Parameter.Add("@Priority")
        Parameter.Add("@Address")
        Parameter.Add("@vo_details")
        Parameter.Add("@po_notes")
        Parameter.Add("@Status")
        Parameter.Add("@OrderRun_No")
        Parameter.Add("@contractEmail")
        Parameter.Add("@voNotAgreedDate")
        Parameter.Add("@calloutNumber")
        Parameter.Add("@CherryPickerCheck")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)

        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, arrValues)


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
    Public Function getOrderAlreadyAssignedSubContractorByOrder(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err



        Dim sp As String = "[Order_get_already_assigned_subcontractor_by_OrderID]"
        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id

        getOrderAlreadyAssignedSubContractorByOrder = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub UpdateOrderSubcontractor(ByRef arrValue As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_order_subcontractor]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@OrderID")
        arrParameter.Add("@StaffID")
        arrParameter.Add("@Date")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Date)
        arrType.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValue)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
#End Region

#Region "Get Unlock Orders"
    Public Function getAllOrderLocked() As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[UserLogin_get_all_order_locked_users]"


        getAllOrderLocked = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function CheckgetOrderDataID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[UserLogin_get_order_locked_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        CheckgetOrderDataID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
#Region "Update Unlock Orders"
    Public Sub SaveOrderUnlockUser(ID As Integer)
        On Error GoTo Err



        Dim sp As String = "[Update_UserLogin_unlocked_order]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID

        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        If Err.Number > 0 Then RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

#End Region

#Region "Get Orderlink data"

    Public Function getOrderLinkManageData(ByRef value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[ORDER_get_link_order_managed_data]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@OrderID")
        Parameter.Add("@Search")
        Parameter.Add("@CompanyID")
        Parameter.Add("@index")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)



        getOrderLinkManageData = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
End Class
