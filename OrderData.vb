Public Class OrderData
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

#Region " Subbie Request ReadOnly Form"

    Public Sub SubbieAppointmentJobRequest(ByRef value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_orders_subbie_request_appointment]"



        Dim Type As New ArrayList
        Dim Parameter As New ArrayList


        Parameter.Add("@ID")
        Parameter.Add("@userName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
#End Region


#Region "Order PO Request"

    Public Function getPhase5VORequestItemCountByOrderID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_phase5_vo_request_item_count_by_OrderID]"

        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int


        getPhase5VORequestItemCountByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function
    Public Function getPhase5VORequestIDByOrderID(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_phase5_vo_request_by_OrderID]"

        Dim Parameter As String = "@OrderID"

        Dim Type As String = SqlDbType.Int

        getPhase5VORequestIDByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Function getOrderVORequestInfoOrderID(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_Advance_Info]"

        Dim Parameter As String = "@OrderID"

        Dim Type As String = SqlDbType.Int

        getOrderVORequestInfoOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Function getOrderPhase3VORequestByOrderID(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_phase3_vo_request_item_by_OrderID]"

        Dim Parameter As String = "@OrderID"

        Dim Type As String = SqlDbType.Int

        getOrderPhase3VORequestByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Function getOrderVORequestByOrderID(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_vo_request_by_OrderID]"

        Dim Parameter As String = "@OrderID"

        Dim Type As String = SqlDbType.Int

        getOrderVORequestByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Sub updateOrderVORequestStatusBYVORequestID(ByRef value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_order_vo_request_status_by_VORequestID]"



        Dim Type As New ArrayList
        Dim Parameter As New ArrayList

        Parameter.Add("@VORequestID")
        Parameter.Add("@StateDate")
        Parameter.Add("@State")
        Parameter.Add("@userName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
#End Region

#Region "Phase 2 WorkTicket"
    Public Function getWorkTicketSubcontractorByOrderID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Order_get_phase2_workticket_subcontractor_by_OrderID]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int



        getWorkTicketSubcontractorByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getWorkTicketByID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Order_get_workticket_By_ID]"
        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int



        getWorkTicketByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getWorkTicketNotesByWorkID(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Order_get_phase2_workticket_subcontractor_notes_by_workID]"



        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@OrderID")
        Parameter.Add("@workID")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)


        getWorkTicketNotesByWorkID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Sub UpdateWorkTicketSubcontractor(value As ArrayList)

        On Error GoTo Err


        Dim sp As String = "[update_Order_phase2_workticket]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@tick1")
        Parameter.Add("@tick2")
        Parameter.Add("@tick3")
        Parameter.Add("@tick4")
        Parameter.Add("@tick5")
        Parameter.Add("@tick6")
        Parameter.Add("@tick7")
        Parameter.Add("@tick8")
        Parameter.Add("@tick9")
        Parameter.Add("@tick10")
        Parameter.Add("@tick11")
        Parameter.Add("@tenant1")
        Parameter.Add("@tenant2")
        Parameter.Add("@contractorCardDate")
        Parameter.Add("@contractorCardTime")
        Parameter.Add("@NoAccessCardDate")
        Parameter.Add("@NoAccessCardTime")
        Parameter.Add("@UserName")



        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
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
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region


#Region "Get Company Data"
    Public Function getAllCompanyInfoForNewOrder() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[ORDER_get_company_info_for_newOrder]"



        getAllCompanyInfoForNewOrder = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAllCompanyInfoForNewOrderSearch(value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[ORDER_get_company_info_for_newOrder_search]"
        Dim Parameter As String = "@company"
        Dim Type As String = SqlDbType.VarChar



        getAllCompanyInfoForNewOrderSearch = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanySubNewOrderByID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[ORDER_get_CompanyBranch_new_order_by_id]"
        Dim Parameter As String = "@id"
        Dim Type As String = SqlDbType.Int



        getCompanySubNewOrderByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function


#End Region

#Region "SubContractor Assignments"

    Public Function getOrderSubContractorAssignmentSearch(value As ArrayList) As SqlClient.SqlDataAdapter
        On Error GoTo Err

        Dim sp As String = "[Order_get_order_SubContractor_assignment_search]"
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

        getOrderSubContractorAssignmentSearch = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function
    Public Function getOrderSubContractorAssignment(value As ArrayList) As SqlClient.SqlDataAdapter
        On Error GoTo Err

        Dim sp As String = "[Order_get_order_SubContractor_assignment]"
        Dim Parameter As New ArrayList

        Parameter.Add("@Year")
        Parameter.Add("@Month")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        getOrderSubContractorAssignment = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function
#End Region
#Region "Get Order list Data"

    Public Function getOrderPhaseText(value As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err

        Dim sp As String = "[Order_get_phase_index_text]"
        Dim Parameter As String = "@Index"
        Dim Type As String = SqlDbType.Int



        getOrderPhaseText = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)



    End Function
    Public Function getOrderPhaseByID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim SP As String = "[Order_get_order_phase_by_id]"
        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int



        getOrderPhaseByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, SP, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getOrderPhaseHistoryByID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim SP As String = "[Order_get_order_phase_history_by_id]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int



        getOrderPhaseHistoryByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, SP, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

    Public Function GetCustomerEmailInfo(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim SP As String = "[EMAIL_get_customer_email_info_by_id]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int



        GetCustomerEmailInfo = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, SP, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function




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
    Public Function getOrderAdvanceInfo(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_Advance_Info]"

        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int



        getOrderAdvanceInfo = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)

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
    Public Function getOrderDataSearchReturnCount(ByRef value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_by_search_ReturnCount]"

        Dim parameter As String = "@PassWhere"
        Dim type As String = SqlDbType.VarChar

        getOrderDataSearchReturnCount = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, parameter, type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderDataQuickSearchbyIndex(value As ArrayList) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[Order_get_dashboard_search_by_index]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Parameter.Add("@Index")
        Parameter.Add("@LoginID")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)





        getOrderDataQuickSearchbyIndex = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderDataSearch(ByRef value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_by_search]"

        Dim Parameter As New ArrayList
        Parameter.Add("@PassWhere")
        Parameter.Add("@PageIndex")
        Parameter.Add("@PageSize")



        Dim Type As New ArrayList
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)


        getOrderDataSearch = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


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



    Public Function GetOrderIDbyInvoiceID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim SP As String = "[Order_get_InvoiceID_by_id]"
        Dim Parameter As String = "@InvoiceID"
        Dim Type As String = SqlDbType.Int



        GetOrderIDbyInvoiceID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, SP, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getApppointmentTimeList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[ORDER_get_appointmentTimes_list]"



        getApppointmentTimeList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getAllApppointmentTimeList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[ORDER_get_all_appointmentTimes_list]"



        getAllApppointmentTimeList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

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
    Public Function getOrderDuplication(ByRef value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_duplicatation]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@OrderNo")
        Parameter.Add("@RefNo")
        Parameter.Add("@ClientNo")
        Parameter.Add("@CallOutNo")
        Parameter.Add("@Address")
        Parameter.Add("@Postcode")


        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)

        getOrderDuplication = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
#Region "Get Order Notes"

    Public Function getQuotationOrdeNotesByOrderID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderQuotation_notes_by_OrderID]"
        Dim strParameter As String = "@Orderid"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getQuotationOrdeNotesByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderWorksheetNotesByOrderID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_worksheet_notes_by_OrderID]"
        Dim strParameter As String = "@Orderid"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getOrderWorksheetNotesByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderPhase5WorksheetNotesByOrderID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_phase5_order_worksheet_notes_by_OrderID]"
        Dim strParameter As String = "@Orderid"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getOrderPhase5WorksheetNotesByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Sub InsertNotes(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[insert_Order_Notes]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@OrderID")
        Parameter.Add("@Notes")
        Parameter.Add("@SubContractorID")
        Parameter.Add("@AppointmentCheck")
        Parameter.Add("@WorksheetCheck")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)




        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

    Public Function UpdateNotes(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_Order_Notes]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@ID")
        Parameter.Add("@OrderID")
        Parameter.Add("@Notes")
        Parameter.Add("@SubContractorID")
        Parameter.Add("@AppointmentCheck")
        Parameter.Add("@WorksheetCheck")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, Values)


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
        Parameter.Add("@CompanySubID")
        Parameter.Add("@OrderNo")
        Parameter.Add("@RefNo")
        Parameter.Add("@ClientNo")
        Parameter.Add("@CallOutNo")
        Parameter.Add("@Address")
        Parameter.Add("@Postcode")
        Parameter.Add("@targetdate")
        Parameter.Add("@jobtypeid")
        Parameter.Add("@projectTypeID")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)

        InsertOrders = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString, sp, Parameter, ParameterOutput, Type, Value)
        Return InsertOrders

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function UpdateOrders(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_Orders]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@id")
        Parameter.Add("@companyID")
        Parameter.Add("@CompanySubID")
        Parameter.Add("@coded")
        Parameter.Add("@CancelIndex")
        Parameter.Add("@Address")
        Parameter.Add("@postcode")
        Parameter.Add("@orderNumber")
        Parameter.Add("@refNumber")
        Parameter.Add("@calloutNumber")
        Parameter.Add("@ClientJobNumber")
        Parameter.Add("@jobTypeID")
        Parameter.Add("@projectTypeID")
        Parameter.Add("@Tenant")
        Parameter.Add("@TenantEmail")
        Parameter.Add("@Priority")
        Parameter.Add("@ContractNumber")
        Parameter.Add("@CherryPickerCheck")
        Parameter.Add("@MaterialOrdered")
        Parameter.Add("@ActionStatus")
        Parameter.Add("@ActionComment")
        Parameter.Add("@PhaseID")
        Parameter.Add("@OrderNotes")
        Parameter.Add("@userID")
        Parameter.Add("@saveforlaterCheck")
        Parameter.Add("@payrollOrderCancel")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateOrderPhase1(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_orders_Phase1]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@id")
        Parameter.Add("@companyID")
        Parameter.Add("@CompanySubID")
        Parameter.Add("@job_type_id")
        Parameter.Add("@project_type_id")
        Parameter.Add("@AppointmentTimeID")
        Parameter.Add("@target_date")
        Parameter.Add("@Start_time")
        Parameter.Add("@End_time")
        Parameter.Add("@order_no")
        Parameter.Add("@ref_no")
        Parameter.Add("@post_code")
        Parameter.Add("@Tenant")
        Parameter.Add("@Priority")
        Parameter.Add("@Address")
        Parameter.Add("@contractEmail")
        Parameter.Add("@calloutNumber")
        Parameter.Add("@ClientJobNumber")
        Parameter.Add("@OrderNotes")
        Parameter.Add("@CancelIndex")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
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
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Sub UpdateOrderPhase2(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_orders_Phase2]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@Tenant")
        Parameter.Add("@Priority")
        Parameter.Add("@contractEmail")
        Parameter.Add("@OrderNotes")
        Parameter.Add("@Action")
        Parameter.Add("@ActionText")
        Parameter.Add("@certRequestedCheck")
        Parameter.Add("@UserName")





        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub UpdateOrderPhase2Confirmation(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_orders_Phase2_Confirmation]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@Action")
        Parameter.Add("@certRequestedCheck")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub UpdateOrderPhase3(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_orders_Phase3Certificate]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@CertificateEmailMessage")
        Parameter.Add("@CertificateStatusID")
        Parameter.Add("@CertificateStatus")
        Parameter.Add("@OrderNotes")
        Parameter.Add("@UserName")



        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub UpdateOrderPhase4(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_orders_Phase4]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@Address")
        Parameter.Add("@postcode")
        Parameter.Add("@orderNumber")
        Parameter.Add("@refNumber")
        Parameter.Add("@ClientJobNumber")
        Parameter.Add("@calloutNumber")
        Parameter.Add("@Tenant")
        Parameter.Add("@TenantEmail")
        Parameter.Add("@MaterialOrdered")
        Parameter.Add("@Priority")
        Parameter.Add("@OrderNotes")
        Parameter.Add("@Confirmation")
        Parameter.Add("@ContractNumber")
        Parameter.Add("@PhaseID")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
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
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)





        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)

        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub UpdateOrderPhase5(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_orders_Phase5]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Function UpdateOrderPhase6(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_orders_Phase6]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@id")
        Parameter.Add("@PO_sent")
        Parameter.Add("@PO_received")
        Parameter.Add("@vo_sent")
        Parameter.Add("@vo_agreed")
        Parameter.Add("@cancelled_date")
        Parameter.Add("@order_no")
        Parameter.Add("@ref_no")
        Parameter.Add("@PO_number")
        Parameter.Add("@Contract_no")
        Parameter.Add("@Tenant")
        Parameter.Add("@vo_details")
        Parameter.Add("@po_notes")
        Parameter.Add("@OrderRun_No")
        Parameter.Add("@contractEmail")
        Parameter.Add("@voNotAgreedDate")
        Parameter.Add("@CherryPickerCheck")
        Parameter.Add("@ClientJobNumber")
        Parameter.Add("@OrderNotes")
        Parameter.Add("@ActionStatus")
        Parameter.Add("@Comment")
        Parameter.Add("@MaterialOrdered")
        Parameter.Add("@Phase")
        Parameter.Add("@UserLoginID")
        Parameter.Add("@SaveForLaterCheck")
        Parameter.Add("@UserName")




        Type.Add(SqlDbType.Int)
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
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
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
    Public Function getOrderBatchApplication(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err



        Dim sp As String = "[Order_get_batch_application_data_by_id]"
        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id

        getOrderBatchApplication = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)


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
