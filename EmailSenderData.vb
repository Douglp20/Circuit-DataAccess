Public Class EmailSenderData
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
#Region "Material"
    Public Function GetMailMaterial(value As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err


        Dim sp As String = "[EMAIL_order_get_wholesaler_by_id]"


        Dim Parameter As String = "@MatID"
        Dim Type As String = SqlDbType.Int

        GetMailMaterial = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub UpdateMailWholesaler(value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_email_order_Wholesaler]" ' "[update_service_mail_Wholesaler]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Parameter.Add("@ID")
        Parameter.Add("@subject")
        Parameter.Add("@emailMessage")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, value)




Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
#End Region

#Region "Order Picture Certificate"

    Public Function getOrderEmailCustomerInfoID(id As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[EMAIL_order_email_customer_info_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getOrderEmailCustomerInfoID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Order Picture"

    Public Function getOrderNotes(id As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[EMAIL_order_picture_notes_by_id]"
        Dim strParameter As String = "@OrderID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getOrderNotes = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getPicturebyOrderID(id As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[EMAIL_order_picture_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getPicturebyOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function


    Public Sub updateEmailOrderPictureSent(ByRef Value As Integer)

        On Error GoTo Err


        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int

        Dim sp As String = "[update_email_order_picture_sent]"

        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
#Region "Order Certificate"
    Public Function getCertificatebyOrderID(id As Integer) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[EMAIL_order_certificate_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getCertificatebyOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Sub updateEmailOrderCertificateSent(ByRef Value As Integer)

        On Error GoTo Err


        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int

        Dim sp As String = "[update_email_order_certificate_sent]"

        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub


#End Region
#Region "Order Worksheet"

    Public Function GetMailOrderWorkSheet(Value As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err


        Dim sp As String = "[EMAIL_order_get_worksheet_subContractor_by_Orderid]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int

        GetMailOrderWorkSheet = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function




    Public Sub UpdateMailOrderWorkSheet(value As Integer)
        On Error GoTo Err


        Dim sp As String = "[update_Service_email_order_worksheet_subContractor]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Int

        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)




Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        If Err.Number > 0 Then RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
#End Region
#Region "Customer Statement"
    Public Function GetCustomerInvoiceStatement(Value As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err


        Dim sp As String = "[EMAIL_company_invoice_statement_by_companyID]"
        Dim Parameter As String = "@companyID"
        Dim Type As String = SqlDbType.Int

        GetCustomerInvoiceStatement = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function CustomerAppointmentStatement(Value As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err


        Dim sp As String = "[EMAIL_company_appointment_Statement_by_companyID]"
        Dim Parameter As String = "@companyID"
        Dim Type As String = SqlDbType.Int

        CustomerAppointmentStatement = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#Region "Timesheet"
    Public Function TimesheetRejectionEmail(value As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err


        Dim sp As String = "[EMAIL_get_timesheet_by_id]"


        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int

        TimesheetRejectionEmail = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#Region "Settings"

    Public Function getEmailSetting() As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[SETTING_get_settings]"



        getEmailSetting = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
End Class
