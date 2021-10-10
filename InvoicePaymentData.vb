Public Class InvoicePaymentData
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
    Public Function getInvoicePaymentCompanyDropdownlist() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Payment_get_invoice_Payment_Company_dropdown_list]"


        getInvoicePaymentCompanyDropdownlist = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getInvoicePaymentInvoices(companyID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Payment_get_invoice_by_companyID]"
        Dim strParameter As String = "@companyID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = companyID

        getInvoicePaymentInvoices = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getInvoicePaymentFindInvoices(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        sp = "[Payment_get_invoice_find_invoiceNo]"

        arrParameter.Add("@CompanyID")
        arrParameter.Add("@invoiceNo")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)



        getInvoicePaymentFindInvoices = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getInvoicePaymentbyInvoiceID(InvoiceID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Payment_get_invoice_payment_by_invoiceID]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = InvoiceID

        getInvoicePaymentbyInvoiceID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getInvoicePaymentOrderItems(InvoiceID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Payment_get_orderitem_by_InvoiceID]"
        Dim strParameter As String = "@InvoiceID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = InvoiceID

        getInvoicePaymentOrderItems = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getInvoicePaymentDetail(InvoiceID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Payment_get_invoiceDetail_By_InvoiceID]"
        Dim strParameter As String = "@InvoiceID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = InvoiceID

        getInvoicePaymentDetail = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region

#Region "Save Data"

    Public Function InsertInvoicePayment(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        sp = "[insert_Invoice_Payment]"

        arrParameter.Add("@InvoiceID")
        arrParameter.Add("@UserName ")

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
