Public Class OrderAttachmentData
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
#Region "Picture"
    Public Sub insertOrderImage(ByRef Value As ArrayList, ImageData As Byte())
        On Error GoTo Err


        Dim sp As String = "[insert_order_image]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Dim pictureParameter As String = "@photo"

        Parameter.Add("@ID")
        Parameter.Add("@OrderID")
        Parameter.Add("@CustomerCheck")
        Parameter.Add("@FileName")
        Parameter.Add("@extention")
        Parameter.Add("@TypeName")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessImageWithParameters(connection.ConnectionString, sp, Parameter, Type, Value, ImageData, pictureParameter)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub UpdateOrderImage(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_order_image]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@ID")
        Parameter.Add("@OrderID")
        Parameter.Add("@CustomerCheck")
        Parameter.Add("@FileName")
        Parameter.Add("@extention")
        Parameter.Add("@TypeName")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

    Public Sub SaveOrderQuotation(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[Update_order_quotation]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@OrderID")
        Parameter.Add("@Description")
        Parameter.Add("@qty")
        Parameter.Add("@location")
        Parameter.Add("@reason")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub UpdateOrderPicture(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_order_picture]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@customerPictureEmailMessage")
        Parameter.Add("@customerPictureStatus")
        Parameter.Add("@customerPictureStatusID")
        Parameter.Add("@UserName")



        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

    Public Sub deleteOrderPicture(ByRef ID As Integer)
        On Error GoTo Err


        Dim sp As String = "[delete_orderPicture]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int





        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, strParameter, strType, ID)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub deleteAllOrderPicture(ByRef value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[delete_all_orderPicture]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@OrderID")
        Parameter.Add("@type")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Function getOrderPicture(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderPicture_by_ID]"

        Dim Parameter As String = "@OrderID"

        Dim Type As String = SqlDbType.Int

        getOrderPicture = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderItemPicture(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderItemPicture_by_ID]"

        Dim Parameter As String = "@OrderID"

        Dim Type As String = SqlDbType.Int

        getOrderItemPicture = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderWorksheetPicture(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderWorkSheet_Picture_by_ID]"

        Dim Parameter As String = "@OrderID"

        Dim Type As String = SqlDbType.Int

        getOrderWorksheetPicture = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getInvoiceAttachment(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Invoice_get_invoice_attachment_by_ID]"

        Dim Parameter As String = "@InvoiceID"

        Dim Type As String = SqlDbType.Int

        getInvoiceAttachment = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderQuotationPicture(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderQuotationPicture_by_ID]"

        Dim Parameter As String = "@OrderID"

        Dim Type As String = SqlDbType.Int

        getOrderQuotationPicture = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderQuotationTestByID(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_order_quotation_test_by_ID]"

        Dim Parameter As String = "@OrderID"

        Dim Type As String = SqlDbType.Int

        getOrderQuotationTestByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderAttachmentMessage(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderAttachment_status_by_id]"

        Dim Parameter As String = "@ID"

        Dim Type As String = SqlDbType.Int

        getOrderAttachmentMessage = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderCertificate(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderCertificate_by_ID]"

        Dim Parameter As String = "@OrderID"

        Dim Type As String = SqlDbType.Int

        getOrderCertificate = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getCustomerAttachment(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Company_get_Attachment_by_companyID]"

        Dim Parameter As String = "@CompanyID"

        Dim Type As String = SqlDbType.Int

        getCustomerAttachment = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getMaterialAttachment(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_Material_Attachment_by_companyID]"

        Dim Parameter As String = "@ID"

        Dim Type As String = SqlDbType.Int

        getMaterialAttachment = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
End Class
