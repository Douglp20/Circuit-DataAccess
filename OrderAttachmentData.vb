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

    Public Sub zzzUpdateAttachmentMessage(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_orderAttachment_status]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@customerPictureEmailMessage")
        Parameter.Add("@customerCertificateEmailMessage")
        Parameter.Add("@customerPictureStatus")
        Parameter.Add("@customerCertificateStatus")
        Parameter.Add("@UserName")



        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
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
    Public Sub UpdateOrderPicture(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_order_picture]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@customerPictureEmailMessage")
        Parameter.Add("@customerPictureStatus")
        Parameter.Add("@UserName")



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
    Public Sub UpdateOrderCertificate(ByRef Value As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[update_order_cerificate]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@customerCertificateEmailMessage")
        Parameter.Add("@customerCertificateStatus")
        Parameter.Add("@UserName")



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
    Public Function getOrderCustomer(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_orderCusAttachment_by_ID]"

        Dim Parameter As String = "@OrderID"

        Dim Type As String = SqlDbType.Int

        getOrderCustomer = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
End Class
