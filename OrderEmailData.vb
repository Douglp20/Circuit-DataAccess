Public Class OrderEmailData
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
    Public Function zzUpdateOrderEmail(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_Order_Email]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@OrderID")
        arrParameter.Add("@certificateCheck")
        arrParameter.Add("@attachmentCheck")
        arrParameter.Add("@certificateMessage")
        arrParameter.Add("@attachmentMessage")
        arrParameter.Add("@UserName")


        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Text)
        arrType.Add(SqlDbType.Text)
        arrType.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getEmailSchedule() As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[email_get_all_scheduled]"


        getEmailSchedule = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getEmailSent() As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[email_get_all_sent]"


        getEmailSent = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
End Class
