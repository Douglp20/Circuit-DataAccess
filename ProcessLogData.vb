Public Class ProcessLogData
    Public Sub New()
    End Sub

    Private WithEvents ViperCon As New Viper.Connection.Connection()
    Private Connection As New Connection

#Region "Error Control"
    Public Event errorMessage(ByVal errDes As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private Sub errorMessage_Event(ByVal errDes As String, ByVal errNo As Integer, ByVal errTrace As String)
        Dim errMessage As String = ">> Called by the module : " + Me.ToString()
        RaiseEvent errorMessage(errDes, errNo, errTrace)
    End Sub

#End Region

#Region "The ProcessLog"
    Public Function ProcessLog(message As String)

        On Error GoTo Err


        Dim sp As String = "[msg_proccesLog]"
        Dim strParameter As String = "@log"
        Dim strType As String = SqlDbType.Char
        Dim strQueryString As String = message

        Dim d As String = Connection.getConnection

        ViperCon.ExecuteProcessWithParameter(Connection.getConnection, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + ToString() + "."
        RaiseEvent errorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region

End Class
