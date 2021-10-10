Public Class msgProcessLog
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New douglas.Viper.Connection.Connection()
    Private connection As New Connection

End Class
