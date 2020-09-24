Public Class Connection
    Private WithEvents VenomRegistry As New Douglas.Venom.Registry.VenomRegistry()
    Private WithEvents VenomSecurity As New Douglas.Venom.Security.Cipher()


#Region "Error Control"
    Public Event errorMessage(ByVal errDes As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private Sub errorMessage_Event(ByVal errDes As String, ByVal errNo As Integer, ByVal errTrace As String)
        Dim errMessage As String = ">> Called by the module : " + Me.ToString()
        RaiseEvent errorMessage(errDes, errNo, errTrace)
    End Sub

#End Region

    Private Const constConnectionString = "ConnectionString"
    Private Const constPWDString = "pwd"
    Private Const constUserNameString = "UserName"
    Private Const constServerString = "Server"
    Private Const constReportDatabaseString = "ReportDatabase"

    Public Function ConnectionString() As String
        Return getConnection()
    End Function
    Public Function getConnection() As String
        Dim connnectionEnvironment As String = System.Configuration.ConfigurationManager.ConnectionStrings(constConnectionString).ToString()
        Return VenomSecurity.Decode(System.Configuration.ConfigurationManager.ConnectionStrings(constConnectionString).ToString())
    End Function
End Class
