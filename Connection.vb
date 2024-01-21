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

    Private Const constConnectionStringProd = "ConnectionString-prod"
    Private Const constPWDStringProd = "pwd-prod"
    Private Const constUserNameStringProd = "UserName-prod"
    Private Const constServerStringProd = "Server-prod"
    Private Const constReportDatabaseStringProd = "ReportDatabase-prod"

    Private Const constConnectionStringTest = "ConnectionString-test"

    Private Const constConnectionStringDev = "ConnectionString-dev"
    Private Const constPWDStringDev = "pwd-dev"
    Private Const constUserNameStringDev = "UserName-dev"
    Private Const constServerStringDev = "Server-dev"
    Private Const constReportDatabaseStringDev = "ReportDatabase-dev"

    Public Function ConnectionString() As String
        Return getConnection()

    End Function
    Public Function getConnection() As String

        Dim curEnvironment As String = VenomRegistry.GetSetting("Settings", modRegistry.Environment, "")
        Dim connnectionEnvironment As String = String.Empty

        If curEnvironment = "" Then
            connnectionEnvironment = System.Configuration.ConfigurationManager.ConnectionStrings(constConnectionStringProd).ToString
        Else
            Select Case curEnvironment
                Case modRegistry.Production
                    connnectionEnvironment = System.Configuration.ConfigurationManager.ConnectionStrings(constConnectionStringProd).ToString
                Case modRegistry.Development
                    connnectionEnvironment = System.Configuration.ConfigurationManager.ConnectionStrings(constConnectionStringDev).ToString
                Case modRegistry.Testing
                    connnectionEnvironment = System.Configuration.ConfigurationManager.ConnectionStrings(constConnectionStringTest).ToString
            End Select

        End If

        Return VenomSecurity.Decode(connnectionEnvironment.ToString())
    End Function
    Public Function ConnectionChecker() As Boolean

        Try


            Dim strConnectionString As String = getConnection()


            Dim cmd As New SqlClient.SqlCommand()
            Dim con As New SqlClient.SqlConnection(strConnectionString)
            con.Open()

            cmd.Connection = con

            Dim apt As New SqlClient.SqlDataAdapter(cmd)

            con.Close()
            ConnectionChecker = True

        Catch ex As Exception
            ConnectionChecker = False
        End Try
    End Function
End Class
