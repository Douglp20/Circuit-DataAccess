Public Class MailServiceData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)

    Public Sub New()
    End Sub

    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
    Private connection As New Connection
#Region "Error Control"
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#End Region

#Region "Get Data"
    Public Function getEmailSetting() As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[SETTING_get_settings]"


        getEmailSetting = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)




        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getEmailRunnerHistory(Value As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[EmailService_get_history_by_runnerID]"

        Dim Parameter As String = "RunnerID"
        Dim Type As String = SqlDbType.Int

        getEmailRunnerHistory = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, Value)




        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

    Public Function getEmailSchedule() As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[EmailService_get_all_scheduled]"


        getEmailSchedule = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

    Public Function getEmailRunnerTreeByScheduleID(Value As ArrayList) As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[EmailService_get_runner_tree_by_scheduleID]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Parameter.Add("@scheduleID")
        Parameter.Add("@index")
        Parameter.Add("@year")
        Parameter.Add("@month")
        Parameter.Add("@Day")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        getEmailRunnerTreeByScheduleID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, Parameter, Type, Value)




        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#Region "Save Data"
    Public Sub UpdateEmailSetting(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_setting_emailservice]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@id")
        Parameter.Add("@emailServiceRunning")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, arrValues)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
#Region "Update"
    Public Sub updateEmailRunnerByScheduleID(Value As ArrayList)
        On Error GoTo Err




        Dim sp As String = "[update_EmailService]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Parameter.Add("@ID")
        Parameter.Add("@value")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, Value)




        Exit Sub


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

#End Region
End Class
