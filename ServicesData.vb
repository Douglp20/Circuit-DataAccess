Public Class ServicesData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)

    Public Sub New()
    End Sub

    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
    Private connection As New Connection

#Region "Phase Service"
    Public Sub ExecPhaseService()
        On Error GoTo Err



        Dim sp As String = "[Update_service_order_dashboard]"



        ViperCon.ExecuteProcess(connection.ConnectionString(), sp)



        Exit Sub


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
#End Region

#Region "Get Data"
    Public Function getServiceSchedule() As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[SERVICE_get_all_scheduled]"


        getServiceSchedule = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getServiceHistory(value As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[Service_get_history_by_scheduleID]"
        Dim Parameter As String = "@scheduleID"
        Dim Type As String = SqlDbType.Int



        getServiceHistory = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getEmailSetting() As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[SETTING_get_settings]"


        getEmailSetting = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)




        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function GetSubContractorWorksheet() As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[Service_get_mail_subcontractor_worksheet]"


        GetSubContractorWorksheet = ViperCon.getSqlDataAdapter(connection.ConnectionString(), sp)


        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getServiceDataByServiceName(value As String) As SqlClient.SqlDataAdapter
        On Error GoTo Err




        Dim sp As String = "[Service_get_data_by_ServiceName]"
        Dim Parameter As String = "@ServiceName"
        Dim Type As String = SqlDbType.VarChar


        getServiceDataByServiceName = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)


        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#Region "Update Data"
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
    Public Sub updateServiceRunnerByScheduleID(Value As ArrayList)
        On Error GoTo Err




        Dim sp As String = "[update_Service_RunnerSwitch]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Parameter.Add("@ID")
        Parameter.Add("@RunnerSwitch")



        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, Value)



        Exit Sub


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub updateServiceResetIntervalTimeByScheduleID(Value As Integer)
        On Error GoTo Err




        Dim sp As String = "[update_Service_reset_interval_time_by_ID]"

        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString(), sp, Parameter, Type, Value)



        Exit Sub


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
#End Region
End Class
