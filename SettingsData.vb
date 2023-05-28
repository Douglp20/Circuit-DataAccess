Public Class SettingsData

    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New douglas.Viper.Connection.Connection()
    Private connection As New Connection
#Region "Error Control"
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#End Region


#Region "Schedule"

    Public Function getSettingCompanySchedule() As SqlClient.SqlDataAdapter
        On Error GoTo Err

        ''Company_get_jobType_by_contactID
        Dim sp As String = "[Setting_get_Company_Schedule]"


        getSettingCompanySchedule = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Sub saveSettingCompanySchedule(ByRef Value As ArrayList)

        On Error GoTo Err


        Dim SP As String = "[SAVE_Setting_Company_Schedule]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@ID")
        Parameter.Add("@company")
        Parameter.Add("@Schedule1")
        Parameter.Add("@Schedule2")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, SP, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region

#Region "Get Data"
    Public Function getSettings() As SqlClient.SqlDataAdapter

        On Error GoTo Err

       
        Dim sp As String = "[SETTING_get_settings]"


        getSettings = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save Data"
    Public Sub SaveSettings(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_settings]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@id")
        Parameter.Add("@Address")
        Parameter.Add("@info")
        Parameter.Add("@signature")
        Parameter.Add("@CC")
        Parameter.Add("@BB")
        Parameter.Add("@reply")
        Parameter.Add("@smtpauthenticate")
        Parameter.Add("@sendusername")
        Parameter.Add("@sendpassword")
        Parameter.Add("@smtpserver")
        Parameter.Add("@sendusing")
        Parameter.Add("@smtpserverport")
        Parameter.Add("@smtpusessl")
        Parameter.Add("@vatReverse")
        Parameter.Add("@CustomerEmail")
        Parameter.Add("@WholesalerEmail")
        Parameter.Add("@emailServiceRunning")
        Parameter.Add("@emailHeader")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, arrValues)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region


End Class
