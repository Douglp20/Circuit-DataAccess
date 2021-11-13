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
    Public Function SaveSettings(ByRef arrValues As ArrayList)

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

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region


End Class
