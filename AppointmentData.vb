Public Class AppointmentData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub


#Region "Error Control"
    Private Sub ErrorMessage_ViperCon(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#End Region


#Region "Get Data"
    Public Function getAppointmentSubContractorByDate(selDate As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err



        Dim sp As String = "[Appointment_get_subcontractor_by_date]"
        Dim strParameter As String = "@date"
        Dim strType As String = SqlDbType.VarChar
        Dim strQueryString As String = selDate


        getAppointmentSubContractorByDate = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getAppointmentByDateReturnCount(selDate As String, staffID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Appointment_get_appointment_by_date_ReturnCount]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@date")
        arrParameter.Add("@staffID")


        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Int)

        Dim arrQueryString As New ArrayList
        arrQueryString.Add(selDate)
        arrQueryString.Add(staffID)

        getAppointmentByDateReturnCount = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
    Public Function getAppointmentByDate(arrValue As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err



        Dim sp As String = "[Appointment_get_appointment_by_date]"

        Dim arrParameter As New ArrayList
        arrParameter.Add("@date")
        arrParameter.Add("@staffID")
        arrParameter.Add("@PageIndex")
        arrParameter.Add("@PageSize")

        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)


        getAppointmentByDate = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, arrParameter, arrType, arrValue)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getJobAppointmentByDate(Value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err



        Dim sp As String = "[Appointment_get_job_appointment_by_date]"

        Dim Parameter As New ArrayList
        Parameter.Add("@date")
        Parameter.Add("@staffID")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.Int)


        getJobAppointmentByDate = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString(), sp, Parameter, Type, Value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getAppointmentListByDate(dte As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err



        Dim sp As String = "[Appointment_get_appointment_list_by_date]"

        Dim strParameter As String = "@date"

        Dim strType As String = SqlDbType.DateTime



        getAppointmentListByDate = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, dte)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getAvailableSubContractorByDate(dte As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Appointment_get_subContractor_list]"
        Dim strParameter As String = "@Date"
        Dim strType As String = SqlDbType.Char

        getAvailableSubContractorByDate = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, dte)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAvailableSubContractorNotOnLeave(dte As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Appointment_get_subContractor_not_on_leave]"
        Dim strParameter As String = "@Date"
        Dim strType As String = SqlDbType.Char

        getAvailableSubContractorNotOnLeave = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, strParameter, strType, dte)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#Region "Save Data"

#End Region

End Class
