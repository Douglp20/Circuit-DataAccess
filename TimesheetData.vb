Public Class TimesheetData
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

#Region "User Timesheet"


    Public Function getTimesheetbyID(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[STAFF_timesheet_get_timesheet_by_ID]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@StaffID")
        Parameter.Add("@ymwID")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)

        getTimesheetbyID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Sub SaveTimesheet(value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[insert_staff_timesheet]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Parameter.Add("@StaffID")
        Parameter.Add("@ymwID")
        Parameter.Add("@mon")
        Parameter.Add("@tue")
        Parameter.Add("@wed")
        Parameter.Add("@thu")
        Parameter.Add("@fri")
        Parameter.Add("@sat")
        Parameter.Add("@sun")
        Parameter.Add("@total")
        Parameter.Add("@startMon")
        Parameter.Add("@starttue")
        Parameter.Add("@startwed")
        Parameter.Add("@startthu")
        Parameter.Add("@startfri")
        Parameter.Add("@startsat")
        Parameter.Add("@startsun")
        Parameter.Add("@endmon")
        Parameter.Add("@endtue")
        Parameter.Add("@endwed")
        Parameter.Add("@endthu")
        Parameter.Add("@endfri")
        Parameter.Add("@endsat")
        Parameter.Add("@endsun")
        Parameter.Add("@release")
        Parameter.Add("@userName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.Float)
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
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
#Region "Approval Timesheet"
    Public Function getReleasedTimesheet(value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[STAFF_timesheet_get_released_timesheet]"
        Dim Parameter As String = "@year"
        Dim type As String = SqlDbType.VarChar

        getReleasedTimesheet = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Approval-Reject Timesheet"
    Public Sub ApprovalRejectTimesheet(value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_staff_timesheet]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Parameter.Add("@ID")
        Parameter.Add("@reject")
        Parameter.Add("@approve")
        Parameter.Add("@userName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
End Class
