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


    Public Function getTimesheetbyID(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[timesheet_get_timesheet_by_ID]"
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

        Dim sp As String = "[insert_timesheet]"


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
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
End Class
