Public Class OrderNoAccessData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub


#Region "Get Data"
    Public Function getNoAccessByOrderID(ByRef OrderID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Order_get_noAccess_by_Orderid]"
        Dim Parameter As String = "@OrderID"
        Dim Type As String = SqlDbType.Bit


        getNoAccessByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, OrderID)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getNoAccessByID(ByRef ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Order_get_noAccess_by_id]"
        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Bit


        getNoAccessByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, ID)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
#Region "Save Data"
    Public Sub DeleteNoAccess(ByRef ID As Integer)

        On Error GoTo Err

        Dim sp As String = "[delete_noAccess]"

        Dim Parameter As String = "@id"

        Dim Type As String = "SqlDbType.Int"



        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString(), sp, Parameter, Type, ID)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
    Public Sub InsertNoAccess(ByRef arrQueryString As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[insert_noAccess]"

        Dim Parameter As New ArrayList
        Parameter.Add("@Orderid")
        Parameter.Add("@accessdate")
        Parameter.Add("@accesstime")
        Parameter.Add("@comment")
        Parameter.Add("@UserLogin")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, arrQueryString)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
    Public Sub UpdateNoAccess(ByRef arrQueryString As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_noAccess]"

        Dim Parameter As New ArrayList
        Parameter.Add("@id")
        Parameter.Add("@Orderid")
        Parameter.Add("@accessdate")
        Parameter.Add("@accesstime")
        Parameter.Add("@comment")
        Parameter.Add("@UserLogin")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString(), sp, Parameter, Type, arrQueryString)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
#End Region

End Class
