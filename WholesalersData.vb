Public Class WholesalersData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#Region "Get Data"
    Public Function getAllWholesalers() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Wholesale_get_all_wholesalers_list]"

        getAllWholesalers = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAllWholesalerDropdownList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Wholesale_get_all_wholesale_Dropdown_list]"

        getAllWholesalerDropdownList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getWholesalersbySearch(searchValue As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Wholesaler_get_wholesalers_by_search]"

        Dim strParameter As String = "@search"
        Dim strType As String = SqlDbType.VarChar
        Dim strQueryString As String = searchValue


        getWholesalersbySearch = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getWholesalersbyID(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Wholesale_get_wholesalers_by_ID]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID


        getWholesalersbyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Get Wholesaler Contact"
    Public Function getWholesalerContactbyWholesalerID(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Wholesale_get_wholesaler_contact_by_wholesalerID]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID


        getWholesalerContactbyWholesalerID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getWholesalerContactbyID(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Wholesale_get_wholesaler_contact_by_ID]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID


        getWholesalerContactbyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Get Wholesaler Material"
    Public Function getWholesalerOrderMaterialByID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Wholesaler_get_ordermaterial_by_ID]"
        Dim strParameter As String = "@Wholesalerid"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getWholesalerOrderMaterialByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
#Region "Save Wholesaler"
    Public Function SaveWholesalerDetail(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList
        ''Insert a new record

        If arrValues(0) = 0 Then
            sp = "[insert_wholesaler_detail]"
            For i As Integer = 1 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        Else
            arrParameter.Add("@id")
            arrType.Add(SqlDbType.Int)
            sp = "[update_wholesaler_detail]"
            For i As Integer = 0 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        End If


        arrParameter.Add("@wholesaler")
        arrParameter.Add("@Address")
        arrParameter.Add("@Postcode")
        arrParameter.Add("@disabled")
        arrParameter.Add("@notes")
        arrParameter.Add("@Email")
        arrParameter.Add("@WWW")
        arrParameter.Add("@account_no")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save Contact Data"
    Public Function SaveWholesalerContactDetail(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList
        'Insert a new record

        If arrValues(0) = 0 Then
            sp = "[insert_wholesaler_contact_detail]"
            For i As Integer = 1 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        Else
            arrParameter.Add("@id")
            arrType.Add(SqlDbType.Int)
            sp = "[update_wholesaler_contact_detail]"
            For i As Integer = 0 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        End If


        arrParameter.Add("@WholesalerID")
        arrParameter.Add("@FirstName")
        arrParameter.Add("@Surname")
        arrParameter.Add("@Email")
        arrParameter.Add("@Telephone")
        arrParameter.Add("@Mobile")
        arrParameter.Add("@JobTitle")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
End Class
