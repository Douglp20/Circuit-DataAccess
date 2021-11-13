Public Class WholesalersData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
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

        Dim sp As String = "[Wholesaler_get_all_wholesaler_list]"

        getAllWholesalers = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAllWholesalerDropdownList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Wholesaler_get_all_wholesaler_Dropdown_list]"

        getAllWholesalerDropdownList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getWholesalersbySearch(searchValue As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Wholesaler_get_wholesaler_by_search]"

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


        Dim sp As String = "[Wholesaler_get_wholesaler_by_ID]"
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


        Dim sp As String = "[Wholesaler_get_wholesaler_contact_by_wholesalerID]"
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


        Dim sp As String = "[Wholesaler_get_wholesaler_contact_by_ID]"
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
    Public Function getWholesalerOrderMaterialByID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Wholesaler_get_ordermaterial_by_ID]"
        Dim Parameter As String = "@Wholesalerid"
        Dim Type As String = SqlDbType.Int



        getWholesalerOrderMaterialByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function


    Public Function getSubContractortMaterial(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Wholesaler_get_wholesaler_by_advance_search]"
        Dim Parameter As New ArrayList
        Parameter.Add("@StaffID")
        Parameter.Add("@Year")
        Parameter.Add("@Month")


        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)

        getSubContractortMaterial = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

#End Region
#Region "Save Wholesaler"

    Public Sub DeleteWholesalerContactDetail(ID As Integer)
        On Error GoTo Err

        Dim SP As String = "[delete_wholesaler_contact]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Function InsertWholesalerDetail(ByRef arrValues As ArrayList) As Integer

        On Error GoTo Err

        Dim storeProcedure As String = "[insert_wholesaler_detail]"
        Dim strParameterOutput As String = "@ID"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@wholesaler")
        arrParameter.Add("@Address")
        arrParameter.Add("@Postcode")
        arrParameter.Add("@disabled")
        arrParameter.Add("@notes")
        arrParameter.Add("@Email")
        arrParameter.Add("@WWW")
        arrParameter.Add("@account_no")
        arrParameter.Add("@SageACRef")
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
        arrType.Add(SqlDbType.VarChar)



        InsertWholesalerDetail = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString, storeProcedure, arrParameter, strParameterOutput, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateWholesalerDetail(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim storeProcedure As String = "[update_wholesaler_detail]"
        Dim arrValuesPass As New ArrayList
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@id")
        arrParameter.Add("@wholesaler")
        arrParameter.Add("@Address")
        arrParameter.Add("@Postcode")
        arrParameter.Add("@disabled")
        arrParameter.Add("@notes")
        arrParameter.Add("@Email")
        arrParameter.Add("@WWW")
        arrParameter.Add("@account_no")
        arrParameter.Add("@SageACRef")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, storeProcedure, arrParameter, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save Contact Data"
    Public Function UpdateWholesalerContactDetail(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_wholesaler_contact_detail]"
        Dim arrValuesPass As New ArrayList

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        'Update a new record

        Parameter.Add("@id")
        Parameter.Add("@WholesalerID")
        Parameter.Add("@FirstName")
        Parameter.Add("@Surname")
        Parameter.Add("@Email")
        Parameter.Add("@Telephone")
        Parameter.Add("@Mobile")
        Parameter.Add("@City")
        Parameter.Add("@JobTitle")
        Parameter.Add("@Notes")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
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


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function InsertWholesalerContactDetail(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[insert_wholesaler_contact_detail]"
        Dim arrValuesPass As New ArrayList

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        'Insert a new record



        Parameter.Add("@WholesalerID")
        Parameter.Add("@FirstName")
        Parameter.Add("@Surname")
        Parameter.Add("@Email")
        Parameter.Add("@Telephone")
        Parameter.Add("@Mobile")
        Parameter.Add("@City")
        Parameter.Add("@JobTitle")
        Parameter.Add("@Notes")
        Parameter.Add("@UserName")

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


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region

#Region "Get City"

    Public Function getWholesalerCity() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Wholesaler_get_city]"



        getWholesalerCity = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
End Class
