Public Class CompanyData
    Public Sub New()
    End Sub

    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Viper.Connection.Connection()
    Private connection As New Connection


#Region "Error Control"
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#End Region


#Region "Get Company Data"
    Public Function getAllCompanyDropdownlist() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Company_get_all_company_dropdown_list]"

        getAllCompanyDropdownlist = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAllCompanyList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Company_get_all_company_list]"

        getAllCompanyList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAllSubCompanyList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Company_get_all_Subcompany_list]"

        getAllSubCompanyList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanybyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_company_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getCompanybyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanybySearch(searchValue As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_company_by_search]"
        Dim strParameter As String = "@search"
        Dim strType As String = SqlDbType.VarChar
        Dim strQueryString As String = searchValue


        getCompanybySearch = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanyContactbyCompanyID(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_company_contact_by_companyID]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID


        getCompanyContactbyCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getCompanyContactbyID(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_company_contact_by_ID]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID


        getCompanyContactbyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

#End Region
#Region "Get Order Item"
    Public Function GetCompanyPricelistbycompanyIDList(companyID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Company_get_company_pricelist_by_companyID]"
        Dim strParameter As String = "@companyID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = companyID



        GetCompanyPricelistbycompanyIDList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Get JobType Data"
    Public Function getAllJobTypeDropDownList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Jobtype_get_all_jobtype_dropdown_list]"

        getAllJobTypeDropDownList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getJobTypebyCompanyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Jobtype_get_jobtype_by_companyID]"
        Dim strParameter As String = "@companyID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getJobTypebyCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanyJobTypebyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Jobtype_get_jobtype_by_ID]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getCompanyJobTypebyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getJobTypeBYCompanyIDDropDownList(companyID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Company_get_jobType_By_companyID_dropdown_list]"
        Dim strParameter As String = "@companyID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = companyID



        getJobTypeBYCompanyIDDropDownList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save Company Data"
    Public Function SaveCompanyDetail(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList
        ''Insert a snew record

        If arrValues(0) = 0 Then
            sp = "[insert_company]"
            For i As Integer = 1 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        Else
            arrParameter.Add("@id")
            arrType.Add(SqlDbType.Int)
            sp = "[update_company]"
            For i As Integer = 0 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        End If


        arrParameter.Add("@Company")
        arrParameter.Add("@Address")
        arrParameter.Add("@Postcode")
        arrParameter.Add("@worksheetnotes")
        arrParameter.Add("@abbreviation")
        arrParameter.Add("@companyType")
        arrParameter.Add("@reportToID")
        arrParameter.Add("@invoice_materials_percent")
        arrParameter.Add("@voids_nonvoids_reports")
        arrParameter.Add("@retention_req")
        arrParameter.Add("@levy_req")
        arrParameter.Add("@daily_appointment_reports")
        arrParameter.Add("@disabled")
        arrParameter.Add("@monthly_valuation_report")
        arrParameter.Add("@UserName")


        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region

#Region "Get Company Contact Data"
    Public Function SaveCompanyContactDetail(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList
        ''Insert a snew record

        If arrValues(0) = 0 Then
            sp = "[insert_company_contact_detail]"
            For i As Integer = 1 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        Else
            arrParameter.Add("@id")
            arrType.Add(SqlDbType.Int)
            sp = "[update_company_contact_detail]"
            For i As Integer = 0 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        End If


        arrParameter.Add("@CompanyID")
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
#Region "Save Company Job Type Data"
    Public Sub SaveCompanyJobTypeDetail(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList
        ''Insert a snew record

        If arrValues(0) = 0 Then
            sp = "[insert_company_jobtype]"
            For i As Integer = 1 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        Else
            arrParameter.Add("@id")
            arrType.Add(SqlDbType.Int)
            sp = "[update_company_jobtype]"
            For i As Integer = 0 To arrValues.Count - 1
                arrValuesPass.Add(arrValues(i))
            Next

        End If

        arrParameter.Add("@Companyid")
        arrParameter.Add("@jobtypeid")
        arrParameter.Add("@percentage")
        arrParameter.Add("@UserName")


        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
#Region "Get Compnay Price List"
    Public Function getCompanyPriceListbyCompanyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_company_pricelist_by_companyID_List]"
        Dim strParameter As String = "@companyID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getCompanyPriceListbyCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save Company Price List"
    Public Sub SaveCompanyPriceList(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList
        sp = "[insert_Update_company_PriceList]"
        For i As Integer = 0 To arrValues.Count - 1
            arrValuesPass.Add(arrValues(i))
        Next

        arrParameter.Add("@CompanyID")
        arrParameter.Add("@code")
        arrParameter.Add("@description")
        arrParameter.Add("@price")
        arrParameter.Add("@disprice")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
#Region "Delete Company Price List"
    Public Sub deleteCompanyPriceListCompanyID(id As Integer)

        On Error GoTo Err


        Dim sp As String = "[delete_company_priceList]"
        Dim strParameter As String = "@companyID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
End Class
