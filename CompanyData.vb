﻿Public Class CompanyData
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

#Region " Sub Company"
    Public Function InsertCompanyBranch(ByRef arrValues As ArrayList) As Integer

        On Error GoTo Err


        Dim storeProcedure As String = "[insert_company_branch]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Dim ParameterOutput As String = "@ID"


        Parameter.Add("@Company")
        Parameter.Add("@CompanyID")
        Parameter.Add("@Address")
        Parameter.Add("@Postcode")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        InsertCompanyBranch = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString, storeProcedure, Parameter, ParameterOutput, Type, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateCompanyBranch(ByRef arrValues As ArrayList)

        On Error GoTo Err


        Dim arrValuesPass As New ArrayList

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Dim SP As String = "[update_company_branch]"



        Parameter.Add("@id")
        Parameter.Add("@TypeIndex")
        Parameter.Add("@Company")
        Parameter.Add("@Address")
        Parameter.Add("@Postcode")
        Parameter.Add("@notes")
        Parameter.Add("@batchOrderLimit")
        Parameter.Add("@SageACRef")
        Parameter.Add("@batchExcelRef")
        Parameter.Add("@invoice_materials_percent")
        Parameter.Add("@disabled")
        Parameter.Add("@voids_nonvoids_reports")
        Parameter.Add("@retention_req")
        Parameter.Add("@levy_req")
        Parameter.Add("@daily_appointment_reports")
        Parameter.Add("@correctionAppChecked")
        Parameter.Add("@monthly_valuation_report")
        Parameter.Add("@vat_charge")
        Parameter.Add("@portal_update")
        Parameter.Add("@batch_invoice")
        Parameter.Add("@emailCancelledCheck")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Text)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, SP, Parameter, Type, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanyBranchByCompanyID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_company_branch_by_CompanyID]"
        Dim Parameter As String = "@companyid"
        Dim Type As String = SqlDbType.Int



        getCompanyBranchByCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function



    Public Function getCompanyBranchByID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_company_branch_by_id]"
        Dim Parameter As String = "@branchID"
        Dim Type As String = SqlDbType.Int



        getCompanyBranchByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region

#Region "Get Company email Data"

    Public Function getCompanyEmailPicture(ByRef value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[company_get_company_email_picture_by_ID]"

        Dim Parameter As String = "@CompanyEmailID"

        Dim Type As String = SqlDbType.Int

        getCompanyEmailPicture = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Function getAllCompanyEmail() As SqlClient.SqlDataAdapter
        On Error GoTo Err

        On Error GoTo Err

        Dim sp As String = "[company_get_company_email]"

        getAllCompanyEmail = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function


Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region

#Region "Get Company Data"
    Public Function getAllCompanyDropdownlist(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Company_get_all_company_dropdown_list]"
        Dim Parameter As String = "@index"
        Dim Type As String = SqlDbType.Int


        getAllCompanyDropdownlist = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanyJobTypePercentage() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Company_get_jobType_percentage]"

        getCompanyJobTypePercentage = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanyByOrderStatus(value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Company_get_company_by_Order_Status]"

        Dim Parameter As String = "@status"
        Dim Type As String = SqlDbType.VarChar


        getCompanyByOrderStatus = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getAllCompanyList(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Company_get_all_company_list]"

        Dim Parameter As String = "@index"
        Dim Type As String = SqlDbType.Int


        getAllCompanyList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAllCompany() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Company_get_all_company]"



        getAllCompany = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getCompanybyID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_company_by_id]"
        Dim Parameter As String = "@id"
        Dim Type As String = SqlDbType.Int



        getCompanybyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getCompanyJobTypebyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_company_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getCompanyJobTypebyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



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
    Public Function getCompanyContactbyCompanyID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_company_contact_by_companyID]"
        Dim Parameter As String = "@companyID"
        Dim Type As String = SqlDbType.Int

        getCompanyContactbyCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)

        'Dim Parameter As New ArrayList
        'Dim Type As New ArrayList




        'Parameter.Add("@companyID")
        'Parameter.Add("@companySubID")
        'Parameter.Add("@Type")

        'Type.Add(SqlDbType.Int)
        'Type.Add(SqlDbType.Int)
        'Type.Add(SqlDbType.VarChar)


        'getCompanyContactbyCompanyID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)



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
    Public Function getBatchExcelRef() As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_application_excel_template_ref]"



        getBatchExcelRef = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)



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
#Region "Get Project Type"

    Public Function getCompanyProjectTypeByCompanyID(value As ArrayList) As SqlClient.SqlDataAdapter
        On Error GoTo Err


        Dim sp As String = "[Company_get_projectType_by_companyID]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@CompanyID")
        Parameter.Add("@CompanySubID")
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)

        getCompanyProjectTypeByCompanyID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#Region "Get Job Type"

    Public Function getCompanyJobTypeListByContactID(id As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err

        ''Company_get_jobType_by_contactID
        Dim sp As String = "[Company_get_jobType_by_contactID]"
        Dim strParameter As String = "@ContactID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getCompanyJobTypeListByContactID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getCompanyJobTypeByCompanyID(value As ArrayList) As SqlClient.SqlDataAdapter
        On Error GoTo Err


        Dim sp As String = "[Company_get_jobType_by_companyID]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@CompanyID")
        Parameter.Add("@CompanySubID")
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        getCompanyJobTypeByCompanyID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#Region "Save Company Data"
    Public Function InsertCompany(ByRef arrValues As ArrayList) As Integer

        On Error GoTo Err


        Dim storeProcedure As String = "[insert_company]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList
        Dim strParameterOutput As String = "@ID"


        arrParameter.Add("@Company")
        arrParameter.Add("@Address")
        arrParameter.Add("@Postcode")
        arrParameter.Add("@UserName")


        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)

        arrType.Add(SqlDbType.VarChar)


        InsertCompany = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString, storeProcedure, arrParameter, strParameterOutput, arrType, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateCompanyDetail(ByRef arrValues As ArrayList)

        On Error GoTo Err


        Dim arrValuesPass As New ArrayList

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Dim SP As String = "[update_company]"



        Parameter.Add("@id")
        Parameter.Add("@Company")
        Parameter.Add("@Address")
        Parameter.Add("@Postcode")
        Parameter.Add("@TypeIndex")
        Parameter.Add("@VORequestLimit")
        Parameter.Add("@worksheetnotes")
        Parameter.Add("@notes")
        Parameter.Add("@invoice_materials_percent")
        Parameter.Add("@voids_nonvoids_reports")
        Parameter.Add("@retention_req")
        Parameter.Add("@levy_req")
        Parameter.Add("@daily_appointment_reports")
        Parameter.Add("@disabled")
        Parameter.Add("@monthly_valuation_report")
        Parameter.Add("@vat_charge")
        Parameter.Add("@portal_update")
        Parameter.Add("@batch_invoice")
        Parameter.Add("@batchOrderLimit")
        Parameter.Add("@SageACRef")
        Parameter.Add("@batchExcelRef")
        Parameter.Add("@correctionAppChecked")
        Parameter.Add("@emailCancelledCheck")
        Parameter.Add("@SubCompanyChecked")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, SP, Parameter, Type, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region

#Region "Get Company Contact Data"
    Public Function InsertContactDetail(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String

        sp = "[insert_company_contact_detail]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList




        Parameter.Add("@CompanyID")
        Parameter.Add("@CompanySubID")
        Parameter.Add("@FirstName")
        Parameter.Add("@Surname")
        Parameter.Add("@Email")
        Parameter.Add("@Telephone")
        Parameter.Add("@Mobile")
        Parameter.Add("@JobTitle")
        Parameter.Add("@Notes")
        Parameter.Add("@Type")
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

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateContactDetail(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String
        Dim arrValuesPass As New ArrayList

        sp = "[update_company_contact_detail]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@id")
        Parameter.Add("@CompanyID")
        Parameter.Add("@CompanySubID")
        Parameter.Add("@FirstName")
        Parameter.Add("@Surname")
        Parameter.Add("@Email")
        Parameter.Add("@Telephone")
        Parameter.Add("@Mobile")
        Parameter.Add("@JobTitle")
        Parameter.Add("@Notes")
        Parameter.Add("@Type")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
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


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save and  Company Project and JobType Data"
    Public Sub DeleteCompanyJobTypeDetail(ByRef value As ArrayList)

        On Error GoTo Err
        Dim sp As String = "[delete_company_jobtype]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@CompanyID")
        Parameter.Add("@CompanySubID")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)

        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub

    Public Sub DeleteCompanyContractType(ByRef value As ArrayList)

        On Error GoTo Err
        Dim sp As String = "[delete_company_contracttype]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@CompanyID")
        Parameter.Add("@CompanySubID")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)

        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub InsertCompanyJobType(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[insert_company_jobtype]"


        Dim Parameter As New ArrayList
        Parameter.Add("@companyid")
        Parameter.Add("@companySubid")
        Parameter.Add("@jobtypeid")
        Parameter.Add("@percentage")
        Parameter.Add("@WorksheetNotes")
        Parameter.Add("@VORequestLimit")
        Parameter.Add("@UserName")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub SaveCompanyProjectType(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[insert_company_projectType]"


        Dim Parameter As New ArrayList
        Parameter.Add("@companyid")
        Parameter.Add("@CompanySubid")
        Parameter.Add("@projecttypeid")
        Parameter.Add("@UserName")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub

    Public Sub DeleteCompanyContactJobTypeDetail(ByRef id As Integer)

        On Error GoTo Err
        Dim sp As String = "[delete_company_contact_jobtype]"
        Dim strParameter As String = "@ContactID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)

        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub SaveCompanyContactJobType(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[insert_company_contact_jobtype]"


        Dim arrParameter As New ArrayList
        arrParameter.Add("@Contactid")
        arrParameter.Add("@jobtypeid")
        arrParameter.Add("@UserName")

        Dim arrType As New ArrayList
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValues)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub DeleteCompanyContactDetail(ByRef id As Integer)

        On Error GoTo Err
        Dim sp As String = "[delete_company_contact]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id



        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)

        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
#Region "Get Compnay Price List"
    Public Function getCompanyContractPriceList(Value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_company_Contract_pricelist]"
        Dim Parameter As New ArrayList

        Parameter.Add("@companyID")
        Parameter.Add("@CompanySubID")
        Parameter.Add("@projectID")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)

        getCompanyContractPriceList = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save Company Price List"
    Public Sub SaveCompanyContractPriceList(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[save_company_Contract_PriceList]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@CompanyID")
        Parameter.Add("@CompanySubID")
        Parameter.Add("@projectID")
        Parameter.Add("@code")
        Parameter.Add("@Short")
        Parameter.Add("@medium")
        Parameter.Add("@Rate")
        Parameter.Add("@discount")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.VarChar)




        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
#Region "Delete Company Price List"
    Public Sub deleteCompanyContractPriceList(Value As ArrayList)

        On Error GoTo Err


        Dim sp As String = "[delete_company_contract_priceList]"
        Dim Parameter As New ArrayList

        Parameter.Add("@companyID")
        Parameter.Add("@CompanySubID")
        Parameter.Add("@projectID")

        Dim Type As New ArrayList
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
#Region "Save Company CompareStatement"

    Public Function getCompanyCompareStatement(Value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Company_get_compareStatement_by_ID]"
        Dim arrParameter As New ArrayList

        Dim Parameter As String = "@companyID"

        Dim Type As String = SqlDbType.Int


        getCompanyCompareStatement = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, Value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Sub SaveCompanyCompareStatement(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[insert_company_Compare_Statement]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@CompanyID")
        Parameter.Add("@OrderNumber")
        Parameter.Add("@JobNumber")
        Parameter.Add("@AltJobNumber")
        Parameter.Add("@OrderDate")
        Parameter.Add("@Address")
        Parameter.Add("@Code")
        Parameter.Add("@Description")
        Parameter.Add("@Quantity")
        Parameter.Add("@TotalValue")
        Parameter.Add("@POCompletionDate")
        Parameter.Add("@ClaimedValue")
        Parameter.Add("@sort")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)



        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub

    Public Sub CompanyCompareStatement(value As Integer)

        On Error GoTo Err


        Dim sp As String = "[Update_company_Compare_Statement]"

        Dim Parameter As String = "@companyID"

        Dim Type As String = SqlDbType.Int


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub deleteCompanyCompareStatement(value As Integer)

        On Error GoTo Err


        Dim sp As String = "[Delete_company_Compare_Statement]"

        Dim Parameter As String = "@companyID"

        Dim Type As String = SqlDbType.Int


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
End Class
