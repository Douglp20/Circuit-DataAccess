﻿Public Class StaffData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub

#Region "Bonus"
    Public Function getOrderBonusByOrderSearch(value As String) As SqlClient.SqlDataAdapter
        On Error GoTo Err



        Dim sp As String = "[Staff_get_order_bonus_by_order_search]"
        Dim Parameter As String = "@Search"
        Dim Type As String = SqlDbType.VarChar


        getOrderBonusByOrderSearch = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString(), sp, Parameter, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        If Err.Number > 0 Then RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

    Public Function getStaffBonusbyStaffID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_get_Bonus_by_staffID]"
        Dim Parameter As String = "@StaffID"
        Dim Type As String = SqlDbType.Int



        getStaffBonusbyStaffID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getStaffBonusbyID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_get_Bonus_by_ID]"
        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int



        getStaffBonusbyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Sub InsertStaffBonus(ByRef Value As ArrayList)

        On Error GoTo Err




        Dim sp As String = "[Insert_StaffBonus]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@StaffID")
        Parameter.Add("@OrderID")
        Parameter.Add("@Hour")
        Parameter.Add("@Rate")
        Parameter.Add("@Labour")
        Parameter.Add("@Total")
        Parameter.Add("@Reason")
        Parameter.Add("@option")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)





        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub UpdateStaffBonus(ByRef value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[Update_StaffBonus]"

        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@ID")
        Parameter.Add("@StaffID")
        Parameter.Add("@Hour")
        Parameter.Add("@Rate")
        Parameter.Add("@Labour")
        Parameter.Add("@Total")
        Parameter.Add("@Reason")
        Parameter.Add("@option")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)





        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub DeleteStaffBonus(ByRef Value As Integer)

        On Error GoTo Err




        Dim sp As String = "[Delete_StaffBonus]"

        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, Parameter, Type, Value)




        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region

#Region "Staff search and FormLoad"

    Public Function getAllStaffUsers() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Staff_get_all_Staff_users]"

        getAllStaffUsers = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getAllSubContractorDropDownList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Staff_get_all_subContractor_dropdown_list]"

        getAllSubContractorDropDownList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function zzzgetStaffEmployeePricework() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Staff_get_employee_pricework]"

        zzzgetStaffEmployeePricework = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function


    Public Function getStaffbySearch(searchValue As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_get_staff_by_search]"
        Dim Parameter As String = "@search"
        Dim strType As String = SqlDbType.VarChar
        Dim strQueryString As String = searchValue


        getStaffbySearch = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

#End Region
#Region "Staff Save"
    Public Function InsertStaffDetail(ByRef arrValue As ArrayList) As Integer

        On Error GoTo Err



        Dim arrQueryString As New ArrayList
        Dim sp As String = "[insert_staff]"
        Dim ParameterOutput As String = "@ID"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        For i As Integer = 1 To arrValue.Count - 1
            arrQueryString.Add(arrValue(i))
        Next

        arrParameter.Add("@Firstname")
        arrParameter.Add("@Surname")
        arrParameter.Add("@NI_no")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)




        InsertStaffDetail = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString, sp, arrParameter, ParameterOutput, arrType, arrQueryString)





        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateStaffDetailTest(ByRef value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_staff]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@id")
        Parameter.Add("@Firstname")
        Parameter.Add("@Surname")
        Parameter.Add("@Trade")
        Parameter.Add("@Address")
        Parameter.Add("@Vehicle_registration")
        Parameter.Add("@NI_no")
        Parameter.Add("@Tel_no")
        Parameter.Add("@Mobile")
        Parameter.Add("@UTR")
        Parameter.Add("@CIS")
        Parameter.Add("@email")
        Parameter.Add("@PAYE_start_date")
        Parameter.Add("@PAYE_end_date")
        Parameter.Add("@Post_code")
        Parameter.Add("@tester")
        Parameter.Add("@dob")
        Parameter.Add("@previous_employer")
        Parameter.Add("@next_of_kin")
        Parameter.Add("@next_of_kin_phone")
        Parameter.Add("@employee_type")
        Parameter.Add("@comp_reg_no")
        Parameter.Add("@comp_name")
        Parameter.Add("@tax_treatment")
        Parameter.Add("@vat_number")
        Parameter.Add("@disabled")
        Parameter.Add("@email_worksheets")
        Parameter.Add("@UserName")
        Parameter.Add("@Notes")
        Parameter.Add("@subContractor")
        Parameter.Add("@HourlyRate")
        Parameter.Add("@voidworker")
        Parameter.Add("@staff")
        Parameter.Add("@Administrator")
        Parameter.Add("@CertificateEngineer")

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
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateStaffDetail(ByRef value As ArrayList, ImageData As Byte())

        On Error GoTo Err

        Dim sp As String = "[update_staff]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList
        Dim pictureParameter As String = "@photo"


        Parameter.Add("@id")
        Parameter.Add("@Firstname")
        Parameter.Add("@Surname")
        Parameter.Add("@Trade")
        Parameter.Add("@Address")
        Parameter.Add("@Vehicle_registration")
        Parameter.Add("@NI_no")
        Parameter.Add("@Tel_no")
        Parameter.Add("@Mobile")
        Parameter.Add("@UTR")
        Parameter.Add("@CIS")
        Parameter.Add("@email")
        Parameter.Add("@PAYE_start_date")
        Parameter.Add("@PAYE_end_date")
        Parameter.Add("@Post_code")
        Parameter.Add("@tester")
        Parameter.Add("@dob")
        Parameter.Add("@previous_employer")
        Parameter.Add("@next_of_kin")
        Parameter.Add("@next_of_kin_phone")
        Parameter.Add("@employee_type")
        Parameter.Add("@comp_reg_no")
        Parameter.Add("@comp_name")
        Parameter.Add("@tax_treatment")
        Parameter.Add("@vat_number")
        Parameter.Add("@disabled")
        Parameter.Add("@email_worksheets")
        Parameter.Add("@UserName")
        Parameter.Add("@Notes")
        Parameter.Add("@subContractor")
        Parameter.Add("@HourlyRate")
        Parameter.Add("@voidworker")
        Parameter.Add("@staff")
        Parameter.Add("@Administrator")
        Parameter.Add("@CertificateEngineer")

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
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.SmallMoney)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)

        ViperCon.ExecuteProcessImageWithParameters(connection.ConnectionString, sp, Parameter, Type, value, ImageData, pictureParameter)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateStaffUserInfo(ByRef value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_staff_user_info]"


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@id")
        Parameter.Add("@Firstname")
        Parameter.Add("@Surname")
        Parameter.Add("@Address")
        Parameter.Add("@Tel_no")
        Parameter.Add("@Mobile")
        Parameter.Add("@email")
        Parameter.Add("@Post_code")
        Parameter.Add("@previous_employer")
        Parameter.Add("@next_of_kin")
        Parameter.Add("@next_of_kin_phone")
        Parameter.Add("@comp_name")
        Parameter.Add("@dob")
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
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function check_for_duplicate_before_update(ByRef arrValue As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_check_for_duplicate_before_update]"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList
        arrParameter.Add("@First_name")
        arrParameter.Add("@Surname")
        arrParameter.Add("@NI_no")


        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)


        check_for_duplicate_before_update = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValue)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Sub zzzUpdateStaffWeekDays(ByRef arrValue As ArrayList)
        On Error GoTo Err



        On Error GoTo Err

        Dim sp As String = "[insert_update_staff_workdays]"


        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@StaffID")
        arrParameter.Add("@Monday")
        arrParameter.Add("@Tuesday")
        arrParameter.Add("@Wednesday")
        arrParameter.Add("@Thursday")
        arrParameter.Add("@Friday")
        arrParameter.Add("@Saturday")
        arrParameter.Add("@Sunday")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValue)


        Exit Sub


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
    Public Sub SaveStaffWeekDay(ByRef Value As ArrayList)
        On Error GoTo Err



        On Error GoTo Err

        Dim sp As String = "[SAVE_staff_workday]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@StaffID")
        Parameter.Add("@Monday")
        Parameter.Add("@Tuesday")
        Parameter.Add("@Wednesday")
        Parameter.Add("@Thursday")
        Parameter.Add("@Friday")
        Parameter.Add("@Saturday")
        Parameter.Add("@Sunday")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
#End Region
#Region "Staff Load Data"

    Public Function getStaffbyID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_get_staff_by_id]"
        Dim Parameter As String = "@id"
        Dim Type As String = SqlDbType.Int



        getStaffbyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getPayrollBACSHistorybyStaffID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Payroll_get_BACS_history_by_StaffID]"
        Dim Parameter As String = "@StaffID"
        Dim Type As String = SqlDbType.Int



        getPayrollBACSHistorybyStaffID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getStaffUserInfoByLoginID(value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_get_staff_user_info_by_LoginID]"
        Dim Parameter As String = "@LoginId"
        Dim Type As String = SqlDbType.VarChar



        getStaffUserInfoByLoginID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

#End Region
#Region "Staff Absence"
    Public Function getAllStaffAbsenceUsers() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Staff_get_all_Staff_absence_users]"

        getAllStaffAbsenceUsers = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getStaffAbsenceBySearch(searchValue As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_get_staff_absence_by_search]"
        Dim Parameter As String = "@search"
        Dim strType As String = SqlDbType.VarChar
        Dim strQueryString As String = searchValue


        getStaffAbsenceBySearch = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getStaffHolidayInfoByDate(value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[StaffAbsence_get_absence_viewer]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@Month")
        Parameter.Add("@Year")

        Type.Add(SqlDbType.Char)
        Type.Add(SqlDbType.Int)


        getStaffHolidayInfoByDate = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getStaffInfoByLoginID(value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_get_info_by_LoginID]"
        Dim Parameter As String = "@LoginID"
        Dim Type As String = SqlDbType.VarChar



        getStaffInfoByLoginID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getStaffInfoByStaffID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_get_info_by_staffID]"
        Dim Parameter As String = "@StaffID"
        Dim Type As String = SqlDbType.VarChar



        getStaffInfoByStaffID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getPublicHolidayByDate(value As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[StaffAbsence_get_public_hoilday_By_date]"
        Dim Parameter As String = "@Date"
        Dim Type As String = SqlDbType.VarChar



        getPublicHolidayByDate = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getStaffAbsenceNew() As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[StaffAbsence_get_Absence_New]"
        'Dim Parameter As String = "@StaffID"
        'Dim Type As String = SqlDbType.Int



        getStaffAbsenceNew = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function getStaffAbsencebyYear(ByRef arrValuesPass As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[StaffAbsence_get_Absence_by_year]"

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList
        arrParameter.Add("@Year")
        arrParameter.Add("@StaffID")

        arrType.Add(SqlDbType.Char)
        arrType.Add(SqlDbType.Int)

        getStaffAbsencebyYear = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getStaffAbsenceViewer(ByRef Value As ArrayList) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[StaffAbsence_get_AbsenceView]"

        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList
        arrParameter.Add("@Year")
        arrParameter.Add("@Month")
        arrParameter.Add("@StaffID")
        arrParameter.Add("@Index")

        arrType.Add(SqlDbType.Char)
        arrType.Add(SqlDbType.Char)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)

        getStaffAbsenceViewer = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameter, arrType, Value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getAllStaffList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Staff_get_all_list_Users]"

        getAllStaffList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getStaffAbsenceDatebyID(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[StaffAbsence_get_AbsenceDate_by_AbsenceID]"
        Dim Parameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID


        getStaffAbsenceDatebyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getStaffDaysOfWeek(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[StaffAbsence_get_staff_workdays_By_StaffID]"
        Dim Parameter As String = "@Staffid"
        Dim Type As String = SqlDbType.Int



        getStaffDaysOfWeek = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Staff Absence Save Data"
    Public Function SaveStaffAbsence(ByRef Value As ArrayList) As Integer

        On Error GoTo Err




        Dim sp As String = "[save_staff_absence]"
        Dim ParameterOutput As String = "@OUT_ID"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@id")
        Parameter.Add("@StaffID")
        Parameter.Add("@startDate")
        Parameter.Add("@endDate")
        Parameter.Add("@NoOfday")
        Parameter.Add("@Notes")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Date)
        Type.Add(SqlDbType.Date)
        Type.Add(SqlDbType.Float)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)



        SaveStaffAbsence = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString, sp, Parameter, ParameterOutput, Type, Value)





        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateStaffAbsence(ByRef arrValuesPass As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_staff_absence]"


        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@id")
        arrParameter.Add("@StaffID")
        arrParameter.Add("@startDate")
        arrParameter.Add("@endDate")
        arrParameter.Add("@NoOfday")
        arrParameter.Add("@Notes")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Date)
        arrType.Add(SqlDbType.Date)
        arrType.Add(SqlDbType.SmallInt)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function InsertStaffAbsenceDate(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[insert_staff_absenceDate]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@staffabsenceID")
        Parameter.Add("@staffID")
        Parameter.Add("@dayName")
        Parameter.Add("@AbsenceDate")
        Parameter.Add("@daySession")


        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Date)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateStaffAbsenceApproval(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[Update_staff_absence_approval]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@ID")
        Parameter.Add("@ActionStatus")
        Parameter.Add("@Notes")
        Parameter.Add("@UserName")


        Type.Add(SqlDbType.Int)
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
#Region "Staff Absence Delete Data"
    Public Function DeleteStaffAbsenceDate(ByRef value As Integer, index As Integer)

        On Error GoTo Err

        Dim sp As String
        If index = 0 Then sp = "[delete_staff_absence]"
        If index = 1 Then sp = "[delete_staff_absenceDate]"
        Dim Parameter As String = "@ID"
        Dim strType As String = SqlDbType.Int


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, Parameter, strType, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region


End Class
