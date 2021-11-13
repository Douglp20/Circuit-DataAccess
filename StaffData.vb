Public Class StaffData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New douglas.Viper.Connection.Connection()
    Private connection As New Connection
    Public Sub New()
    End Sub
    Private Sub ErrorMessage_event(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub

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


    Public Function getStaffbySearch(searchValue As String) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_get_staff_by_search]"
        Dim strParameter As String = "@search"
        Dim strType As String = SqlDbType.VarChar
        Dim strQueryString As String = searchValue


        getStaffbySearch = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



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
        Dim strParameterOutput As String = "@ID"
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




        InsertStaffDetail = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString, sp, arrParameter, strParameterOutput, arrType, arrQueryString)





        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateStaffDetail(ByRef arrValuesPass As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_staff]"


        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList



        arrParameter.Add("@id")
        arrParameter.Add("@Firstname")
        arrParameter.Add("@Surname")
        arrParameter.Add("@Trade")
        arrParameter.Add("@Address")
        arrParameter.Add("@Vehicle_registration")
        arrParameter.Add("@NI_no")
        arrParameter.Add("@Tel_no")
        arrParameter.Add("@Mobile")
        arrParameter.Add("@UTR")
        arrParameter.Add("@CIS")
        arrParameter.Add("@email")
        arrParameter.Add("@PAYE_start_date")
        arrParameter.Add("@PAYE_end_date")
        arrParameter.Add("@Post_code")
        arrParameter.Add("@tester")
        arrParameter.Add("@dob")
        arrParameter.Add("@previous_employer")
        arrParameter.Add("@next_of_kin")
        arrParameter.Add("@next_of_kin_phone")
        arrParameter.Add("@employee_type")
        arrParameter.Add("@comp_reg_no")
        arrParameter.Add("@comp_name")
        arrParameter.Add("@tax_treatment")
        arrParameter.Add("@vat_number")
        arrParameter.Add("@disabled")
        arrParameter.Add("@email_worksheets")
        arrParameter.Add("@UserName")
        arrParameter.Add("@Notes")
        arrParameter.Add("@subContractor")
        arrParameter.Add("@HourlyRate")
        arrParameter.Add("@voidworker")
        arrParameter.Add("@staffCheck")


        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.DateTime)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.SmallMoney)
        arrType.Add(SqlDbType.Bit)
        arrType.Add(SqlDbType.Bit)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


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
    Public Sub UpdateStaffWeekDays(ByRef arrValue As ArrayList)
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
#End Region
#Region "Staff Load Data"
    Public Function getStaffbyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Staff_get_staff_by_id]"
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getStaffbyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

#End Region
#Region "Staff Absence Load Data"
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
        Dim strParameter As String = "@id"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID


        getStaffAbsenceDatebyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getStaffDaysOfWeek(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[StaffAbsence_get_staff_workdays_By_StaffID]"
        Dim strParameter As String = "@Staffid"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = ID


        getStaffDaysOfWeek = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Staff Absence Save Data"
    Public Function InsertStaffAbsence(ByRef arrValue As ArrayList) As Integer

        On Error GoTo Err



        Dim arrQueryString As New ArrayList
        Dim sp As String = "[insert_staff_absence]"
        Dim strParameterOutput As String = "@ID"
        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList

        For i As Integer = 1 To arrValue.Count - 1
            arrQueryString.Add(arrValue(i))
        Next

        arrParameter.Add("@StaffID")
        arrParameter.Add("@startDate")
        arrParameter.Add("@endDate")
        arrParameter.Add("@NoOfday")
        arrParameter.Add("@Notes")
        arrParameter.Add("@UserName")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.Date)
        arrType.Add(SqlDbType.Date)
        arrType.Add(SqlDbType.SmallInt)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.VarChar)



        InsertStaffAbsence = ViperCon.ExecuteProcessWithParametersReturnInteger(connection.ConnectionString, sp, arrParameter, strParameterOutput, arrType, arrQueryString)





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

    Public Function InsertStaffAbsenceDate(ByRef arrValuesPass As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[insert_staff_absenceDate]"


        Dim arrParameter As New ArrayList
        Dim arrType As New ArrayList


        arrParameter.Add("@staffabsenceID")
        arrParameter.Add("@dayName")
        arrParameter.Add("@AbsenceDate")

        arrType.Add(SqlDbType.Int)
        arrType.Add(SqlDbType.VarChar)
        arrType.Add(SqlDbType.Date)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameter, arrType, arrValuesPass)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Staff Absence Delete Data"
    Public Function DeleteStaffAbsenceDate(ByRef queryString As String)

        On Error GoTo Err

        Dim sp As String = "[delete_staff_absenceDate]"


        Dim strParameter As String = "@staffabsenceID"
        Dim strType As String = SqlDbType.Int


        ViperCon.ExecuteProcessWithParameter(connection.ConnectionString, sp, strParameter, strType, queryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
End Class
