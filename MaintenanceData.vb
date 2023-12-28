Public Class MaintenanceData
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

#Region "Get ProjectType Data"
    Public Function getAllProjectTypeList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Maintenance_get_projectType]"

        getAllProjectTypeList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
#End Region
#Region "Get JobType Data"
    Public Function getAllJobTypeList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Maintenance_get_jobtype]"

        getAllJobTypeList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getAllOrderStatusList() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Maintenance_get_order_status]"

        getAllOrderStatusList = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function


    Public Function getAllJobTypeByCompanyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Maintenance_get_all_projectType_by_companyID]"
        Dim strParameter As String = "@companyID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getAllJobTypeByCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function


    Public Function getAllProjectTypeByCompanyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Maintenance_get_all_projectType_by_companyID]"
        Dim strParameter As String = "@companyID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getAllProjectTypeByCompanyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getJobTypebyCompanyID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err


        Dim sp As String = "[Maintenance_get_jobtype_by_companyID]"
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


        Dim sp As String = "[Maintenance_get_jobtype_by_ID]"
        Dim strParameter As String = "@ID"
        Dim strType As String = SqlDbType.Int
        Dim strQueryString As String = id


        getCompanyJobTypebyID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, strParameter, strType, strQueryString)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getJobTypeBYCompanyIDDropDownList(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Maintenance_get_jobType_By_companyID_dropdown_list]"


        Dim Parameter As String = "@CompanyID"
        Dim Type As String = SqlDbType.Int
        Dim strQueryString As String = id

        getJobTypeBYCompanyIDDropDownList = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, id)




        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region



#Region " Project Type Maintenance"

    Public Function getALLProjectType() As SqlClient.SqlDataAdapter
        On Error GoTo Err


        Dim sp As String = "[Maintenance_get_all_projectType]"



        getALLProjectType = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
#End Region
#Region "Job Type Maintenance"

    Public Function getALLJobType() As SqlClient.SqlDataAdapter
        On Error GoTo Err

        ''Company_get_jobType_by_contactID
        Dim sp As String = "[Maintenance_get_all_jobType]"



        getALLJobType = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

#End Region
#Region "Code Maintenance"

    Public Function getAllMaintenance(value As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err

        ''Company_get_jobType_by_contactID
        Dim sp As String = "[Maintenance_get_all_Maintenance_by_index]"


        Dim Parameter As String = "@index"
        Dim Type As String = SqlDbType.Int

        getAllMaintenance = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getMaintenanceByIndex(value As Integer) As SqlClient.SqlDataAdapter
        On Error GoTo Err

        ''Company_get_jobType_by_contactID
        Dim sp As String = "[Maintenance_get_by_index]"


        Dim Parameter As String = "@index"
        Dim Type As String = SqlDbType.Int

        getMaintenanceByIndex = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function getMaintenanceByCompany(value As ArrayList) As SqlClient.SqlDataAdapter
        On Error GoTo Err

        ''Company_get_jobType_by_contactID
        Dim sp As String = "[Maintenance_get_by_companyID]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@companyID")
        Parameter.Add("@companysubID")
        Parameter.Add("@index")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)

        getMaintenanceByCompany = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, Parameter, Type, value)

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

#End Region
#Region "Save Data"
    Public Function SaveMaintenance(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        sp = "[Save_maintenance]"

        Parameter.Add("@ID")
        Parameter.Add("@index")
        Parameter.Add("@Name")
        Parameter.Add("@disabled")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function SaveCityDetail(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        sp = "[Update_maintenance_city]"

        Parameter.Add("@ID")
        Parameter.Add("@City")
        Parameter.Add("@disabled")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function SaveJobTypeDetail(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        sp = "[Update_maintenance_jobtype]"

        Parameter.Add("@ID")
        Parameter.Add("@code")
        Parameter.Add("@desc")
        Parameter.Add("@disabled")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function SaveProjectTypeDetail(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        sp = "[Update_maintenance_projecttype]"

        Parameter.Add("@ID")
        Parameter.Add("@code")
        Parameter.Add("@desc")
        Parameter.Add("@disabled")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function SaveZoneDetail(ByRef Values As ArrayList)

        On Error GoTo Err

        Dim sp As String


        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        sp = "[Update_maintenance_zone]"

        Parameter.Add("@ID")
        Parameter.Add("@zone")
        Parameter.Add("@area")
        Parameter.Add("@rate")
        Parameter.Add("@disabled")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Money)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Values)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
End Class
