Public Class TemplateData
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New douglas.Viper.Connection.Connection()
    Private connection As New Connection

    Public Sub New()
    End Sub


#Region "Error Control"
    Private Sub ErrorMessage_ViperCon(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String) Handles ViperCon.ErrorMessage
        Dim ErrMessage As String = ">> Called by the module : " + Me.ToString
        RaiseEvent ErrorMessage(errDesc, errNo, ErrMessage + vbCrLf + errTrace)
    End Sub
#End Region


#Region "Get Data"
    Public Function getTemplateByType(type As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Template_get_template_by_type]"
        Dim Parameter As String = "@TypeID"
        Dim SQLType As String = SqlDbType.Int


        getTemplateByType = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, SQLType, type)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getTemplateByID(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Template_get_template_by_id]"
        Dim Parameter As String = "@ID"
        Dim SQLType As String = SqlDbType.Int


        getTemplateByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, SQLType, ID)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getTemplateCompany() As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Template_get_all_company]"



        getTemplateCompany = ViperCon.getSqlDataAdapter(connection.ConnectionString, sp)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getTemplateCompanyByID(ID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Template_get_company_by_ID]"
        Dim Parameter As String = "@ID"
        Dim SQLType As String = SqlDbType.Int


        getTemplateCompanyByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, SQLType, ID)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getTemplateEmailDetailByID(id As Integer, index As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Template_get_emailDetail_by_ID]"
        Dim arrParameters As New ArrayList
        Dim arrSQLTypes As New ArrayList
        Dim arrQuery As New ArrayList

        arrParameters.Add("@ID")
        arrParameters.Add("@Index")

        arrSQLTypes.Add(SqlDbType.Int)
        arrSQLTypes.Add(SqlDbType.Int)

        arrQuery.Add(id)
        arrQuery.Add(index)


        getTemplateEmailDetailByID = ViperCon.getSqlDataAdapterWithParameters(connection.ConnectionString, sp, arrParameters, arrSQLTypes, arrQuery)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
#Region "Save Data"
    Public Sub InsertTemplate(arrQueryString As ArrayList)
        On Error GoTo Err

        Dim sp As String = "[insert_template]"
        Dim arrParameters As New ArrayList
        arrParameters.Add("@template")
        arrParameters.Add("@disabled")
        arrParameters.Add("@templateName")
        arrParameters.Add("@typeID")
        arrParameters.Add("@UserName")


        Dim arrTypes As New ArrayList
        arrTypes.Add(SqlDbType.Char)
        arrTypes.Add(SqlDbType.Bit)
        arrTypes.Add(SqlDbType.Char)
        arrTypes.Add(SqlDbType.Int)
        arrTypes.Add(SqlDbType.Char)




        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameters, arrTypes, arrQueryString)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub UpdateTemplate(arrQueryString As ArrayList)
        On Error GoTo Err

        Dim sp As String = "[update_template]"
        Dim arrParameters As New ArrayList
        arrParameters.Add("@ID")
        arrParameters.Add("@template")
        arrParameters.Add("@disabled")
        arrParameters.Add("@templateName")
        arrParameters.Add("@typeID")
        arrParameters.Add("@UserName")


        Dim arrTypes As New ArrayList
        arrTypes.Add(SqlDbType.Int)
        arrTypes.Add(SqlDbType.Char)
        arrTypes.Add(SqlDbType.Bit)
        arrTypes.Add(SqlDbType.Char)
        arrTypes.Add(SqlDbType.Int)
        arrTypes.Add(SqlDbType.Char)




        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, arrParameters, arrTypes, arrQueryString)



        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region



End Class
