Public Class MaterialStockData
    Public Sub New()
    End Sub

    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private WithEvents ViperCon As New Douglas.Viper.Connection.Connection()
    Private connection As New Connection
#Region "Get Order Material"
    Public Function getMaterialBySearch(ByRef value As String) As SqlClient.SqlDataAdapter


        On Error GoTo Err


        Dim sp As String = "[MATERIAL_get_material_by_search]"
        Dim Parameter As String = "@search"
        Dim Type As String = SqlDbType.VarChar

        getMaterialBySearch = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function




    Public Function getOrderMaterialByID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_ordermaterial_by_ID]"
        Dim Parameter As String = "@id"
        Dim Type As String = SqlDbType.Int
        Dim QueryString As String = id


        getOrderMaterialByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, QueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function

    Public Function getOrderMaterialItemByID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_Order_material_item_by_materialID]"
        Dim Parameter As String = "@materialID"
        Dim Type As String = SqlDbType.Int
        Dim QueryString As String = id


        getOrderMaterialItemByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, QueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getWholesalerCityByID(id As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_wholesaler_city_wholesaleID]"
        Dim Parameter As String = "@id"
        Dim Type As String = SqlDbType.Int
        Dim QueryString As String = id


        getWholesalerCityByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, QueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getOrderMaterialsByOrderID(value As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_ordermaterials_by_OrderID]"
        Dim Parameter As String = "@Orderid"
        Dim Type As String = SqlDbType.Int



        getOrderMaterialsByOrderID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, value)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Function getOrderMaterialsByID(MaterialID As Integer) As SqlClient.SqlDataAdapter

        On Error GoTo Err

        Dim sp As String = "[Order_get_ordermaterial_by_ID]"
        Dim Parameter As String = "@ID"
        Dim Type As String = SqlDbType.Int
        Dim QueryString As String = MaterialID


        getOrderMaterialsByID = ViperCon.getSqlDataAdapterWithParameter(connection.ConnectionString, sp, Parameter, Type, QueryString)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Function
    Public Sub DeleteMaterial(ByRef Value As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[delete_OrderMaterial]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@ID")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, Value)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
#End Region


#Region "Save Material"


    Public Sub InsertMaterial(ByRef arrValue As ArrayList)
        On Error GoTo Err


        Dim sp As String = "[insert_Order_Material]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList



        Parameter.Add("@OrderID")
        Parameter.Add("@WholesalerID")
        Parameter.Add("@City")
        Parameter.Add("@StaffID")
        Parameter.Add("@InvoiceNo")
        Parameter.Add("@Amount")
        Parameter.Add("@Notes")
        Parameter.Add("@Invoicedate")
        Parameter.Add("@item")
        Parameter.Add("@InStockDate")
        Parameter.Add("@Status")
        Parameter.Add("@SubContractorPaidCheck")
        Parameter.Add("@SentCheck")
        Parameter.Add("@EmailMessage")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)




        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, arrValue)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Function UpdateMaterial(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_Order_Material]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList

        Parameter.Add("@id")
        Parameter.Add("@OrderID")
        Parameter.Add("@WholesalerID")
        Parameter.Add("@City")
        Parameter.Add("@StaffID")
        Parameter.Add("@InvoiceNo")
        Parameter.Add("@Amount")
        Parameter.Add("@Notes")
        Parameter.Add("@Invoicedate")
        Parameter.Add("@item")
        Parameter.Add("@InStockDate")
        Parameter.Add("@Status")
        Parameter.Add("@SubContractorPaidCheck")
        Parameter.Add("@SentCheck")
        Parameter.Add("@EmailMessage")
        Parameter.Add("@UserName")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.DateTime)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.VarChar)

        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function UpdateMaterialItem(ByRef arrValues As ArrayList)

        On Error GoTo Err

        Dim sp As String = "[update_Order_Material_item]"
        Dim Parameter As New ArrayList
        Dim Type As New ArrayList


        Parameter.Add("@id")
        Parameter.Add("@MaterialID")
        Parameter.Add("@Item")
        Parameter.Add("@Qty")
        Parameter.Add("@Delete")
        Parameter.Add("@UserLogin")

        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.VarChar)
        Type.Add(SqlDbType.Int)
        Type.Add(SqlDbType.Bit)
        Type.Add(SqlDbType.VarChar)


        ViperCon.ExecuteProcessWithParameters(connection.ConnectionString, sp, Parameter, Type, arrValues)


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region

End Class
