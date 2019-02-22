Imports System.Data.Sql
Imports System.Data.SqlClient

Public Module DA

#Region " ¤×¤×¤×¤×¤×¤×¤×¤ Constants ¤×¤×¤×¤×¤×¤×¤×¤ "

    Private Const mconNotConnected As String = "< Not Connected >"
    'Used to access the DB
    'Private Const mconAccessConnectString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};User Id=admin;Password=;" 'Text in braces {} is replaceable parameters
    Private Const mconAccessConnectString As String = "Driver={Microsoft Access Driver (*.mdb)};Dbq={0};Uid=Admin;Pwd=;"
    '*Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mydatabase.mdb;Jet OLEDB:System Database=system.mdw;User ID=myUsername;Password=myPassword;
    '*Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mydatabase.mdb;User Id=admin;Password=;
    '*Driver={Microsoft Access Driver (*.mdb)};Dbq=C:\mydatabase.mdb;Uid=Admin;Pwd=;
    Private Const mconSQLConnectString As String = "Data Source={0};Initial Catalog={1};User ID={2};Password={3};database={4};" 'Text in braces {} is replaceable parameters
    'Private Const mconSQLConnectStringOLE As String = "Provider=SQLNCLI10;Server={0};Database={1};Uid=sa; Pwd=navaho:spaces;" 'Text in braces {} is replaceable parameters
    'Private Const mconSQLConnectStringOLE As String = "Provider=SQLNCLI;Server={0};Database={1};Uid=sa; Pwd=navaho:spaces;" 'Text in braces {} is replaceable parameters
    Private Const mconSQLConnectStringOLE As String = "Provider=SQLOLEDB;Data Source={0};Initial Catalog={1};User ID={2};Password={3};"




#End Region

#Region " ¤×¤×¤×¤×¤×¤×¤×¤ Enumerations ¤×¤×¤×¤×¤×¤×¤×¤ "

    Public Enum eDatabaseType
        Access
        SQLServer
    End Enum

    Public Enum eLocking
        None
        Intent
        Full
    End Enum

#End Region

#Region " ¤×¤×¤×¤×¤×¤ Private variables ¤×¤×¤×¤×¤×¤ "

    Private meDataBase As eDatabaseType
    Private mstrDatabase As String
    Private mstrDatabasePath As String
    Private mstrServer As String
    Private mstrUID As String
    Private mstrPassword As String
    Private mstrErrorMessage As String
    Private mConn As SqlConnection

#End Region

#Region " ¤×¤×¤×¤×¤ Read only properties ¤×¤×¤×¤×¤ "

    Public ReadOnly Property Connected() As Boolean
        Get
            If mConn Is Nothing Then
                Return False
            Else
                Return mConn.ConnectionString <> vbNullString
            End If
        End Get
    End Property

    Public ReadOnly Property Connection() As String
        Get
            If mConn Is Nothing Then
                Return mconNotConnected
            Else
                Return mConn.ConnectionString
            End If
        End Get
    End Property

    Public ReadOnly Property ErrorMessage() As String
        Get
            Return mstrErrorMessage
        End Get
    End Property

#End Region

#Region " ¤×¤×¤×¤×¤×¤ Public properties ¤×¤×¤×¤×¤×¤ "

    Public Property DatabaseType() As eDatabaseType
        Get
            Return meDataBase
        End Get
        Set(ByVal value As eDatabaseType)
            meDataBase = value
        End Set
    End Property

    Public Property Database() As String
        Get
            Return mstrDatabase
        End Get
        Set(ByVal value As String)
            mstrDatabase = value
        End Set
    End Property

    Public Property DatabasePath() As String
        Get
            Return mstrDatabasePath
        End Get
        Set(ByVal value As String)
            mstrDatabasePath = value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return mstrPassword
        End Get
        Set(ByVal value As String)
            mstrPassword = value
        End Set
    End Property

    Public Property Server() As String
        Get
            Return mstrServer
        End Get
        Set(ByVal value As String)
            mstrServer = value
        End Set
    End Property

    Public Property UserID() As String
        Get
            Return mstrUID
        End Get
        Set(ByVal value As String)
            mstrUID = value
        End Set
    End Property

#End Region

    Public Sub Disconnect()

        If Not mConn Is Nothing Then
            mConn.Dispose()
            mConn = Nothing
        End If

    End Sub

    Public Function Init() As Boolean

        Dim strConnect As String


        strConnect = vbNullString

        Try

            If mConn Is Nothing Then

                mConn = New SqlConnection

            End If


            If mConn.ConnectionString = vbNullString Then

                Select Case meDataBase
                    Case eDatabaseType.Access
                        'strConnect = [String].Format(mconAccessConnectString, mstrDatabasePath)
                        strConnect = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & mstrDatabasePath & ";Uid=Admin;Pwd=;"
                    Case eDatabaseType.SQLServer
                        strConnect = [String].Format(mconSQLConnectString, mstrServer, mstrDatabase, mstrUID, mstrPassword, mstrDatabase)
                End Select

                With mConn
                    .ConnectionString = strConnect
                    .Open()
                    If .ConnectionString <> vbNullString Then
                        Return True
                    End If
                End With

            Else

                Return True

            End If

        Catch ex As Exception

            mConn = Nothing
            Return False

        End Try

    End Function

    Public Function ToString() As String

        Dim strConnection As String

        If Not mConn Is Nothing Then
            strConnection = mConn.ConnectionString
        Else
            strConnection = "(Not connected)"
        End If

        Return "Encapsulates all the SQL.  Connection = " & strConnection

    End Function

    Public Function GetActiveServers() As String()

        Dim intCount As Integer
        Dim intIndex As Integer
        Dim strServer As String()
        Dim Enumerator As SqlDataSourceEnumerator
        Dim Row As DataRow
        Dim DT As DataTable


        Enumerator = SqlDataSourceEnumerator.Instance
        DT = Enumerator.GetDataSources()
        intCount = DT.Rows.Count - 1

        If DT.Rows.Count > 0 Then

            ReDim strServer(intCount)

            For Each Row In DT.Rows

                If Not String.IsNullOrEmpty(Row("InstanceName").ToString()) Then
                    strServer(intIndex) = String.Format("{0}\{1}", Row("ServerName"), Row("InstanceName"))
                Else
                    strServer(intIndex) = Row("ServerName").ToString()
                End If

                intIndex += 1

            Next

            Return strServer

        Else

            Return Nothing

        End If

    End Function

    Public Function GetDatabases(ByVal strServer As String) As String()

        Dim intCount As Integer
        Dim intIndex As Integer
        Dim strDatabase As String()
        Dim Row As DataRow
        Dim DT As DataTable
        Dim Conn As SqlConnection


        Conn = New SqlConnection

        Try

            With Conn
                .ConnectionString = [String].Format(mconSQLConnectString, strServer, "", mstrUID, mstrPassword, "")
                .Open()
                DT = .GetSchema(SqlClientMetaDataCollectionNames.Databases)
            End With

            intCount = DT.Rows.Count - 1

            If intCount > 0 Then

                ReDim strDatabase(intCount)

                For Each Row In DT.Rows

                    strDatabase(intIndex) = String.Format("{0}", Row(0))
                    intIndex += 1

                Next

                Return strDatabase

            Else

                Return Nothing

            End If

        Catch ex As Exception

            MsgBox("Error retrieving databases: " & ex.Message)
            Return Nothing

        End Try

    End Function

    Public Sub GetFields(ByVal strTable As String, ByRef colFields As Depot.Tables.Table.Fields)

        Dim strTableName As String()
        Dim Field As Depot.Tables.Table.Fields.Field
        Dim Row As DataRow
        Dim FieldList As DataTable


        If mConn.ConnectionString <> vbNullString Then

            ReDim strTableName(3)
            strTableName(2) = strTable
            'Catalog, Schema, Table, Column
            FieldList = mConn.GetSchema("Columns", strTableName)
            colFields = New Depot.Tables.Table.Fields

            For Each Row In FieldList.Rows

                '0 = catalog; 1 = schema; 2 = Table Name; 3 = Column Name; 4 = Ordinal; 5 = Default; 6 = Nullable; 7 = Datatype; 8 = MaxLength;
                '   9 = Octet length; 10 = Precision; 11 = NP Radix; 12 = Scale; 13 = DataTime Precision; 14 = CSC; 15 = CSS; 16 = CSN;
                '   17 colation 
                Field = New Depot.Tables.Table.Fields.Field
                With Field
                    ''Debug.Print(Row(3).ToString & "-" & Row(8).ToString)
                    'If TypeOf Row(8) Is Integer Then
                    '    .DataLength = CType(Row(8), Integer)
                    'End If
                    .DBDataType = Row(7).ToString
                    'If meDataBase = eDatabaseType.SQLServer Then
                    '    .VBDataType = DataTypeMappingSQLServer(.DBDataType)
                    'ElseIf meDataBase = eDatabaseType.Access Then
                    '    .VBDataType = DataTypeMappingAccess(.DBDataType)
                    'End If
                    .Name = Row(3).ToString
                    .Ordinal = CType(Row(4), Integer)
                    '.Required = Not Row(6).ToString = "YES"

                    ''Fields that are uniqueidentifiers are either PKeys or FKeys; make then read-only
                    'If .DBDataType = "uniqueidentifier" Then
                    '    .ReadOnly = True
                    'End If

                    '.UseQuotes = CheckDataTypeUseQuotes(.DBDataType)
                End With

                colFields.Add(Field)

            Next

            ''Retrieve Primary Key information
            'Dim r As Integer
            'Dim strColumn As String
            'Dim Conn As System.Data.OleDb.OleDbConnection
            'Dim DT As DataTable


            'Conn = New System.Data.OleDb.OleDbConnection
            'Conn.ConnectionString = String.Format(mconSQLConnectStringOLE, mstrServer, mstrDatabase, mstrUID, mstrPassword)
            'Conn.Open()

            'DT = Conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Primary_Keys, _
            '    New Object() {Nothing, Nothing, strTable})

            'For r = 0 To DT.Rows.Count - 1
            '    strColumn = DT.Rows(r)!COLUMN_NAME.ToString
            '    colFields(strColumn).PKey = True
            '    colFields(strColumn).UseAsLookup = True
            'Next

            'Conn.Close()

        End If

    End Sub

    Public Function GetStoredProceduress() As List(Of String)

        Dim strSQL As String
        Dim lisStoredProceduresNames As List(Of String) = Nothing
        Dim Reader As System.Data.SqlClient.SqlDataReader = Nothing
        Dim cmd As System.Data.SqlClient.SqlCommand


        Try

            strSQL = "SELECT [name]" & _
                    " FROM SYS.OBJECTS" & _
                    " WHERE type = 'P'" & _
                    " ORDER BY [name]"
            cmd = New System.Data.SqlClient.SqlCommand(strSQL, mConn)
            cmd.CommandTimeout = 0
            Reader = cmd.ExecuteReader()
            If Reader.Read Then
                lisStoredProceduresNames = New List(Of String)

                Do
                    lisStoredProceduresNames.Add(Reader(0).ToString)
                Loop While Reader.Read

            End If

            Return lisStoredProceduresNames


        Catch ex As Exception

            MsgBox("Error getting stored procedures: " & ex.Message)
            Return Nothing

        Finally

            Reader.Close()

        End Try

    End Function

    Public Function GetStoredProceduressDetails() As List(Of ItemInformation)

        Dim strSQL As String
        Dim StoredProceduresDetail As ItemInformation
        Dim lisStoredProceduresDetail As List(Of ItemInformation) = Nothing
        Dim Reader As System.Data.SqlClient.SqlDataReader = Nothing
        Dim cmd As System.Data.SqlClient.SqlCommand


        Try

            strSQL = "SELECT [name], OBJECT_DEFINITION(OBJECT_ID)" & _
                    " FROM SYS.OBJECTS" & _
                    " WHERE type = 'P'" & _
                    " ORDER BY [name]"
            cmd = New System.Data.SqlClient.SqlCommand(strSQL, mConn)
            cmd.CommandTimeout = 0
            Reader = cmd.ExecuteReader()
            If Reader.Read Then
                lisStoredProceduresDetail = New List(Of ItemInformation)

                Do
                    StoredProceduresDetail = New ItemInformation
                    With StoredProceduresDetail
                        .Definition = Reader(1).ToString
                        .Name = Reader(0).ToString
                    End With
                    lisStoredProceduresDetail.Add(StoredProceduresDetail)
                Loop While Reader.Read

            End If

            Return lisStoredProceduresDetail


        Catch ex As Exception

            MsgBox("Error getting stored procedures detail: " & ex.Message)
            Return Nothing

        Finally

            If Reader IsNot Nothing Then
                Reader.Close()
            End If

        End Try

    End Function

    Public Function GetTables() As String()

        Dim intIndex As Integer
        Dim strTables As String()
        Dim Row As DataRow
        Dim TableList As DataTable


        If mConn.ConnectionString <> vbNullString Then

            TableList = mConn.GetSchema("Tables")

            ReDim strTables(TableList.Rows.Count)

            For Each Row In TableList.Rows
                '0 = catalog; 1 = schema; 2 = Name; 3 = Type
                If Row.Item(3).ToString = "BASE TABLE" Then
                    strTables(intIndex) = Row.Item(2).ToString
                    intIndex += 1
                    'Debug.Print(intIndex.ToString)
                End If
            Next

            ReDim Preserve strTables(intIndex - 1)
            Return strTables

        Else

            Return Nothing

        End If

    End Function

    Public Function GetViews() As List(Of String)

        Dim strSQL As String
        Dim lisViewNames As List(Of String) = Nothing
        Dim Reader As System.Data.SqlClient.SqlDataReader = Nothing
        Dim cmd As System.Data.SqlClient.SqlCommand


        Try

            strSQL = "SELECT [name]" & _
                    " FROM SYS.VIEWS" & _
                    " ORDER BY [name]"
            cmd = New System.Data.SqlClient.SqlCommand(strSQL, mConn)
            cmd.CommandTimeout = 0
            Reader = cmd.ExecuteReader()
            If Reader.Read Then
                lisViewNames = New List(Of String)

                Do
                    lisViewNames.Add(Reader(0).ToString)
                Loop While Reader.Read

            End If

            Return lisViewNames


        Catch ex As Exception

            MsgBox("Error getting views: " & ex.Message)
            Return Nothing

        Finally

            Reader.Close()

        End Try

    End Function

    Public Function GetViewsDetails() As List(Of ItemInformation)

        Dim strSQL As String
        Dim ViewDetail As ItemInformation
        Dim lisViewDetails As List(Of ItemInformation) = Nothing
        Dim Reader As System.Data.SqlClient.SqlDataReader = Nothing
        Dim cmd As System.Data.SqlClient.SqlCommand


        Try

            strSQL = "SELECT SV.name, definition " & _
                        "FROM sys.all_sql_modules A " & _
                            "LEFT JOIN sys.views SV ON a.object_id = SV.object_id " & _
                        "WHERE A.object_id IN (SELECT object_id FROM SYS.VIEWS) " & _
                        "ORDER BY SV.name"
            cmd = New System.Data.SqlClient.SqlCommand(strSQL, mConn)
            cmd.CommandTimeout = 0
            Reader = cmd.ExecuteReader()
            If Reader.Read Then
                lisViewDetails = New List(Of ItemInformation)

                Do
                    ViewDetail = New ItemInformation
                    With ViewDetail
                        .Definition = (Reader(1).ToString)
                        .Name = Reader(0).ToString
                    End With
                    lisViewDetails.Add(ViewDetail)
                Loop While Reader.Read

            End If

            Return lisViewDetails


        Catch ex As Exception

            MsgBox("Error getting views: " & ex.Message)
            Return Nothing

        Finally

            Reader.Close()

        End Try

    End Function

    Public Function SearchData(ByVal strTable As String, ByVal strField As String, ByVal strSearchString As String) As Boolean?


        Dim strSQL As String
        Dim Reader As System.Data.SqlClient.SqlDataReader = Nothing
        Dim cmd As System.Data.SqlClient.SqlCommand


        Try

            strSQL = "SELECT [" & strField & "]" & _
                    " FROM " & strTable & _
                    " WHERE [" & strField & "] LIKE '" & strSearchString & "'"
            cmd = New System.Data.SqlClient.SqlCommand(strSQL, mConn)
            cmd.CommandTimeout = 0
            Reader = cmd.ExecuteReader()
            If Reader.Read Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception

            MsgBox("Error searching databases: " & ex.Message)
            Return Nothing

        Finally

            Reader.Close()

        End Try

    End Function

    Public Function SearchViews(ByVal strSearchString As String) As List(Of String)


        Dim strSQL As String
        Dim lisFound As List(Of String) = Nothing
        Dim Reader As System.Data.SqlClient.SqlDataReader = Nothing
        Dim cmd As System.Data.SqlClient.SqlCommand


        Try

            strSQL = "SELECT SV.name, definition " & _
                        "FROM sys.all_sql_modules A " & _
                            "LEFT JOIN sys.views SV ON a.object_id = SV.object_id " & _
                        "WHERE A.object_id IN (SELECT object_id FROM SYS.VIEWS) " & _
                            " AND definition LIKE '%" & strSearchString & "%'"
            cmd = New System.Data.SqlClient.SqlCommand(strSQL, mConn)
            cmd.CommandTimeout = 0
            Reader = cmd.ExecuteReader()

            If Reader.Read Then
                lisFound = New List(Of String)

                Do
                    lisFound.Add(Reader(0).ToString)
                Loop While Reader.Read

            End If

            Return lisFound


        Catch ex As Exception

            MsgBox("Error searching databases: " & ex.Message)
            Return Nothing

        Finally

            Reader.Close()

        End Try

    End Function

End Module
