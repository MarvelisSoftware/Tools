Imports System.Data.Sql
Imports System.Data.SqlClient

Public Module DA

#Region " ¤×¤×¤×¤×¤×¤×¤×¤ Constants ¤×¤×¤×¤×¤×¤×¤×¤ "

    Private Const mconNotConnected As String = "< Not Connected >"
    'Used to access the DB
    Private Const mconSQLConnectString As String = "Data Source={0};Initial Catalog={1};User ID={2};Password={3};database={4};" 'Text in braces {} is replaceable parameters
    'Private Const mconSQLConnectStringOLE As String = "Provider=SQLNCLI10;Server={0};Database={1};Uid=sa; Pwd=navaho:spaces;" 'Text in braces {} is replaceable parameters
    'Private Const mconSQLConnectStringOLE As String = "Provider=SQLNCLI;Server={0};Database={1};Uid=sa; Pwd=navaho:spaces;" 'Text in braces {} is replaceable parameters
    Private Const mconSQLConnectStringOLE As String = "Provider=SQLOLEDB;Data Source={0};Initial Catalog={1};User ID={2};Password={3};"




#End Region

#Region " ¤×¤×¤×¤×¤×¤×¤×¤ Enumerations ¤×¤×¤×¤×¤×¤×¤×¤ "

    Public Enum eLocking
        None
        Intent
        Full
    End Enum

#End Region

#Region " ¤×¤×¤×¤×¤×¤ Private variables ¤×¤×¤×¤×¤×¤ "

    Private mstrErrorMessage As String

#End Region

#Region " ¤×¤×¤×¤×¤ Read only properties ¤×¤×¤×¤×¤ "

    Public ReadOnly Property ErrorMessage() As String
        Get
            Return mstrErrorMessage
        End Get
    End Property

#End Region


    Public Function GetActiveServers() As String()

        Dim intCount As Integer
        Dim intIndex As Integer
        Dim strServer As String()
        Dim Enumerator As SqlDataSourceEnumerator
        Dim Row As DataRow
        Dim DT As DataTable


        Try

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


        Catch ex As Exception
            mstrErrorMessage = ex.Message

        End Try

    End Function

    Public Function GetData(ByVal strTable As String, col As DataColumnCollection, ByVal strServer As String, ByVal strDatabase As String, ByVal strUser As String, ByVal strPassword As String) As DataTable

        Dim DA As SqlDataAdapter
        Dim DC As DataColumn
        Dim DT As DataTable
        Dim strFieldList As String = String.Empty


        Try

            For Each DC In col
                strFieldList &= DC.ColumnName & ", "
            Next
            If col.Count > 0 Then
                strFieldList = strFieldList.Substring(0, strFieldList.Length - 2)
            End If

            Using Conn As New SqlConnection
                Conn.ConnectionString = [String].Format(mconSQLConnectString, strServer, strDatabase, strUser, strPassword, strDatabase)
                Conn.Open()

                Using cmd As New SqlCommand("SELECT " & strFieldList & " FROM " & strTable, Conn)
                    DT = New DataTable(strTable)
                    DT.Load(cmd.ExecuteReader(CommandBehavior.KeyInfo))
                    'DA = New SqlDataAdapter("SELECT * FROM " & strTable, [String].Format(mconSQLConnectString, strServer, strDatabase, strUser, strPassword, strDatabase))
                    DA = New SqlDataAdapter(cmd)
                    DA.Fill(DT)
                End Using

            End Using

            Return DT


        Catch ex As Exception

            mstrErrorMessage = "Error retrieving data: " & ex.Message
            Return Nothing

        End Try

    End Function

    Public Function GetDatabases(ByVal strServer As String, ByVal strUser As String, ByVal strPassword As String) As String()

        Dim intCount As Integer
        Dim intIndex As Integer
        Dim strDatabase As String()
        Dim Row As DataRow
        Dim DT As DataTable


        Try

            Using Conn As New SqlConnection
                Conn.ConnectionString = [String].Format(mconSQLConnectString, strServer, "", strUser, strPassword, "")
                Conn.Open()
                DT = Conn.GetSchema(SqlClientMetaDataCollectionNames.Databases)
            End Using

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

            mstrErrorMessage = "Error retrieving databases: " & ex.Message
            Return Nothing

        End Try

    End Function

    Public Function GetFields(ByVal strTable As String, ByVal strServer As String, ByVal strDatabase As String, ByVal strUser As String, ByVal strPassword As String) As String()

        Dim strTableName As String()
        Dim strField As String()
        Dim r As Integer
        Dim Row As DataRow
        Dim FieldList As DataTable


        Using Conn As New SqlConnection
            Conn.ConnectionString = [String].Format(mconSQLConnectString, strServer, strDatabase, strUser, strPassword, strDatabase)
            Conn.Open()


            ReDim strTableName(3)
            strTableName(2) = strTable
            'Catalog, Schema, Table, Column
            FieldList = Conn.GetSchema("Columns", strTableName)
            ReDim strField(FieldList.Rows.Count - 1)

        End Using

        For Each Row In FieldList.Rows

            '0 = catalog; 1 = schema; 2 = Table Name; 3 = Column Name; 4 = Ordinal; 5 = Default; 6 = Nullable; 7 = Datatype; 8 = MaxLength;
            '   9 = Octet length; 10 = Precision; 11 = NP Radix; 12 = Scale; 13 = DataTime Precision; 14 = CSC; 15 = CSS; 16 = CSN;
            '   17 colation 
            strField(r) = Row(3).ToString
            r += 1
        Next

        Return strField

    End Function

    Public Function GetTable(ByVal strTable As String, ByVal strServer As String, ByVal strDatabase As String, ByVal strUser As String, ByVal strPassword As String) As DataTable

        Return GetTable(New String() {}, strTable, strServer, strDatabase, strUser, strPassword)

    End Function

    Public Function GetTable(ByVal strFields As String(), ByVal strTable As String, ByVal strServer As String, ByVal strDatabase As String, ByVal strUser As String, ByVal strPassword As String) As DataTable

        Dim DC As DataColumn
        Dim DT As DataTable
        Dim r As Integer
        Dim strSQL As String
        Dim strPKey As String = String.Empty


        Using Conn As New SqlConnection

            Conn.ConnectionString = [String].Format(mconSQLConnectString, strServer, strDatabase, strUser, strPassword, strDatabase)
            Conn.Open()

            DT = New DataTable
            If strFields.Length < 1 Then
                strSQL = "SELECT * FROM " & strTable
            Else
                strSQL = "SELECT * "
                For r = 0 To strFields.Length - 1
                    strSQL &= "[" & strFields(r) & "], "
                Next
                strSQL = strSQL.Substring(0, strSQL.Length - 2)
                strSQL = " FROM " & strTable
            End If

            Using cmd As SqlCommand = New SqlCommand(strSQL, Conn)
                DT.Load(cmd.ExecuteReader(CommandBehavior.KeyInfo))
                If DT.PrimaryKey.Length > 0 Then
                    strPKey = " ORDER BY "
                    For Each DC In DT.PrimaryKey
                        strPKey &= "[" & DC.ColumnName & "], "
                    Next
                    strPKey = strPKey.Substring(0, strPKey.Length - 2)
                End If
                cmd.CommandText = strSQL & strPKey
                DT.Load(cmd.ExecuteReader())
            End Using

        End Using

        Return DT

    End Function

    Public Function GetTables(ByVal strServer As String, ByVal strDatabase As String, ByVal strUser As String, ByVal strPassword As String) As String()

        Dim intIndex As Integer
        Dim strTables As String()
        Dim Row As DataRow
        Dim TableList As DataTable


        Try

            Using Conn As New SqlConnection

                Conn.ConnectionString = [String].Format(mconSQLConnectString, strServer, strDatabase, strUser, strPassword, strDatabase)
                Conn.Open()

                If Conn.ConnectionString <> vbNullString Then

                    TableList = Conn.GetSchema("Tables")

                    ReDim strTables(TableList.Rows.Count)

                    For Each Row In TableList.Rows
                        '0 = catalog; 1 = schema; 2 = Name; 3 = Type
                        If Row.Item(3).ToString = "BASE TABLE" Then
                            strTables(intIndex) = Row.Item(2).ToString
                            intIndex += 1
                            'Debug.Print(intIndex.ToString)
                        End If
                    Next

                End If

            End Using


            ReDim Preserve strTables(intIndex - 1)
            Return strTables

        Catch ex As Exception

            mstrErrorMessage = "Error retrieving tables: " & ex.Message
            Return Nothing

        End Try

    End Function

    Public Function GetViews() As List(Of String)

        Dim strSQL As String
        Dim lisViewNames As List(Of String) = Nothing
        Dim Reader As System.Data.SqlClient.SqlDataReader = Nothing
        Dim cmd As System.Data.SqlClient.SqlCommand


        Try

            Using Conn As New SqlConnection

                strSQL = "SELECT [name]" & _
                        " FROM SYS.VIEWS" & _
                        " ORDER BY [name]"
                cmd = New System.Data.SqlClient.SqlCommand(strSQL, Conn)
                cmd.CommandTimeout = 0
                Reader = cmd.ExecuteReader()
                If Reader.Read Then
                    lisViewNames = New List(Of String)

                    Do
                        lisViewNames.Add(Reader(0).ToString)
                    Loop While Reader.Read

                End If

                Return lisViewNames

            End Using


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

            Using Conn As New SqlConnection

                strSQL = "SELECT SV.name, definition " & _
                            "FROM sys.all_sql_modules A " & _
                                "LEFT JOIN sys.views SV ON a.object_id = SV.object_id " & _
                            "WHERE A.object_id IN (SELECT object_id FROM SYS.VIEWS) " & _
                            "ORDER BY SV.name"
                cmd = New System.Data.SqlClient.SqlCommand(strSQL, Conn)
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

            End Using


        Catch ex As Exception

            MsgBox("Error getting views: " & ex.Message)
            Return Nothing

        Finally

            Reader.Close()

        End Try

    End Function

    Public Function PrimaryKey(PKeys As DataColumn()) As String

        Dim DC As DataColumn
        Dim strPrimaryKey As String = String.Empty


        For Each DC In PKeys
            strPrimaryKey &= DC.ColumnName & ":"
        Next

        If strPrimaryKey.Length > 0 Then
            strPrimaryKey = strPrimaryKey.Substring(0, strPrimaryKey.Length - 1)
        End If

        Return strPrimaryKey

    End Function

    Public Function PrimaryKeyValue(PKeys As DataColumn(), DT As DataTable, DR As DataRow) As String

        Dim DC As DataColumn
        Dim r As Integer
        Dim strValue As String = String.Empty


        For Each DC In PKeys
            For r = 0 To DT.Columns.Count - 1
                If DC.ColumnName = DT.Columns(r).ColumnName Then
                    strValue &= DR.Item(r).ToString & ":"
                End If
            Next
        Next

        If strValue.Length > 0 Then
            strValue = strValue.Substring(0, strValue.Length - 1)
        End If

        Return strValue

    End Function

    Public Function PKeyWhere(PKeys As DataColumn(), DT As DataTable, DR As DataRow) As String

        Dim DC As DataColumn
        Dim r As Integer
        Dim strWhere As String = String.Empty


        For Each DC In PKeys
            For r = 0 To DT.Columns.Count - 1
                If DC.ColumnName = DT.Columns(r).ColumnName Then
                    strWhere &= DC.ColumnName & " = " & QuotesIfNeeded(DT.Columns(r).DataType, DR.Item(r)) & " AND "
                End If
            Next
        Next

        If strWhere.Length > 0 Then
            strWhere = strWhere.Substring(0, strWhere.Length - 5)
        End If

        Return strWhere

    End Function

    'Public Function PKeyWhere(PKeys As Dictionary(Of String, String)) As String

    '    Dim KeyValue As KeyValuePair(Of String, String)
    '    Dim PKeysKeyParsed As String()
    '    Dim PKeysValueParsed As String()
    '    Dim r As Integer
    '    Dim strWhere As String = String.Empty


    '    For Each KeyValue In PKeys
    '        PKeysKeyParsed = KeyValue.Key.Split(":")
    '        PKeysValueParsed = KeyValue.Value.Split(":")
    '        For r = 0 To PKeysKeyParsed.Length - 1
    '                    strWhere &= DC.ColumnName & " = " & QuotesIfNeeded(DT.Columns(r).DataType, DR.Item(r)) & " AND "
    '        Next
    '        For Each DC In PKeys
    '            For r = 0 To DT.Columns.Count
    '                If DC.ColumnName = DT.Columns(r).ColumnName Then
    '                    strWhere &= DC.ColumnName & " = " & QuotesIfNeeded(DT.Columns(r).DataType, DR.Item(r)) & " AND "
    '                End If
    '            Next
    '        Next

    '        If strWhere.Length > 0 Then
    '            strWhere = strWhere.Substring(0, strWhere.Length - 5)
    '        End If

    '        Return strWhere

    'End Function

End Module
