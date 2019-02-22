Public Class Types

    Public Class PrimaryKeys
        Inherits Dictionary(Of PrimaryKey, String)


        Public Overloads Function Add(DCs As DataColumn())

            Dim DC As DataColumn
            Dim PK As PrimaryKey = Nothing
            Dim strPrimaryKey As String = String.Empty


            For Each DC In DCs
                strPrimaryKey &= DC.ColumnName & ":"
            Next

            Return PK

        End Function

        Public Overrides Function ToString() As String

            Dim KeyValue As KeyValuePair(Of PrimaryKey, String)
            Dim PKey As PrimaryKey
            Dim strWhere As String = String.Empty


            For Each KeyValue In Me
                PKey = CType(KeyValue.Key, PrimaryKey)
                strWhere &= PKey.Name & " = " & QuotesIfNeeded(PKey.Datatype, PKey.Value) & " AND "
            Next

            If strWhere.Length > 0 Then
                strWhere = strWhere.Substring(0, strWhere.Length - 5)
            End If

            Return strWhere

        End Function

    End Class


End Class
