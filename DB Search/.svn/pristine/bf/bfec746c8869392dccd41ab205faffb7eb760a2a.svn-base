Module DataType

    Public Function DataTypeMappingSQLServer(ByVal strDBType As String) As String

        Select Case strDBType
            Case "bigint"               'Integer data from -2^63 through 2^63-1
                Return "long"
            Case "int"                  'Integer data from -2^31 through 2^31 - 1
                Return "integer"
            Case "smallint"             'Integer data from -2^15 through 2^15 - 1
                Return "short"
            Case "tinyint"              'Integer data from 0 through 255
                Return "byte"
            Case "bit"                  'Integer data with either a 1 or 0 value
                Return "boolean"
            Case "decimal"              'Fixed precision and scale numeric data from -10^38 +1 through 10^38 -1
                Return "single"
            Case "numeric"              'Fixed precision and scale numeric data from -10^38 +1 through 10^38 -1
                Return "single"
            Case "money"                'Monetary data values from -2^63 through 2^63 - 1
                Return "double"
            Case "smallmoney"           'Monetary data values from -214,748.3648 through +214,748.3647
                Return "single"
            Case "float"                'Floating precision number data from -1.79E + 308 through 1.79E + 308
                Return "double"
            Case "real"                 'Floating precision number data from -3.40E + 38 through 3.40E + 38
                Return "single"
            Case "datetime"             'Date and time data from January 1, 1753, through December 31, 9999, with an accuracy of 3.33 milliseconds
                Return "datetime"
            Case "smalldatetime"        'Date and time data from January 1, 1900, through June 6, 2079, with an accuracy of one minute
                Return "datetime"
            Case "char"                 'Fixed-length character data with a maximum length of 8,000 characters
                Return "string"
            Case "varchar"              'Variable-length data with a maximum of 8,000 characters
                Return "string"
            Case "text"                 'Variable-length data with a maximum length of 2^31 - 1 characters
                Return "string"
            Case "nchar"                'Fixed-length Unicode data with a maximum length of 4,000 characters
                Return "string"
            Case "nvarchar"             'Variable-length Unicode data with a maximum length of 4,000 characters
                Return "string"
            Case "ntext"                'Variable-length Unicode data with a maximum length of 2^30 - 1 characters
                Return "string"
            Case "binary"               'Fixed-length binary data with a maximum length of 8,000 bytes
                Return "byte()"
            Case "varbinary"            'Variable-length binary data with a maximum length of 8,000 bytes
                Return "byte()"
            Case "image"                'Variable-length binary data with a maximum length of 2^31 - 1 bytes
                Return "byte()"
            Case "cursor"               'A reference to a cursor
                Return "xxx"
            Case "sql_variant"          'A data type that stores values of various data types, except text, ntext, timestamp, and sql_variant
                Return "xxx"
            Case "table"                'A special data type used to store a result set for later processing
                Return "xxx"
            Case "timestamp"            'A database-wide unique number that gets updated every time a row gets updated
                Return "datetime"
            Case "uniqueidentifier"     'A globally unique identifier
                Return "string"
            Case Else
                Return "ERR"
        End Select

    End Function

    Public Function DataTypeMappingAccess(ByVal strDBType As String) As String

        'Access DataTypes
        'Text 	    Text or combinations of text and numbers, such as addresses. Also numbers that do not require calculations, such as phone numbers, part numbers, or postal codes. 	Up to 255 characters.
        'Microsoft Access only stores the characters entered in a field; it does not store space characters for unused positions in a Text field. To control the maximum number of characters that can be entered, set the FieldSize property.
        'Memo 	    Lengthy text and numbers, such as notes or descriptions. 	Up to 64,000 characters.
        'Number 	Numeric data to be used for mathematical calculations, except calculations involving money (use Currency type). Set the FieldSize property to define the specific Number type. 	1, 2, 4, or 8 bytes. 16 bytes for Replication ID (GUID) only.
        'Date/Time 	Dates and times. 	8 bytes
        'Currency 	Currency values. Use the Currency data type to prevent rounding off during calculations. Accurate to 15 digits to the left of the decimal point and 4 digits to the right. 	8 bytes
        'Auto-number 	Unique sequential (incrementing by 1) or random numbers automatically inserted when a record is added. 	4 bytes. 16 bytes for Replication ID (GUID) only.
        'Yes/No 	Fields that will contain only one of two values, such as Yes/No, True/False, On/Off. 	1 bit
        'OLE Objects 	Objects (such as Microsoft Word documents, Microsoft Excel spreadsheets, pictures, sounds, or other binary data), created in other programs using the OLE protocol that can be linked to or embedded in a Microsoft Access table. You must use a bound object frame in a form or report to display the OLE object. 	Up to 1 gigabyte (limited by disk space).
        'Hyperlink 	Field that will store hyperlinks. A hyperlink can be a UNC path or a URL. 	Up to 64,000 characters.
        'Lookup Wizard 	Creates a field that allows you to choose a value from another table or from a list of values using a combo box. Choosing this option in the data type list starts a wizard to define this for you. 	The same size as the primary key field that is also the Lookup field; typically 4 bytes.

        Return vbNullString

    End Function

    Public Function CheckDataTypeIsString(ByVal strDataType As String) As Boolean

        Select Case strDataType
            Case "char", "varchar", "text", "nchar", "nvarchar", "ntext", "uniqueidentifier"
                Return True
            Case Else
                Return False
        End Select

    End Function

    Public Function CheckDataTypeUseQuotes(ByVal strDataType As String) As Boolean
        'change CheckDataTypeVsQuotes to CheckDataTypeUseQuotes
        Select Case strDataType
            Case "datetime", "smalldatetime", "char", "varchar", "text", "nchar", "nvarchar", "ntext", "uniqueidentifier"
                Return True
            Case Else
                Return False
        End Select

    End Function

    Public Function CheckDataTypeforConvert(ByVal strDataType As String) As Boolean

        Select Case strDataType
            Case "boolean", "byte", "short", "integer", "long", "single", "double", "decimal", "datetime"
                Return True
            Case Else
                Return False
        End Select

    End Function

    Public Function GetPrefixFromDataType(ByVal strType As String) As String

        If strType.Length > 2 Then
            'Arrays will use the same prefix as non-arrays
            If Right(strType, 2) = "()" Then
                strType = Truncate(strType, 2)
            End If
            Select Case LCase(strType)
                Case "constant"
                    Return "con"
                Case "boolean" '4 bytes; True or False
                    Return "b"
                Case "byte" '1 byte; 0 to 255 (unsigned)
                    Return "byt"
                Case "char" '2 bytes; 0 to 65535 (unsigned)
                    Return "chr"
                Case "short", "int16" '2 bytes; -32,768 to 32,767 
                    Return "sht"
                Case "integer", "int32" '4 bytes; -2,147,483,648 to +2,147,483,647
                    Return "int"
                Case "long", "int64" '8 bytes; -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 
                    Return "lng"
                Case "single" '4 bytes; -3.402823E38 to -1.401298E-45 for negative values; 1.401298E-45 to 3.402823E38 for positive values 
                    Return "sng"
                Case "double" '8 bytes; -1.79769313486231E308 to -4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values 
                    Return "dbl"
                Case "decimal" '16 bytes; +/-79,228,162,514,264,337,593,543,950,335 with no decimal point; +/-7.9228162514264337593543950335 with 28 places to the right of the decimal
                    Return "dec"
                Case "string" '10 bytes + (2 * string length); 0 to approximately two billion Unicode characters 
                    Return "str"
                Case "datetime" '8 bytes
                    Return "dtm"
                Case Else
                    'Unhandled type
                    Return ""
            End Select
        Else
            'Error return type
            Return "xxx"

        End If

    End Function

    Public Function GetDatatypeFromPrefix(ByVal strVar As String) As String

        If strVar.Length >= 3 Then
            Select Case LCase(Left(strVar, 3))
                Case "con"
                    Return "constant"
                Case "byt" '1 byte; 0 to 255 (unsigned)
                    Return "byte"
                Case "chr" '2 bytes; 0 to 65535 (unsigned)
                    Return "char"
                Case "sht" '2 bytes; -32,768 to 32,767 
                    Return "short"
                Case "int" '4 bytes; -2,147,483,648 to +2,147,483,647
                    Return "integer"
                Case "lng" '8 bytes; -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 
                    Return "long"
                Case "sng" '4 bytes; -3.402823E38 to -1.401298E-45 for negative values; 1.401298E-45 to 3.402823E38 for positive values 
                    Return "single"
                Case "dbl" '8 bytes; -1.79769313486231E308 to -4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values 
                    Return "double"
                Case "dec" '16 bytes; +/-79,228,162,514,264,337,593,543,950,335 with no decimal point; +/-7.9228162514264337593543950335 with 28 places to the right of the decimal
                    Return "decimal"
                Case "str" '10 bytes + (2 * string length); 0 to approximately two billion Unicode characters 
                    Return "string"
                Case "dtm" '8 bytes
                    Return "datetime"
                Case Else
                    'Check for boolean
                    If Left(strVar, 1) = "b" Then
                        'boolean
                        Return "boolean"
                    Else
                        'Unhandled type
                        Return "ERR"
                    End If
            End Select

        Else
            'Error return type
            Return "xxx"

        End If

    End Function


End Module
