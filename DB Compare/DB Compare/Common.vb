Module Common

    Public Const mconSeparator As String = ":"


    Public Sub ParseConcat(ByVal strFull As String, ByRef strFirst As String, ByRef strLast As String)

        Dim intLocation As Integer


        intLocation = InStr(strFull, mconSeparator)
        If intLocation > 1 And intLocation < strFull.Length Then
            strFirst = Left(strFull, intLocation - 1)
            strLast = Mid(strFull, intLocation + 1)
        End If

    End Sub

    Public Function QuotesIfNeeded(DataType As Type, Value As String) As String

        Select Case True
            Case DataType Is GetType(String)
                Return "'" & Value & "'"
            Case Else
                Return Value
        End Select

    End Function

    Public Function Truncate(ByVal strString As String, ByVal intTruncateLength As Integer) As String

        Dim intStringLength As Integer
        Dim strReturnedString As String


        If strString <> vbNullString Then
            strReturnedString = strString
            intStringLength = strString.Length

            If intStringLength > intTruncateLength Then
                strReturnedString = Left(strString, intStringLength - intTruncateLength)
            End If

            Return strReturnedString

        Else

            Return Nothing

        End If

    End Function


End Module
