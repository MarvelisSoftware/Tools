Public Module Structures

    Public Structure ItemInformation
        Public Definition As String
        Public Name As String

        Public Overrides Function ToString() As String
            Return Name
        End Function
    End Structure

    Public Structure PrimaryKey
        Public Datatype As Type
        Public Name As String
        Public Value As String

        Public Overrides Function ToString() As String
            Return Name
        End Function
    End Structure

End Module
