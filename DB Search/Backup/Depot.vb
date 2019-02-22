Public Class Depot


    Public Enum eDBDataType
        Unknown
        Bit
        Tinyint
        Smallint
        Int
        Bigint
        Smallmoney
        [Decimal]
        Numeric
        Money
        Float
        Real
        Timestamp
        Smalldatetime
        Datetime
        [Char]
        Text
        Varchar
        Nchar
        Ntext
        Nvarchar
        Uniqueidentifier
        Sql_variant
        Binary
        Varbinary
        Image
        Cursor
        Table
    End Enum

    Public Enum eFormType
        Windows = 1
        IMS
    End Enum

    Public Enum eProjectSource
        [New] = 1
        Load
    End Enum

    Public Enum eProjectType
        EXE = 1
        DLL
    End Enum

    Public Enum eReadType
        DataAdaptor = 1
        Recordset
    End Enum

    Public Enum eType
        NoneSet
        Parent
        Child
    End Enum


    Private meDatabaseType As eDatabaseType
    Private meFormType As eFormType
    Private meProjectSource As eProjectSource
    Private meProjectType As eProjectType
    Private meReadType As eReadType
    Private mbCreateClasses As Boolean
    Private mbCreateCollection As Boolean
    Private mbCreateChangeTables As Boolean
    Private mbCreateDataAccess As Boolean
    Private mbCreateFindType As Boolean
    Private mbCreateForm As Boolean
    Private mbCreateInterface As Boolean
    Private mbCreateTopLeveLCollection As Boolean
    Private mbMainDA As Boolean
    Private mbStoredProcedure As Boolean
    Private mbRemoveTablePrefix As Boolean
    Private mstrProjectName As String
    Private mstrProjectPath As String
    Private mstrFormFileName As String
    Private mstrFormName As String
    Private mstrSPName As String
    Private mstrTopLevelName As String
    Private mstrTablePrefix As String
    Private mFormCode As FormData
    Private mcolTables As Tables
    Private mcolLinks As Links


#Region " ¤×¤×¤×¤×¤×¤×¤ Classes ¤×¤×¤×¤×¤×¤×¤ "

    Public Class FormData

        Public strClearControls As String()
        Public strEnableControls As String()
        Public strDisplayControlsFromFields As String()
        Public strStoreControlsToFields As String()
        Private mstrVerifyControls As String

        Public CreateNewObjectMethod As String
        Public DataEntryFormName As String
        Public LoadObjectMethod As String
        Public LookupControl As String
        Public TopLevelObject As String

        Public Sub ClearCodeArrays()

            Erase strClearControls
            Erase strEnableControls
            Erase strDisplayControlsFromFields
            Erase strStoreControlsToFields

        End Sub

        Public Property ClearControls() As String
            Get
                Return strClearControls.GetUpperBound(0).ToString
            End Get
            Set(ByVal value As String)

                Dim intIndex As Integer


                If strClearControls IsNot Nothing Then
                    intIndex = strClearControls.GetUpperBound(0) + 1
                End If
                ReDim Preserve strClearControls(intIndex)
                strClearControls(intIndex) = value

            End Set
        End Property

        Public Property ClearControls(ByVal intIndex As Integer) As String
            Get
                Return strClearControls(intIndex)
            End Get
            Set(ByVal value As String)

                If intIndex >= 0 And intIndex <= strClearControls.GetUpperBound(0) Then
                    strClearControls(intIndex) = value
                End If

            End Set
        End Property

        Public Property EnableControls() As String
            Get
                Return strEnableControls.GetUpperBound(0).ToString
            End Get
            Set(ByVal value As String)

                Dim intIndex As Integer


                If strEnableControls IsNot Nothing Then
                    intIndex = strEnableControls.GetUpperBound(0) + 1
                End If
                ReDim Preserve strEnableControls(intIndex)
                strEnableControls(intIndex) = value

            End Set
        End Property

        Public Property EnableControls(ByVal intIndex As Integer) As String
            Get
                Return strEnableControls(intIndex)
            End Get
            Set(ByVal value As String)

                If intIndex >= 0 And intIndex <= strEnableControls.GetUpperBound(0) Then
                    strEnableControls(intIndex) = value
                End If

            End Set
        End Property

        Public Property DisplayControlsFromFields() As String
            Get
                Return strDisplayControlsFromFields.GetUpperBound(0).ToString
            End Get
            Set(ByVal value As String)

                Dim intIndex As Integer


                If strDisplayControlsFromFields IsNot Nothing Then
                    intIndex = strDisplayControlsFromFields.GetUpperBound(0) + 1
                End If
                ReDim Preserve strDisplayControlsFromFields(intIndex)
                strDisplayControlsFromFields(intIndex) = value

            End Set
        End Property

        Public Property DisplayControlsFromFields(ByVal intIndex As Integer) As String
            Get
                Return strDisplayControlsFromFields(intIndex)
            End Get
            Set(ByVal value As String)

                If intIndex >= 0 And intIndex <= strDisplayControlsFromFields.GetUpperBound(0) Then
                    strDisplayControlsFromFields(intIndex) = value
                End If

            End Set
        End Property

        Public Property StoreControlsToFields() As String
            Get
                Return strStoreControlsToFields.GetUpperBound(0).ToString
            End Get
            Set(ByVal value As String)

                Dim intIndex As Integer


                If strStoreControlsToFields IsNot Nothing Then
                    intIndex = strStoreControlsToFields.GetUpperBound(0) + 1
                End If
                ReDim Preserve strStoreControlsToFields(intIndex)
                strStoreControlsToFields(intIndex) = value

            End Set
        End Property

        Public Property StoreControlsToFields(ByVal intIndex As Integer) As String
            Get
                Return strStoreControlsToFields(intIndex)
            End Get
            Set(ByVal value As String)

                If intIndex >= 0 And intIndex <= strStoreControlsToFields.GetUpperBound(0) Then
                    strStoreControlsToFields(intIndex) = value
                End If

            End Set
        End Property

        Public Property VerifyControls() As String
            Get
                Return mstrVerifyControls
            End Get
            Set(ByVal value As String)

                mstrVerifyControls = value

            End Set
        End Property

    End Class


    Public Class Links

        Inherits CollectionBase


        Public Class Link

            Private meType As eType
            Private mstrName As String       'Name is Table (owner of the link) & mconSeparator & Linked table (can be a parent or child)
            Private mstrLocalTableName As String
            Private mstrForiegnTableName As String
            Private mcolLocalKeys As Depot.Tables.Table.Fields
            Private mcolForiegnKeys As Depot.Tables.Table.Fields

            Property Type() As eType
                Get
                    Return meType
                End Get
                Set(ByVal value As eType)
                    meType = value
                End Set
            End Property

            Property Name() As String
                Get
                    Return mstrName
                End Get
                Set(ByVal value As String)
                    mstrName = value
                End Set
            End Property

            Property LocalTableName() As String
                Get
                    Return mstrLocalTableName
                End Get
                Set(ByVal value As String)
                    mstrLocalTableName = value
                End Set
            End Property

            Property ForiegnTableName() As String
                Get
                    Return mstrForiegnTableName
                End Get
                Set(ByVal value As String)
                    mstrForiegnTableName = value
                End Set
            End Property

            Property LocalKeys() As Depot.Tables.Table.Fields
                Get
                    Return mcolLocalKeys
                End Get
                Set(ByVal value As Depot.Tables.Table.Fields)
                    mcolLocalKeys = value
                End Set
            End Property

            Property ForiegnKeys() As Depot.Tables.Table.Fields
                Get
                    Return mcolForiegnKeys
                End Get
                Set(ByVal value As Depot.Tables.Table.Fields)
                    mcolForiegnKeys = value
                End Set
            End Property


            Public Sub New()
                mcolLocalKeys = New Depot.Tables.Table.Fields
                mcolForiegnKeys = New Depot.Tables.Table.Fields
            End Sub

            Public Sub AddFields(ByVal LocalField As Depot.Tables.Table.Fields.Field, ByVal ForiegnField As Depot.Tables.Table.Fields.Field)

                mcolLocalKeys.Add(LocalField)
                mcolForiegnKeys.Add(ForiegnField)

            End Sub

            Public Sub DeleteLink(ByVal strLocal As String, ByVal strForiegn As String)

                Dim intIndex As Integer


                intIndex = CheckLink(strLocal, strForiegn)
                If intIndex > 0 Then
                    mcolLocalKeys.RemoveAt(intIndex)
                    mcolForiegnKeys.RemoveAt(intIndex)
                End If

            End Sub

            Public Function CheckLink(ByVal strLocal As String, ByVal strForiegn As String) As Integer

                Dim bFound As Boolean
                Dim intIndex As Integer
                Dim Field As Depot.Tables.Table.Fields.Field


                intIndex = -1
                For Each Field In LocalKeys
                    intIndex += 1
                    If Field.Name = strLocal AndAlso mcolForiegnKeys(intIndex).Name = strForiegn Then
                        If mcolForiegnKeys(intIndex).Name = strForiegn Then
                            bFound = True
                            Exit For
                        End If
                    End If
                Next

                If bFound Then
                    Return intIndex
                Else
                    Return -1
                End If

            End Function

        End Class

        'Collection
        Public Shadows ReadOnly Property Count() As Integer
            Get
                Return MyBase.List.Count
            End Get
        End Property

        Public ReadOnly Property GetChildren(ByVal strTableName As String) As String()
            Get

                Dim intCounter As Integer
                Dim strChildren As String()
                Dim Link As Link


                ReDim strChildren(9)
                For Each Link In List
                    If Link.ForiegnTableName = strTableName And Link.Type = eType.Child Then
                        strChildren(intCounter) = Link.LocalTableName
                        intCounter += 1
                        If intCounter / 10 = Int(intCounter / 10) Then
                            ReDim Preserve strChildren(intCounter + 9)
                        End If
                    End If
                Next

                If intCounter > 0 Then
                    ReDim Preserve strChildren(intCounter - 1)
                    Return strChildren
                Else
                    Return Nothing
                End If

            End Get
        End Property

        Public ReadOnly Property GetParent(ByVal strTableName As String) As String
            Get

                Dim strParent As String = vbNullString
                Dim Link As Link


                For Each Link In List
                    If Link.ForiegnTableName = strTableName And Link.Type = eType.Parent Then
                        strParent = Link.LocalTableName
                    End If
                Next

                Return strParent

            End Get
        End Property

        Public ReadOnly Property GetChildLinks(ByVal strTableName As String, ByVal strParentTableName As String) As String(,)
            Get

                Dim r As Integer
                Dim strLinks As String(,) = {}
                Dim Link As Link


                For Each Link In List
                    'If strParentTableName = vbNullString Then
                    '    If Link.ForiegnTableName = strTableName AndAlso Link.Type = eType.Child Then
                    '        ReDim strLinks(1, Link.LocalKeys.Count - 1)
                    '        For r = 0 To Link.LocalKeys.Count - 1
                    '            strLinks(0, r) = Link.LocalKeys(r).Name
                    '            strLinks(1, r) = Link.ForiegnKeys(r).Name
                    '        Next
                    '    End If
                    'Else
                    If Link.ForiegnTableName = strTableName AndAlso Link.Type = eType.Child _
                    AndAlso Link.LocalTableName = strParentTableName Then
                        ReDim strLinks(1, Link.LocalKeys.Count - 1)
                        For r = 0 To Link.LocalKeys.Count - 1
                            strLinks(0, r) = Link.LocalKeys(r).Name
                            strLinks(1, r) = Link.ForiegnKeys(r).Name
                        Next
                    End If
                    'End If
                Next

                Return strLinks

            End Get
        End Property

        Public ReadOnly Property GetForiegnChildKeys(ByVal strTableName As String, ByVal strChildTableName As String) As String()
            Get

                Dim r As Integer
                Dim strKeys As String() = {}
                Dim Link As Link


                For Each Link In List
                    If Link.LocalTableName = strTableName AndAlso Link.Type = eType.Child _
                        AndAlso Link.ForiegnTableName = strChildTableName Then
                        ReDim strKeys(Link.ForiegnKeys.Count - 1)
                        For r = 0 To Link.ForiegnKeys.Count - 1
                            strKeys(r) = Link.ForiegnKeys(r).Name
                        Next
                    End If
                Next

                Return strKeys

            End Get
        End Property

        Public ReadOnly Property GetForiegnParentKeys(ByVal strTableName As String, ByVal strParentTableName As String) As String()
            Get

                Dim r As Integer
                Dim strKeys As String() = {}
                Dim Link As Link


                For Each Link In List
                    If Link.LocalTableName = strTableName AndAlso Link.Type = eType.Parent _
                        AndAlso Link.ForiegnTableName = strParentTableName Then
                        ReDim strKeys(Link.ForiegnKeys.Count - 1)
                        For r = 0 To Link.ForiegnKeys.Count - 1
                            strKeys(r) = Link.ForiegnKeys(r).Name
                        Next
                    End If
                Next

                Return strKeys

            End Get
        End Property

        Public ReadOnly Property GetLocalChildKeys(ByVal strTableName As String, ByVal strChildTableName As String) As String()
            Get

                Dim r As Integer
                Dim strKeys As String() = {}
                Dim Link As Link


                For Each Link In List
                    If Link.LocalTableName = strTableName AndAlso Link.Type = eType.Child _
                        AndAlso Link.ForiegnTableName = strChildTableName Then
                        ReDim strKeys(Link.LocalKeys.Count - 1)
                        For r = 0 To Link.LocalKeys.Count - 1
                            strKeys(r) = Link.LocalKeys(r).Name
                        Next
                    End If
                Next

                Return strKeys

            End Get
        End Property

        Public ReadOnly Property GetLocalParentKeys(ByVal strTableName As String, ByVal strParentTableName As String) As String()
            Get

                Dim r As Integer
                Dim strKeys As String() = {}
                Dim Link As Link


                For Each Link In List
                    If Link.LocalTableName = strTableName AndAlso Link.Type = eType.Parent _
                        AndAlso Link.ForiegnTableName = strParentTableName Then
                        ReDim strKeys(Link.LocalKeys.Count - 1)
                        For r = 0 To Link.LocalKeys.Count - 1
                            strKeys(r) = Link.LocalKeys(r).Name
                        Next
                    End If
                Next

                Return strKeys

            End Get
        End Property

        Public ReadOnly Property GetParentLinks(ByVal strTableName As String, Optional ByVal strChildTableName As String = vbNullString) As String(,)
            Get

                Dim r As Integer
                Dim strLinks As String(,) = {}
                Dim Link As Link


                For Each Link In List
                    If strChildTableName = vbNullString Then
                        If Link.LocalTableName = strTableName AndAlso Link.Type = eType.Parent Then
                            ReDim strLinks(1, Link.LocalKeys.Count - 1)
                            For r = 0 To Link.LocalKeys.Count - 1
                                strLinks(0, r) = Link.LocalKeys(r).Name
                                strLinks(1, r) = Link.ForiegnKeys(r).Name
                            Next
                        End If
                    Else
                        If Link.LocalTableName = strTableName AndAlso Link.Type = eType.Parent _
                        AndAlso Link.ForiegnTableName = strChildTableName Then
                            ReDim strLinks(1, Link.LocalKeys.Count - 1)
                            For r = 0 To Link.LocalKeys.Count - 1
                                strLinks(0, r) = Link.LocalKeys(r).Name
                                strLinks(1, r) = Link.ForiegnKeys(r).Name
                            Next
                        End If
                    End If
                Next

                Return strLinks

            End Get
        End Property

        Public ReadOnly Property GetRelatedLink(ByVal strLinkName As String) As Link
            Get
                Dim bFound As Boolean
                Dim intLocation As Integer
                Dim strFindLocal As String
                Dim strFindForiegn As String
                Dim Link As Link = Nothing


                intLocation = InStr(strLinkName, mconSeparator)
                If intLocation > 1 And intLocation < strLinkName.Length Then
                    strFindLocal = Left(strLinkName, intLocation - 1)
                    strFindForiegn = Mid(strLinkName, intLocation + 1)
                    For Each Link In List
                        With Link
                            If .ForiegnTableName = strFindForiegn And .LocalTableName = strFindLocal Then
                                bFound = True
                                Exit For
                            End If
                        End With
                    Next
                End If

                If bFound Then
                    Return Link
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public ReadOnly Property IsChildLink(ByVal strChildTableName As String, ByVal strTableName As String, ByVal strFieldName As String) As String
            Get

                Dim bFound As Boolean
                Dim r As Integer
                Dim strLinkedField As String = vbNullString
                Dim Link As Link


                For Each Link In List
                    If Link.ForiegnTableName = strTableName AndAlso Link.Type = eType.Child AndAlso _
                    Link.LocalTableName = strChildTableName Then
                        For r = 0 To Link.ForiegnKeys.Count - 1
                            If Link.ForiegnKeys(r).Name = strFieldName Then
                                strLinkedField = Link.LocalKeys(r).Name
                                bFound = True
                                Exit For
                            End If
                        Next
                    End If
                Next

                If bFound Then
                    Return strLinkedField
                Else
                    Return Nothing
                End If

            End Get
        End Property

        Public ReadOnly Property IsParentLink(ByVal strParentTableName As String, ByVal strTableName As String, ByVal strFieldName As String) As String
            Get

                Dim bFound As Boolean
                Dim r As Integer
                Dim strLinkedField As String = vbNullString
                Dim Link As Link


                For Each Link In List
                    If Link.ForiegnTableName = strTableName AndAlso Link.Type = eType.Parent AndAlso _
                    Link.LocalTableName = strParentTableName Then
                        For r = 0 To Link.ForiegnKeys.Count - 1
                            If Link.ForiegnKeys(r).Name = strFieldName Then
                                strLinkedField = Link.LocalKeys(r).Name
                                bFound = True
                                Exit For
                            End If
                        Next
                    End If
                Next

                If bFound Then
                    Return strLinkedField
                Else
                    Return Nothing
                End If

            End Get
        End Property

        Default Public Property Item(ByVal index As Integer) As Link
            Get
                Return CType(MyBase.List.Item(index), Link)
            End Get
            Set(ByVal value As Link)
                MyBase.List.Item(index) = value
            End Set

        End Property

        Default Public ReadOnly Property Item(ByVal index As String, Optional ByVal bIgnoreCase As Boolean = False) As Link
            Get

                Dim bFound As Boolean
                Dim Link As Link


                If bIgnoreCase Then
                    index = UCase(index)
                End If
                Link = Nothing
                For Each Link In MyBase.List
                    If bIgnoreCase Then
                        If UCase(Link.Name) = index Then
                            bFound = True
                            Exit For
                        End If
                    Else
                        If Link.Name = index Then
                            bFound = True
                            Exit For
                        End If
                    End If
                Next
                If bFound Then
                    Return Link
                Else
                    Return Nothing
                End If
            End Get

        End Property

        Public Overrides Function ToString() As String

            Return "Collection of links, count = " & MyBase.List.Count.ToString

        End Function

        Public Overloads Function Add() As Link

            Dim Link As New Link

            MyBase.List.Add(Link)
            Return Link

        End Function

        Public Overloads Sub Add(ByVal Link As Link)

            MyBase.List.Add(Link)

        End Sub

        Public Function Contains(ByVal Link As Link) As Boolean

            Return List.Contains(Link)

        End Function

        Public Sub CopyTo(ByVal Array() As Link, ByVal intIndex As Integer)

            List.CopyTo(Array, intIndex)

        End Sub

        Public Function HasLinkForTable(ByVal strLocalTable As String, ByVal strForiegnTable As String) As Link

            Dim bFound As Boolean
            Dim Link As Link = Nothing


            For Each Link In List
                If Link.LocalTableName = strLocalTable And Link.ForiegnTableName = strForiegnTable Then
                    bFound = True
                    Exit For
                End If
            Next

            If bFound Then
                Return Link
            Else
                Return Nothing
            End If

        End Function

        Public Sub Insert(ByVal intIndex As Integer, ByVal Link As Link)

            List.Insert(intIndex, Link)

        End Sub

        Public Function IndexOf(ByVal Link As Link) As Integer

            Return List.IndexOf(Link)

        End Function

        Public Shadows Sub Remove(ByVal Link As Link)

            MyBase.List.Remove(Link)

        End Sub

    End Class

    Public Class Tables

        Inherits CollectionBase


        Public Class Table

            Private mbTopParent As Boolean
            Private mbCreatedConstants As Boolean
            Private mintCodeIndex As Integer
            Private mintNestingLevel As Integer
            'Private mintNestingLevelActual As Integer
            Private mstrClassName As String
            Private mstrClassNameVar As String
            Private mstrCollectionName As String
            Private mstrCollectionNameVar As String
            Private mstrCollectionProperty As String
            Private mstrFieldConstantList As String
            Private mstrGridDisplayMethod As String
            Private mstrGridName As String
            Private mstrGridFillMethod As String
            Private mintIndexColumn As Integer
            Private mstrIndexVar As String
            Private mstrLoadedVar As String
            Private mstrLookupDimVars As String
            Private mstrLookupList As String
            Private mstrLookupModularVars As String
            Private mstrLookupVars As String
            Private mstrNiceName As String
            Private mstrNewFieldList As String
            Private mstrObjectAccessClass As String
            Private mstrObjectAccessCollection As String
            Private mstrParent As String
            Private mstrParentLink As String
            Private mstrLinkToParent As String
            Private mstrSelectList As String
            Private mstrSelectListAsParameters As String
            Private mstrTableName As String
            Private mstrTypeDefClass As String
            Private mstrTypeDefCollection As String
            Private mstrChildren As String()
            Private mstrChildrenLink As String()
            Private mstrLinkToChildren As String()
            Private mstrLookup As String()
            Private mFieldsCollection As Fields

            Public Class Fields

                Inherits CollectionBase

                Private mbHasGUID As Boolean
                Private mstrKeyControl As String
                Private mstrFirstNonKeyControl As String

                Public Class Field

                    'Private intColumnLength As Integer

                    Public ControlName As String
                    Public ControlProperty As String
                    Public ControlType As String
                    Public DataLength As Integer
                    Public DBDataType As String
                    Public FKey As Boolean
                    Public FKeyParentField As String
                    Public Name As String
                    Public Hidden As Boolean                'Used in LOHCO's change tracking
                    Public Ordinal As Integer
                    Public PKey As Boolean
                    Public [ReadOnly] As Boolean            'Used in LOHCO's change tracking
                    Public Required As Boolean
                    Public UnNeededOnEdit As Boolean        'Used in LOHCO's change tracking for fields not needed to be updated on a change
                    Public UseAsLookup As Boolean
                    Public UseQuotes As Boolean
                    Public VBDataType As String

                    Public ReadOnly Property CheckDBDataType() As eDBDataType
                        Get

                            Select Case DBDataType
                                Case "bit"
                                    Return eDBDataType.Bit
                                Case "tinyint"
                                    Return eDBDataType.Tinyint
                                Case "smallint"
                                    Return eDBDataType.Smallint
                                Case "int"
                                    Return eDBDataType.Int
                                Case "bigint"
                                    Return eDBDataType.Bigint
                                Case "smallmoney"
                                    Return eDBDataType.Smallmoney
                                Case "decimal"
                                    Return eDBDataType.Decimal
                                Case "numeric"
                                    Return eDBDataType.Numeric
                                Case "money"
                                    Return eDBDataType.Money
                                Case "float"
                                    Return eDBDataType.Float
                                Case "real"
                                    Return eDBDataType.Real
                                Case "timestamp"
                                    Return eDBDataType.Timestamp
                                Case "smalldatetime"
                                    Return eDBDataType.Smalldatetime
                                Case "datetime"
                                    Return eDBDataType.Datetime
                                Case "char"
                                    Return eDBDataType.Char
                                Case "text"
                                    Return eDBDataType.Text
                                Case "varchar"
                                    Return eDBDataType.Varchar
                                Case "nchar"
                                    Return eDBDataType.Nchar
                                Case "ntext"
                                    Return eDBDataType.Ntext
                                Case "nvarchar"
                                    Return eDBDataType.Nvarchar
                                Case "uniqueidentifier"
                                    Return eDBDataType.Uniqueidentifier
                                Case "binary"
                                    Return eDBDataType.Binary
                                Case "varbinary"
                                    Return eDBDataType.Varbinary
                                Case "image"
                                    Return eDBDataType.Image
                                Case "sql_variant"
                                    Return eDBDataType.Sql_variant
                                Case "cursor"
                                    Return eDBDataType.Cursor
                                Case "table"
                                    Return eDBDataType.Table
                                Case Else
                                    Return eDBDataType.Unknown
                            End Select

                        End Get
                    End Property

                    Public ReadOnly Property ColumnLength() As Integer
                        Get
                            If DataLength > 1000 Then
                                Return 500
                            ElseIf DataLength > 0 Then
                                Return DataLength
                            Else
                                Return 50
                            End If
                        End Get
                    End Property

                    Public ReadOnly Property GetGridDataType() As String
                        Get

                            Select Case VBDataType
                                Case Else
                                    Return VBDataType
                            End Select

                        End Get
                    End Property

                    Public ReadOnly Property IsBoolean() As Boolean
                        Get
                            Select Case DBDataType
                                Case "bit"
                                    Return True
                                Case Else
                                    Return False
                            End Select
                        End Get
                    End Property

                    Public ReadOnly Property IsDate() As Boolean
                        Get
                            Select Case DBDataType
                                Case "datetime", "smalldatetime"
                                    Return True
                                Case Else
                                    Return False
                            End Select
                        End Get
                    End Property

                    Public ReadOnly Property IsDateTime() As Boolean
                        Get
                            Select Case DBDataType
                                Case "datetime", "smalldatetime", "timestamp"
                                    Return True
                                Case Else
                                    Return False
                            End Select
                        End Get
                    End Property

                    Public ReadOnly Property IsNumeric() As Boolean
                        Get
                            Select Case DBDataType
                                Case "decimal", "numeric", "money", "smallmoney", "float", "real", _
                                "bigint", "int", "smallint", "tinyint"
                                    Return True
                                Case Else
                                    Return False
                            End Select
                        End Get
                    End Property

                    Public ReadOnly Property IsString() As Boolean
                        Get
                            Select Case DBDataType
                                Case "char", "varchar", "text", "nchar", "nvarchar", "ntext", "uniqueidentifier"
                                    Return True
                                Case Else
                                    Return False
                            End Select
                        End Get
                    End Property

                    Public ReadOnly Property RequiresLength() As Boolean
                        Get
                            Select Case DBDataType
                                Case "decimal", "numeric", "money", "smallmoney", "float", "real", "char", "varchar", _
                                    "text", "nchar", "nvarchar", "ntext", "binary", "varbinary", "image"
                                    Return True
                                Case "bigint", "int", "smallint", "tinyint", "bit", "datetime", "smalldatetime", _
                                    "uniqueidentifier", "timestamp", "uniqueidentifier"
                                    Return False
                                Case "cursor", "sql_variant", "table"
                            End Select
                        End Get
                    End Property

                    Public ReadOnly Property UseToString() As Boolean
                        Get
                            Select Case DBDataType
                                Case "decimal", "numeric", "money", "smallmoney", "float", "real", "bigint", "int", "smallint", "tinyint"
                                    Return True
                                Case Else
                                    Return False
                            End Select
                        End Get
                    End Property

                End Class

                Public Shadows ReadOnly Property Count() As Integer
                    Get
                        Return MyBase.List.Count
                    End Get
                End Property

                Public ReadOnly Property GetOrderedCollection() As Fields
                    Get
                        Dim intIndex As Integer
                        Dim Field As New Field
                        Dim colFields As New Fields

                        For r = 0 To Me.Count - 1
                            For Each Field In MyBase.List
                                If Field.Ordinal = r Then
                                    colFields.Add(Field)
                                    intIndex += 1
                                    Exit For
                                End If
                            Next
                        Next
                        Return colFields
                    End Get
                End Property

                Public Property HasGUID() As Boolean
                    Get
                        Return mbHasGUID
                    End Get
                    Set(ByVal value As Boolean)
                        mbHasGUID = value
                    End Set
                End Property

                Public Property FirstNonKeyControl() As String
                    Get
                        Return mstrFirstNonKeyControl
                    End Get
                    Set(ByVal value As String)
                        mstrFirstNonKeyControl = value
                    End Set
                End Property

                Public Property KeyControl() As String
                    Get
                        Return mstrKeyControl
                    End Get
                    Set(ByVal value As String)
                        mstrKeyControl = value
                    End Set
                End Property

                Default Public Property Item(ByVal index As Integer) As Field
                    Get
                        Return CType(MyBase.List.Item(index), Field)
                    End Get
                    Set(ByVal value As Field)
                        MyBase.List.Item(index) = value
                    End Set

                End Property

                Default Public ReadOnly Property Item(ByVal index As String, Optional ByVal bIgnoreCase As Boolean = False) As Field
                    Get

                        Dim bFound As Boolean
                        Dim Field As Field


                        If bIgnoreCase Then
                            index = UCase(index)
                        End If
                        Field = Nothing
                        For Each Field In MyBase.List
                            If bIgnoreCase Then
                                If UCase(Field.Name) = index Then
                                    bFound = True
                                    Exit For
                                End If
                            Else
                                If Field.Name = index Then
                                    bFound = True
                                    Exit For
                                End If
                            End If
                        Next
                        If bFound Then
                            Return Field
                        Else
                            Return Nothing
                        End If
                    End Get

                End Property

                Public Overrides Function ToString() As String

                    Return "Collection of fields, count = " & MyBase.List.Count.ToString

                End Function

                Public Overloads Function Add() As Field

                    Dim Field As New Field

                    MyBase.List.Add(Field)
                    Return Field

                End Function

                Public Overloads Sub Add(ByVal Field As Field)

                    MyBase.List.Add(Field)

                End Sub

                Public Function Contains(ByVal Field As Field) As Boolean

                    Return List.Contains(Field)

                End Function

                Public Sub CopyTo(ByVal Array() As Field, ByVal intIndex As Integer)

                    List.CopyTo(Array, intIndex)

                End Sub

                Public Sub Insert(ByVal intIndex As Integer, ByVal Field As Field)

                    List.Insert(intIndex, Field)

                End Sub

                Public Function IndexOf(ByVal Field As Field) As Integer

                    Return List.IndexOf(Field)

                End Function

                Public Shadows Sub Remove(ByVal Field As Field)

                    MyBase.List.Remove(Field)

                End Sub

            End Class

#Region " ¤×¤×¤×¤×¤×¤×¤ Properties ¤×¤×¤×¤×¤×¤×¤ "

            Public ReadOnly Property HasChildren() As Boolean

                Get
                    Return Not mstrChildren Is Nothing
                End Get

            End Property

            Public ReadOnly Property HasParent() As Boolean

                Get
                    Return Not mstrParent = vbNullString
                End Get

            End Property

            Public ReadOnly Property IsAChild(ByVal strChild As String) As Boolean

                Get

                    Dim bFound As Boolean
                    Dim r As Integer

                    If mstrChildren Is Nothing Then
                        Return False
                    Else
                        For r = 0 To mstrChildren.GetUpperBound(0)
                            If strChild = mstrChildren(r) Then
                                bFound = True
                                Exit For
                            End If
                        Next
                        If bFound Then
                            Return True
                        Else
                            Return False
                        End If
                    End If

                End Get

            End Property

            Public ReadOnly Property LookupFieldCount() As Integer
                Get
                    Dim intCount As Integer
                    Dim Field As Depot.Tables.Table.Fields.Field

                    For Each Field In mFieldsCollection
                        If Field.UseAsLookup Then
                            intCount += 1
                        End If
                    Next

                    Return intCount
                End Get
            End Property

            Public Property Children() As String()

                Get
                    Return mstrChildren
                End Get
                Set(ByVal value() As String)
                    mstrChildren = value
                End Set

            End Property

            Public Property ClassName() As String
                Get
                    Return mstrClassName
                End Get
                Set(ByVal value As String)
                    mstrClassName = value
                End Set
            End Property

            Public Property ClassNameVar() As String
                Get
                    Return mstrClassNameVar
                End Get
                Set(ByVal value As String)
                    mstrClassNameVar = value
                End Set
            End Property

            Public Property CodeIndex() As Integer
                Get
                    Return mintCodeIndex
                End Get
                Set(ByVal value As Integer)
                    mintCodeIndex = value
                End Set
            End Property

            Public Property CollectionName() As String
                Get
                    Return mstrCollectionName
                End Get
                Set(ByVal value As String)
                    mstrCollectionName = value
                End Set
            End Property

            Public Property CollectionNameVar() As String
                Get
                    Return mstrCollectionNameVar
                End Get
                Set(ByVal value As String)
                    mstrCollectionNameVar = value
                End Set
            End Property

            Public Property CollectionProperty() As String
                Get
                    Return mstrCollectionProperty
                End Get
                Set(ByVal value As String)
                    mstrCollectionProperty = value
                End Set
            End Property

            Public Property CreatedConstants() As Boolean

                Get
                    Return mbCreatedConstants
                End Get
                Set(ByVal value As Boolean)
                    mbCreatedConstants = value
                End Set

            End Property

            Public Property FieldsCollection() As Fields

                Get
                    Return mFieldsCollection
                End Get
                Set(ByVal value As Fields)
                    mFieldsCollection = value
                End Set

            End Property

            Public Property FieldConstantList() As String
                Get
                    Return mstrFieldConstantList
                End Get
                Set(ByVal value As String)
                    mstrFieldConstantList = value
                End Set
            End Property

            Public Property GridDisplayMethod() As String
                Get
                    Return mstrGridDisplayMethod
                End Get
                Set(ByVal value As String)
                    mstrGridDisplayMethod = value
                End Set
            End Property

            Public Property GridName() As String
                Get
                    Return mstrGridName
                End Get
                Set(ByVal value As String)
                    mstrGridName = value
                End Set
            End Property

            Public Property GridFillMethod() As String
                Get
                    Return mstrGridFillMethod
                End Get
                Set(ByVal value As String)
                    mstrGridFillMethod = value
                End Set
            End Property

            Public Property IndexColumn() As Integer
                Get
                    Return mintIndexColumn
                End Get
                Set(ByVal value As Integer)
                    mintIndexColumn = value
                End Set
            End Property

            Public Property IndexVar() As String
                Get
                    Return mstrIndexVar
                End Get
                Set(ByVal value As String)
                    mstrIndexVar = value
                End Set
            End Property

            Public Property LoadedVar() As String
                Get
                    Return mstrLoadedVar
                End Get
                Set(ByVal value As String)
                    mstrLoadedVar = value
                End Set
            End Property

            Public Property Lookup() As String()
                Get
                    Return mstrLookup
                End Get
                Set(ByVal value As String())
                    mstrLookup = value
                End Set
            End Property

            Public Property LookupDimVars() As String
                Get
                    Return mstrLookupDimVars
                End Get
                Set(ByVal value As String)
                    mstrLookupDimVars = value
                End Set
            End Property

            Public Property LookupList() As String
                Get
                    Return mstrLookupList
                End Get
                Set(ByVal value As String)
                    mstrLookupList = value
                End Set
            End Property

            Public Property LookupModularVars() As String
                Get
                    Return mstrLookupModularVars
                End Get
                Set(ByVal value As String)
                    mstrLookupModularVars = value
                End Set
            End Property

            Public Property LookupVars() As String
                Get
                    Return mstrLookupVars
                End Get
                Set(ByVal value As String)
                    mstrLookupVars = value
                End Set
            End Property

            'Public Property NestingLevel() As Integer
            '    Get
            '        Return mintNestingLevel
            '    End Get
            '    Set(ByVal value As Integer)
            '        mintNestingLevel = value
            '    End Set

            'End Property

            'Public Property NestingLevelActual() As Integer
            '    Get
            '        Return mintNestingLevelActual
            '    End Get
            '    Set(ByVal value As Integer)
            '        mintNestingLevelActual = value
            '    End Set

            'End Property

            Public Property NewFieldList() As String
                'Error.  Kill this property.
                Get
                    Return mstrNewFieldList
                End Get
                Set(ByVal value As String)
                    mstrNewFieldList = value
                End Set
            End Property

            Public Property NiceName() As String
                'Depreciated
                Get
                    If Not mstrNiceName Is Nothing Then
                        Return mstrNiceName
                    Else
                        Return mstrTableName
                    End If
                End Get
                Set(ByVal value As String)
                    mstrNiceName = value
                End Set

            End Property

            Public Property ObjectAccessClass() As String
                Get
                    Return mstrObjectAccessClass
                End Get
                Set(ByVal value As String)
                    mstrObjectAccessClass = value
                End Set

            End Property

            Public Property ObjectAccessCollection() As String
                Get
                    Return mstrObjectAccessCollection
                End Get
                Set(ByVal value As String)
                    mstrObjectAccessCollection = value
                End Set

            End Property

            Public Property Parent() As String

                Get
                    Return mstrParent
                End Get
                Set(ByVal value As String)
                    mstrParent = value
                End Set

            End Property

            Public Property SelectList() As String
                Get
                    Return mstrSelectList
                End Get
                Set(ByVal value As String)
                    mstrSelectList = value
                End Set
            End Property

            Public Property SelectListAsParameters() As String
                Get
                    Return mstrSelectListAsParameters
                End Get
                Set(ByVal value As String)
                    mstrSelectListAsParameters = value
                End Set
            End Property

            Public Property TableName() As String

                Get
                    Return mstrTableName
                End Get
                Set(ByVal value As String)
                    mstrTableName = value
                End Set

            End Property

            Public Property TopParent() As Boolean

                Get
                    Return mbTopParent
                End Get
                Set(ByVal value As Boolean)
                    mbTopParent = value
                End Set

            End Property

            Public Property TypeDefClass() As String

                Get
                    Return mstrTypeDefClass
                End Get
                Set(ByVal value As String)
                    mstrTypeDefClass = value
                End Set

            End Property

            Public Property TypeDefCollection() As String

                Get
                    Return mstrTypeDefCollection
                End Get
                Set(ByVal value As String)
                    mstrTypeDefCollection = value
                End Set

            End Property

#End Region

#Region " ¤×¤×¤×¤×¤×¤×¤ Methods ¤×¤×¤×¤×¤×¤×¤ "

            Public Sub New()

                mFieldsCollection = New Fields

            End Sub

            Public Function AddChild(ByVal strChild As String) As Boolean

                Dim bFound As Boolean
                Dim r As Integer
                Dim intNewIndex As Integer


                If Not mstrChildren Is Nothing Then

                    For r = 0 To mstrChildren.GetUpperBound(0)
                        If mstrChildren(r) = strChild Then
                            bFound = True
                        End If
                    Next

                End If

                If Not bFound Then

                    If mstrChildren Is Nothing Then
                        ReDim mstrChildren(0)
                        mstrChildren(0) = strChild
                        ReDim mstrChildrenLink(0)
                        ReDim mstrLinkToChildren(0)
                    Else
                        intNewIndex = mstrChildren.GetUpperBound(0) + 1
                        ReDim Preserve mstrChildren(intNewIndex)
                        mstrChildren(intNewIndex) = strChild
                        ReDim Preserve mstrChildrenLink(intNewIndex)
                        ReDim Preserve mstrLinkToChildren(intNewIndex)
                    End If

                    Return True

                Else

                    Return False

                End If

            End Function


            Public Function AddChildLink(ByVal strChild As String, ByVal strChildLink As String, ByVal strLinkToChild As String) As Boolean

                Dim bFoundChild As Boolean
                'Dim bFoundLink As Boolean
                Dim r As Integer
                Dim intFoundChild As Integer
                'Dim intFoundLinktoChild As Integer
                'Dim intNewIndex As Integer


                If Not mstrChildren Is Nothing And Not mstrLinkToChildren Is Nothing Then

                    intFoundChild = -1
                    'intFoundLinktoChild = -1

                    'For r = 0 To mstrChildrenLink.GetUpperBound(0)
                    '    If mstrChildrenLink(r) = strChildLink Then
                    '        intFoundChildLink = r
                    '        bFound = True
                    '        Exit For
                    '    End If
                    'Next strChild

                    'For r = 0 To mstrLinkToChildren.GetUpperBound(0)
                    '    If mstrLinkToChildren(r) = strLinkToChild Then
                    '        intFoundLinktoChild = r
                    '        'bFound = True And bFound
                    '        bFound = True
                    '        Exit For
                    '    End If
                    'Next

                    For r = 0 To mstrChildren.GetUpperBound(0)
                        If mstrChildren(r) = strChild Then
                            intFoundChild = r
                            'bFound = True And bFound
                            bFoundChild = True
                            Exit For
                        End If
                    Next

                End If

                If bFoundChild Then

                    mstrChildrenLink(intFoundChild) = strChildLink
                    mstrLinkToChildren(intFoundChild) = strLinkToChild

                    Return True

                Else

                    'If Not mstrLinkToChildren Is Nothing Then
                    '    For r = 0 To mstrLinkToChildren.GetUpperBound(0)
                    '        If mstrLinkToChildren(r) = strLinkToChild Then
                    '            intFoundLinktoChild = r
                    '            'bFound = True And bFound
                    '            bFoundLink = True
                    '            Exit For
                    '        End If
                    '    Next
                    'End If

                    'If intFoundLinktoChild = -1 Then

                    '    ReDim mstrChildrenLink(0)
                    '    mstrChildrenLink(0) = strChildLink
                    '    ReDim mstrLinkToChildren(0)
                    '    mstrLinkToChildren(0) = strLinkToChild
                    'Else
                    '    intNewIndex = mstrChildrenLink.GetUpperBound(0) + 1
                    '    ReDim Preserve mstrChildrenLink(intNewIndex)
                    '    mstrChildrenLink(intNewIndex) = strChildLink
                    '    ReDim Preserve mstrLinkToChildren(intNewIndex)
                    '    mstrLinkToChildren(intNewIndex) = strLinkToChild
                    'End If

                    Return False

                End If

            End Function

            Public Sub ClearChildren()

                mstrChildren = Nothing

            End Sub

            Public Sub ClearChildLinks()

                mstrChildrenLink = Nothing
                mstrLinkToChildren = Nothing

            End Sub

            Public Function RemoveChild(ByVal strChild As String) As Boolean

                Dim bFound As Boolean
                Dim r As Integer
                Dim intIndex As Integer
                Dim intMaxIndex As Integer


                If mstrChildren Is Nothing Then
                    Return False
                End If

                intIndex = -1
                intMaxIndex = mstrChildren.GetUpperBound(0)

                For r = 0 To intMaxIndex
                    If mstrChildren(r) = strChild Then
                        intIndex = r
                        bFound = True
                        If r = 0 And intMaxIndex = 0 Then
                            'Found at first and only child
                            mstrChildren = Nothing
                            Return True
                        ElseIf r = intMaxIndex Then
                            'Found at last child and there is more than 1
                            ReDim Preserve mstrChildren(mstrChildren.GetUpperBound(0) - 1)
                            Return True
                        End If
                    End If
                    If bFound And intIndex <> r Then 'If we found one and we moved past it
                        'Move current child up one 
                        mstrChildren(r - 1) = mstrChildren(r)
                    End If
                Next
                If bFound Then
                    'Shorten the size by 1
                    ReDim Preserve mstrChildren(mstrChildren.GetUpperBound(0) - 1)
                End If

            End Function

#End Region

        End Class

        Public Shadows ReadOnly Property Count() As Integer
            Get
                Return MyBase.List.Count
            End Get
        End Property

        Default Public Property Item(ByVal index As Integer) As Table
            Get
                Return CType(MyBase.List.Item(index), Table)
            End Get
            Set(ByVal value As Table)
                MyBase.List.Item(index) = value
            End Set

        End Property

        Default Public Property Item(ByVal index As String, Optional ByVal bIgnoreCase As Boolean = False) As Table
            Get

                Dim bFound As Boolean
                Dim Table As Table


                If bIgnoreCase Then
                    index = UCase(index)
                End If
                Table = Nothing
                For Each Table In MyBase.List
                    If bIgnoreCase Then
                        If UCase(Table.TableName) = index Then
                            bFound = True
                            Exit For
                        End If
                    Else
                        If Table.TableName = index Then
                            bFound = True
                            Exit For
                        End If
                    End If
                Next
                If bFound Then
                    Return Table
                Else
                    Return Nothing
                End If
            End Get

            Set(ByVal value As Table)
                Dim bFound As Boolean
                Dim intIndex As Integer
                Dim Table As Table


                If bIgnoreCase Then
                    index = UCase(index)
                End If
                Table = Nothing
                For Each Table In MyBase.List
                    If bIgnoreCase Then
                        If UCase(Table.TableName) = index Then
                            bFound = True
                            intIndex = MyBase.List.IndexOf(Table)
                            Exit For
                        End If
                    Else
                        If Table.TableName = index Then
                            bFound = True
                            intIndex = MyBase.List.IndexOf(Table)
                            Exit For
                        End If
                    End If
                Next
                If bFound Then
                    MyBase.List.Item(intIndex) = value
                End If
            End Set

        End Property

        Public Overrides Function ToString() As String

            Return "Collection of tables, count = " & MyBase.List.Count.ToString

        End Function

        Public Overloads Function Add() As Table

            Dim Table As New Table

            MyBase.List.Add(Table)
            Return Table

        End Function

        Public Overloads Sub Add(ByVal Table As Table)

            MyBase.List.Add(Table)

        End Sub

        Public Function Contains(ByVal Table As Table) As Boolean

            Return List.Contains(Table)

        End Function

        Public Sub CopyTo(ByVal Array() As Table, ByVal intIndex As Integer)

            List.CopyTo(Array, intIndex)

        End Sub

        Public Sub Insert(ByVal intIndex As Integer, ByVal Table As Table)

            List.Insert(intIndex, Table)

        End Sub

        Public Function IndexOf(ByVal Table As Table) As Integer

            Return List.IndexOf(Table)

        End Function

        Public Shadows Sub Remove(ByVal Table As Table)

            MyBase.List.Remove(Table)

        End Sub

    End Class

#End Region

    Public Property CreateClasses() As Boolean
        Get
            Return mbCreateClasses
        End Get
        Set(ByVal value As Boolean)
            mbCreateClasses = value
        End Set
    End Property

    Public Property CreateCollection() As Boolean
        Get
            Return mbCreateCollection
        End Get
        Set(ByVal value As Boolean)
            mbCreateCollection = value
        End Set
    End Property

    Public Property CreateChangeTables() As Boolean
        Get
            Return mbCreateChangeTables
        End Get
        Set(ByVal value As Boolean)
            mbCreateChangeTables = value
        End Set
    End Property

    Public Property CreateDataAccess() As Boolean
        Get
            Return mbCreateDataAccess
        End Get
        Set(ByVal value As Boolean)
            mbCreateDataAccess = value
        End Set
    End Property

    Public Property CreateFindType() As Boolean
        Get
            Return mbCreateFindType
        End Get
        Set(ByVal value As Boolean)
            mbCreateFindType = value
        End Set
    End Property

    Public Property CreateForm() As Boolean
        Get
            Return mbCreateForm
        End Get
        Set(ByVal value As Boolean)
            mbCreateForm = value
        End Set
    End Property

    Public Property CreateInterface() As Boolean
        Get
            Return mbCreateInterface
        End Get
        Set(ByVal value As Boolean)
            mbCreateInterface = value
        End Set
    End Property

    Public Property CreateStoredProcedure() As Boolean
        Get
            Return mbStoredProcedure
        End Get
        Set(ByVal value As Boolean)
            mbStoredProcedure = value
        End Set
    End Property

    Public Property CreateTopLeveLCollection() As Boolean
        Get
            Return mbCreateTopLeveLCollection
        End Get
        Set(ByVal value As Boolean)
            mbCreateTopLeveLCollection = value
        End Set
    End Property

    Public Property DatabaseType() As eDatabaseType
        Get
            Return meDatabaseType
        End Get
        Set(ByVal value As eDatabaseType)
            meDatabaseType = value
        End Set
    End Property

    Public Property FormCode() As FormData
        Get
            Return mFormCode
        End Get
        Set(ByVal value As FormData)
            mFormCode = value
        End Set
    End Property

    Public Property FormFileName() As String
        Get
            Return mstrFormFileName
        End Get
        Set(ByVal value As String)
            mstrFormFileName = value
        End Set
    End Property

    Public Property FormName() As String
        Get
            Return mstrFormName
        End Get
        Set(ByVal value As String)
            mstrFormName = value
        End Set
    End Property

    Public Property FormType() As eFormType
        Get
            Return meFormType
        End Get
        Set(ByVal value As eFormType)
            meFormType = value
        End Set
    End Property

    Public Property LinkCollection() As Links
        Get
            Return mcolLinks
        End Get
        Set(ByVal value As Links)
            mcolLinks = value
        End Set
    End Property

    Public Property ProjectName() As String
        Get
            Return mstrProjectName
        End Get
        Set(ByVal value As String)
            mstrProjectName = value
        End Set
    End Property

    Public Property ProjectPath() As String
        Get
            Return mstrProjectPath
        End Get
        Set(ByVal value As String)
            mstrProjectPath = value
        End Set
    End Property

    Public Property ProjectSource() As eProjectSource
        Get
            Return meProjectSource
        End Get
        Set(ByVal value As eProjectSource)
            meProjectSource = value
        End Set
    End Property

    Public Property ProjectType() As eProjectType
        Get
            Return meProjectType
        End Get
        Set(ByVal value As eProjectType)
            meProjectType = value
        End Set
    End Property

    Public Property ReadType() As eReadType
        Get
            Return meReadType
        End Get
        Set(ByVal value As eReadType)
            meReadType = value
        End Set
    End Property

    Public Property RemoveTablePrefix() As Boolean
        Get
            Return mbRemoveTablePrefix
        End Get
        Set(ByVal value As Boolean)
            mbRemoveTablePrefix = value
        End Set
    End Property

    Public Property StoredProcedureName() As String
        Get
            Return mstrSPName
        End Get
        Set(ByVal value As String)
            mstrSPName = value
        End Set
    End Property

    Public Property TableCollection() As Tables
        Get
            Return mcolTables
        End Get
        Set(ByVal value As Tables)
            mcolTables = value
        End Set
    End Property

    Public Property TablePrefix() As String
        Get
            Return mstrTablePrefix
        End Get
        Set(ByVal value As String)
            mstrTablePrefix = value
        End Set
    End Property

    Public Property TopLevelName() As String
        Get
            Return mstrTopLevelName
        End Get
        Set(ByVal value As String)
            mstrTopLevelName = value
        End Set
    End Property

    Public Property UseMainDataAccess() As Boolean
        Get
            Return mbMainDA
        End Get
        Set(ByVal value As Boolean)
            mbMainDA = value
        End Set
    End Property


    Public Sub New()

        meReadType = eReadType.DataAdaptor
        mFormCode = New FormData
        mcolTables = New Tables
        mcolLinks = New Links

    End Sub


    Public Sub BuildFieldList(ByVal Table As Depot.Tables.Table, ByVal strModularPrefix As String)

        Dim intCounter As Integer
        Dim intLookupCounter As Integer
        Dim strLookupDimVars As String
        Dim strLookupList As String
        Dim strLookupModularVars As String
        Dim strLookupVars As String
        Dim strQuote As String
        Dim strSelectList As String
        Dim strSelectListAsParameters As String
        Dim strLookup() As String
        Dim colFields As Tables.Table.Fields
        Dim Field As Depot.Tables.Table.Fields.Field

        strLookupDimVars = vbNullString
        strLookupList = vbNullString
        strLookupModularVars = vbNullString
        strLookupVars = vbNullString
        strSelectList = vbNullString
        strSelectListAsParameters = vbNullString
        colFields = Table.FieldsCollection

        ReDim strLookup(9)

        For Each Field In colFields
            'Build a select list
            strSelectList &= Field.Name & ", "
            'Build a select list with parameters
            strSelectListAsParameters &= "{" & intCounter.ToString & "}, "
            intCounter += 1
            'Is this field a lookup field?
            If Field.UseAsLookup Then
                'check size of array
                If intLookupCounter / 10 = Int(intLookupCounter / 10) Then
                    'Should never get here unless there are more than 10 lookup fields, but CYA.
                    ReDim strLookup(strLookup.GetUpperBound(0) + 10)
                End If
                'Save the lookup field in an array
                strLookup(intLookupCounter) = Field.Name
                'Check if we must use quotes
                strQuote = vbNullString
                If Field.UseQuotes Then
                    strQuote = "'"
                End If
                'Build the lists
                strLookupList &= Field.Name & " = " & strQuote & "{" & intLookupCounter.ToString & "}" & strQuote & " AND "
                strLookupVars &= GetPrefixFromDataType(Field.VBDataType) & Field.Name & ", "
                strLookupDimVars &= "byval " & GetPrefixFromDataType(Field.VBDataType) & Field.Name & " as " & Field.VBDataType & ", "
                strLookupModularVars &= strModularPrefix & GetPrefixFromDataType(Field.VBDataType) & Field.Name & ", "
                intLookupCounter += 1
            End If
        Next
        'Clean up
        If intLookupCounter > 0 Then
            ReDim Preserve strLookup(intLookupCounter - 1)
        End If
        strSelectList = Truncate(strSelectList, 2)
        strSelectListAsParameters = Truncate(strSelectListAsParameters, 2)
        strLookupList = Truncate(strLookupList, 5)
        strLookupVars = Truncate(strLookupVars, 2)
        strLookupDimVars = Truncate(strLookupDimVars, 2)
        strLookupModularVars = Truncate(strLookupModularVars, 2)

        With Table
            .Lookup = strLookup
            .LookupDimVars = strLookupDimVars
            .LookupList = strLookupList
            .LookupModularVars = strLookupModularVars
            .LookupVars = strLookupVars
            .SelectList = strSelectList
            .SelectListAsParameters = strSelectListAsParameters
        End With


    End Sub

    Public Function GetTopTable() As Tables.Table

        Dim bFound As Boolean
        Dim Table As Tables.Table = Nothing


        For Each Table In mcolTables
            If Table.TopParent Then
                bFound = True
                Exit For
            End If
        Next

        If bFound Then
            Return Table
        Else
            Return Nothing
        End If

    End Function

    Public Sub SaveDepot(ByVal strPath As String, ByVal Depot As Depot)

        Dim r, rr As Integer
        Dim intFileNumber As Integer


        Try

            intFileNumber = FreeFile()
            FileOpen(intFileNumber, strPath, OpenMode.Binary)
            With Depot
                FilePutObject(intFileNumber, .CreateChangeTables)
                FilePutObject(intFileNumber, .CreateClasses)
                FilePutObject(intFileNumber, .CreateCollection)
                FilePutObject(intFileNumber, .CreateDataAccess)
                FilePutObject(intFileNumber, .CreateForm)
                FilePutObject(intFileNumber, .CreateInterface)
                FilePutObject(intFileNumber, .CreateStoredProcedure)
                FilePutObject(intFileNumber, .CreateTopLeveLCollection)
                FilePutObject(intFileNumber, .FormType)
                FilePutObject(intFileNumber, .ProjectName)
                FilePutObject(intFileNumber, .ProjectPath)
                FilePutObject(intFileNumber, .ProjectSource)
                FilePutObject(intFileNumber, .RemoveTablePrefix)
                r = .TableCollection.Count
                FilePutObject(intFileNumber, r)
                For r = 0 To r - 1
                    With .TableCollection(r)
                        FilePutObject(intFileNumber, .TableName)
                        rr = .FieldsCollection.Count
                        FilePutObject(intFileNumber, rr)
                        For rr = 0 To rr - 1
                            With .FieldsCollection(rr)
                                FilePutObject(intFileNumber, .DataLength)
                                FilePutObject(intFileNumber, .DBDataType)
                                FilePutObject(intFileNumber, .Hidden)
                                FilePutObject(intFileNumber, .Name)
                                FilePutObject(intFileNumber, .ReadOnly)
                                FilePutObject(intFileNumber, .Required)
                                FilePutObject(intFileNumber, .Ordinal)
                                FilePutObject(intFileNumber, .PKey)
                                FilePutObject(intFileNumber, .VBDataType)
                            End With
                        Next
                    End With
                Next
            End With

        Catch ex As Exception

        Finally

            FileClose(intFileNumber)

        End Try

    End Sub

    'Public Sub LoadDepot(ByVal strPath As String, ByRef Depot As Depot)

    '    Dim r, rr As Integer
    '    Dim intFileNumber As Integer
    '    Dim Table As Depot.Tables.Table
    '    Dim Field As Depot.Tables.Table.Fields.Field


    '    intFileNumber = FreeFile()
    '    FileOpen(intFileNumber, strPath, OpenMode.Binary)
    '    With Depot
    '        FileGetObject(intFileNumber, .CreateChangeTables)
    '        FileGetObject(intFileNumber, .CreateClasses)
    '        FileGetObject(intFileNumber, .CreateCollection)
    '        FileGetObject(intFileNumber, .CreateDataAccess)
    '        FileGetObject(intFileNumber, .CreateForm)
    '        FileGetObject(intFileNumber, .CreateInterface)
    '        FileGetObject(intFileNumber, .CreateStoredProcedure)
    '        FileGetObject(intFileNumber, .CreateTopLeveLCollection)
    '        FileGetObject(intFileNumber, .FormType)
    '        FileGetObject(intFileNumber, .ProjectName)
    '        FileGetObject(intFileNumber, .ProjectPath)
    '        FileGetObject(intFileNumber, .ProjectSource)
    '        FileGetObject(intFileNumber, .RemoveTablePrefix)
    '        FileGetObject(intFileNumber, r)
    '        .TableCollection = New Depot.Tables
    '        For r = 0 To r - 1
    '            Table = .TableCollection.Add
    '            FileGetObject(intFileNumber, Table.TableName)
    '            FileGetObject(intFileNumber, rr)
    '            .TableCollection(r).FieldsCollection = New Depot.Tables.Table.Fields
    '            For rr = 0 To rr - 1
    '                Field = .TableCollection(r).FieldsCollection.Add
    '                With Field
    '                    FileGetObject(intFileNumber, .DataLength)
    '                    FileGetObject(intFileNumber, .DBDataType)
    '                    FileGetObject(intFileNumber, .Hidden)
    '                    FileGetObject(intFileNumber, .Name)
    '                    FileGetObject(intFileNumber, .ReadOnly)
    '                    FileGetObject(intFileNumber, .Required)
    '                    FileGetObject(intFileNumber, .Ordinal)
    '                    FileGetObject(intFileNumber, .PKey)
    '                    FileGetObject(intFileNumber, .VBDataType)
    '                    .UseAsLookup = .PKey
    '                End With
    '            Next
    '        Next
    '    End With

    '    FileClose(intFileNumber)

    'End Sub

End Class
