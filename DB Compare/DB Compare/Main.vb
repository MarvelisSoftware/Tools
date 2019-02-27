'Created by John Maher

Imports System.Math


Public Class frmMain


    Private Const mconFilterMargin As Integer = 29


    Private Enum eFilterState
        Hide
        Show
    End Enum

    Private Enum eImage
        Field
        Table
        FieldSelected
        TableSelected
    End Enum

    Private Enum eWhich
        None
        Source
        Target
    End Enum


    Private meFilter As eFilterState
    Private mbInReload As Boolean
    Private mdictSource As Dictionary(Of String, DataTable)
    Private mdictTarget As Dictionary(Of String, DataTable)
    Private mlisTree As List(Of TreeNode)


    Private Property datProduction As Object


    Public Sub New()

        InitializeComponent()
        ToggleFilter(meFilter)
        CreateContextMenu()

    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim r As Integer
        Dim strServers As String()

        strServers = GetActiveServers()
        For r = 0 To strServers.GetUpperBound(0)
            cboServerSource.Items.Add(strServers(r))
            cboServerTarget.Items.Add(strServers(r))
        Next

        txtPasswordSource.Text = My.Settings.PasswordSource
        txtPasswordTarget.Text = My.Settings.PasswordTarget
        txtUserSource.Text = My.Settings.UserSource
        txtUserTarget.Text = My.Settings.UserTarget
        cboServerSource.Text = My.Settings.ServerSource
        cboServerTarget.Text = My.Settings.ServerTarget
        cboDatabaseSource.Text = My.Settings.DBSource
        cboDatabaseTarget.Text = My.Settings.DBTarget

    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click

        Dim frmAboutbox As frmAbout


        frmAboutbox = New frmAbout
        frmAboutbox.ShowDialog()

    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        Me.Dispose()

    End Sub

    Private Sub btnClearAll_Click(sender As Object, e As EventArgs) Handles btnClearAll.Click

        Dim Node As TreeNode


        For Each Node In treSource.Nodes
            Node.Checked = False
        Next

    End Sub

    Private Sub btnAllSelect_Click(sender As Object, e As EventArgs) Handles btnAllSelect.Click

        Dim Node As TreeNode


        For Each Node In treSource.Nodes
            Node.Checked = True
        Next

    End Sub

    Private Sub btnDiff_Click(sender As Object, e As EventArgs) Handles btnDiff.Click

        Diff()

    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click

        mbInReload = True
        PopulateTables()
        mbInReload = False

    End Sub

    Private Sub cboSource_Changed(sender As Object, e As EventArgs) Handles cboServerSource.SelectedIndexChanged, cboServerSource.Validating, txtPasswordSource.Validated, txtUserSource.Validated

        Try

            Me.Cursor = Cursors.WaitCursor

            FillDBConbo(eWhich.Source, DirectCast(sender, Control).Parent)

        Catch ex As Exception

            MsgBox("Error in cboSource_Changed: " & ex.Message)

        Finally

            Me.Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub cboTarget_Changed(sender As Object, e As EventArgs) Handles cboServerTarget.SelectedIndexChanged, cboServerTarget.Validating, txtPasswordTarget.Validated, txtUserTarget.Validated

        Try

            Me.Cursor = Cursors.WaitCursor

            FillDBConbo(eWhich.Target, DirectCast(sender, Control).Parent)

        Catch ex As Exception

            MsgBox("Error in cboTarget_Changed: " & ex.Message)

        Finally

            Me.Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub treSource_AfterCheck(sender As Object, e As TreeViewEventArgs)

        Dim ChildNode As TreeNode
        Dim ParentNode As TreeNode


        ParentNode = e.Node
        For Each ChildNode In ParentNode.Nodes
            ChildNode.Checked = ParentNode.Checked
        Next

    End Sub

    Private Sub treSource_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs)

        Dim Node As TreeNode = e.Node


        If Node.Nodes.Count = 0 Then
            PopulateFields(Node)
        End If

    End Sub

    Private Sub txtFilterTable_TextChanged(sender As Object, e As EventArgs) Handles txtFilterTable.TextChanged

        FilterTree(txtFilterTable.Text)

    End Sub


    Private Sub CreateContextMenu()

        Dim cms As ContextMenu


        cms = New ContextMenu
        cms.MenuItems.Add("Toggle &Filter")
        AddHandler (cms.MenuItems(0).Click), AddressOf ToggleFilter

        treSource.ContextMenu = cms

    End Sub

    Private Sub FillDBConbo(Which As eWhich, ctl As Control)

        Dim cboDatabase As ComboBox = Nothing
        Dim strDB As String()
        Dim grp As GroupBox
        Dim intIndex As Integer
        Dim strPassword As String = String.Empty
        Dim strServer As String = String.Empty
        Dim strUser As String = String.Empty


        Try

            grp = DirectCast(ctl, GroupBox)
            Select Case Which
                Case eWhich.Source
                    cboDatabase = DirectCast(grp.Controls("cboDatabaseSource"), ComboBox)
                    strPassword = grp.Controls("txtPasswordSource").Text
                    strServer = grp.Controls("cboServerSource").Text
                    strUser = grp.Controls("txtUserSource").Text
                Case eWhich.Target
                    cboDatabase = DirectCast(grp.Controls("cboDatabaseTarget"), ComboBox)
                    strPassword = grp.Controls("txtPasswordTarget").Text
                    strServer = grp.Controls("cboServerTarget").Text
                    strUser = grp.Controls("txtUserTarget").Text
            End Select

            cboDatabase.Items.Clear()
            cboDatabase.Text = String.Empty
            strDB = DA.GetDatabases(strServer, strUser, strPassword)

            If Not strDB Is Nothing Then

                With cboDatabase
                    For intIndex = 0 To strDB.GetUpperBound(0)
                        .Items.Add(strDB(intIndex))
                    Next
                End With

            End If

        Catch ex As Exception

            MsgBox("Error in FillDBConbo: " & ex.Message)

        End Try

    End Sub

    Private Sub FilterTree(strFilter As String)

        Dim CurrentNode As TreeNode


        Try

            Me.Cursor = Cursors.WaitCursor

            With treSource

                .BeginUpdate()
                .Nodes.Clear()

                For Each CurrentNode In mlisTree
                    If CurrentNode.Text.Length >= strFilter.Length AndAlso CurrentNode.Text.Substring(0, strFilter.Length).ToUpper = strFilter.ToUpper Or strFilter = String.Empty Then
                        treSource.Nodes.Add(CurrentNode)
                    End If
                Next
                .EndUpdate()

            End With

        Finally
            Me.Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub PopulateFields(Node As TreeNode)

        Dim r As Integer
        Dim strField As String
        Dim strFields() As String
        Dim NewNode As TreeNode


        Try

            strFields = DA.GetFields(Node.Text, cboServerSource.Text, cboDatabaseSource.Text, txtUserSource.Text, txtPasswordSource.Text)

            For Each strField In strFields

                NewNode = New TreeNode(strFields(r), eImage.Field, eImage.FieldSelected)
                Node.Nodes.Add(NewNode)
                r += 1

            Next


        Catch ex As Exception

            MsgBox("Error in GetTables: " & ex.Message)

        End Try

    End Sub

    Private Sub PopulateTables()

        Dim r As Integer
        Dim Node As TreeNode
        Dim strTables() As String


        Try

            strTables = DA.GetTables(cboServerSource.Text, cboDatabaseSource.Text, txtUserSource.Text, txtPasswordSource.Text)
            mlisTree = New List(Of TreeNode)()

            With treSource

                .Nodes.Clear()
                .BeginUpdate()
                For r = 0 To strTables.GetUpperBound(0)
                    Node = New TreeNode(strTables(r), eImage.Table, eImage.TableSelected)
                    .Nodes.Add(Node)
                    mlisTree.Add(Node)
                Next
                .Sort()
                .EndUpdate()

            End With

            If txtFilterTable.Text <> String.Empty Then
                FilterTree(txtFilterTable.Text)
            End If


        Catch ex As Exception

            MsgBox("Error in GetTables: " & ex.Message)

        End Try

    End Sub

    Private Sub Diff()

        Dim colS As DataColumnCollection
        Dim colT As DataColumnCollection
        Dim DGV As DataGridView
        Dim DT As DataTable = Nothing
        Dim DTS As DataTable
        Dim DTT As DataTable
        Dim strMessage As String = String.Empty
        Dim TP As TabPage
        Dim strTableName As String
        Dim TableNode As TreeNode


        mdictSource = New Dictionary(Of String, DataTable)
        mdictTarget = New Dictionary(Of String, DataTable)
        tcTableResults.TabPages.Clear()

        For Each TableNode In treSource.Nodes

            If TableNode.Checked Then

                strTableName = TableNode.Text
                DGV = New DataGridView
                DGV.AllowUserToAddRows = False
                DGV.Dock = DockStyle.Fill
                DGV.Name = "dgv" & strTableName
                DGV.Visible = True
                TP = New TabPage
                TP.Name = "tp" & strTableName
                TP.Text = strTableName
                TP.Controls.Add(DGV)
                tcTableResults.TabPages.Add(TP)

                DTS = DA.GetTable(strTableName, cboServerSource.Text, cboDatabaseSource.Text, txtUserSource.Text, txtPasswordSource.Text)
                mdictSource.Add(strTableName, DT)
                DTT = DA.GetTable(strTableName, cboServerTarget.Text, cboDatabaseTarget.Text, txtUserTarget.Text, txtPasswordTarget.Text)
                mdictTarget.Add(strTableName, DT)

                If PrimarykeyCountMatches(DTS, DTT) Then

                    If PrimarykeysMatch(DGV, DTS, DTT) Then

                        If ColumnCountMatches(DTS, DTT) Then

                            If ColumnsMatch(DGV, DTS, DTT) Then

                                colS = DTS.Columns
                                DTS = DA.GetData(strTableName, colS, cboServerSource.Text, cboDatabaseSource.Text, txtUserSource.Text, txtPasswordSource.Text)
                                colT = DTS.Columns
                                DTT = DA.GetData(strTableName, colT, cboServerTarget.Text, cboDatabaseTarget.Text, txtUserTarget.Text, txtPasswordTarget.Text)

                                If DataMatches(DGV, DTS, colS, DTT, colT) Then

                                    strMessage &= strTableName & " matches." & Environment.NewLine

                                End If

                            End If

                        Else

                            DisplayDiffColumns(DGV, DTS, DTT)

                        End If

                    End If

                Else

                    DisplayDiffPrimarykeys(DGV, DTS, DTT)

                End If

            End If

        Next

        If strMessage <> String.Empty Then
            MessageBox.Show(strMessage, "Results")
        End If

    End Sub

    Private Sub DisplayDiffColumns(DGV As DataGridView, DTS As DataTable, DTT As DataTable)

        Dim DC As DataGridViewColumn
        Dim DR As DataGridViewRow
        Dim intNewRow As Integer
        Dim intOffset As Integer
        Dim r As Integer

        For r = 0 To DTS.Columns.Count - 1
            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Source: Field " & (r + 1).ToString
            DC.ReadOnly = True
            DGV.Columns.Add(DC)
        Next
        For r = 0 To DTT.Columns.Count - 1
            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Target: Field " & (r + 1).ToString
            DC.ReadOnly = True
            DGV.Columns.Add(DC)
        Next
        intNewRow = DGV.Rows.Add()
        DR = DGV.Rows(intNewRow)
        DR.HeaderCell.Value = "Fields"
        For r = 0 To DTS.Columns.Count - 1
            DR.Cells(r).Value = DTS.Columns(r)
        Next
        intOffset = r
        For r = 0 To DTT.Columns.Count - 1
            DR.Cells(r + intOffset).Value = DTT.Columns(r)
        Next

    End Sub

    Private Sub DisplayDiffPrimarykeys(DGV As DataGridView, DTS As DataTable, DTT As DataTable)

        Dim DC As DataGridViewColumn
        Dim DR As DataGridViewRow
        Dim intNewRow As Integer
        Dim intOffset As Integer
        Dim r As Integer

        For r = 0 To DTS.PrimaryKey.Length - 1
            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Source: PKey" & (r + 1).ToString
            DC.ReadOnly = True
            DGV.Columns.Add(DC)
        Next
        For r = 0 To DTT.PrimaryKey.Length - 1
            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Target: PKey" & (r + 1).ToString
            DC.ReadOnly = True
            DGV.Columns.Add(DC)
        Next
        intNewRow = DGV.Rows.Add()
        DR = DGV.Rows(intNewRow)
        DR.HeaderCell.Value = "PKeys"
        For r = 0 To DTS.PrimaryKey.Length - 1
            DR.Cells(r).Value = DTS.PrimaryKey(r)
        Next
        intOffset = r
        For r = 0 To DTT.PrimaryKey.Length - 1
            DR.Cells(r + intOffset).Value = DTT.PrimaryKey(r)
        Next

    End Sub

    Private Function PrimarykeyCountMatches(DTS As DataTable, DTT As DataTable) As Boolean

        Return DTS.PrimaryKey.Length = DTT.PrimaryKey.Length

    End Function

    Private Function PrimarykeysMatch(DGV As DataGridView, DTS As DataTable, DTT As DataTable) As Boolean

        Dim DC As DataGridViewColumn
        Dim dictDiffSource As Dictionary(Of Integer, String)
        Dim dictDiffTarget As Dictionary(Of Integer, String)
        Dim DR As DataGridViewRow
        Dim KeyValue As KeyValuePair(Of Integer, String)
        Dim bMatch As Boolean = False
        Dim intNewRow As Integer
        Dim s As IOrderedEnumerable(Of DataColumn)
        Dim strSCaption As String
        Dim intSIndex As Integer
        Dim t As IOrderedEnumerable(Of DataColumn)
        Dim strTCaption As String
        Dim intTIndex As Integer

        s = DTS.PrimaryKey.OrderBy(Function(c) c.Caption)
        t = DTT.PrimaryKey.OrderBy(Function(c) c.Caption)

        dictDiffSource = New Dictionary(Of Integer, String)
        dictDiffTarget = New Dictionary(Of Integer, String)

        Do
            strSCaption = s(intSIndex).Caption.ToLower
            strTCaption = t(intTIndex).Caption.ToLower
            If strSCaption = strTCaption Then
                If s(intSIndex).DataType <> t(intTIndex).DataType OrElse s(intSIndex).MaxLength <> t(intSIndex).MaxLength Then
                    dictDiffSource.Add(intSIndex, "DataType Mismatch: " & s(intSIndex).DataType.FullName & "(" & s(intSIndex).MaxLength.ToString & ")")
                    dictDiffTarget.Add(intTIndex, "DataType Mismatch: " & t(intTIndex).DataType.FullName & "(" & t(intSIndex).MaxLength.ToString & ")")
                End If
                intSIndex += 1
                intTIndex += 1
            Else
                If strSCaption > strTCaption Then
                    dictDiffTarget.Add(intTIndex, "Missing from Source")
                    intTIndex += 1
                Else 'strSCaption < strTCaption
                    dictDiffSource.Add(intSIndex, "Missing from Target")
                    intSIndex += 1
                End If
            End If
        Loop While intSIndex < s.Count And intTIndex < t.Count


        While intSIndex < s.Count
            dictDiffSource.Add(intSIndex, "Missing from Target")
            intSIndex += 1
        End While

        While intTIndex < t.Count
            dictDiffTarget.Add(intTIndex, "Missing from Source")
            intTIndex += 1
        End While


        If dictDiffSource.Count = 0 And dictDiffTarget.Count = 0 Then
            bMatch = True
        Else

            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Table"
            DC.ReadOnly = True
            DGV.Columns.Add(DC)
            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Column"
            DC.ReadOnly = True
            DGV.Columns.Add(DC)
            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Status"
            DC.ReadOnly = True
            DGV.Columns.Add(DC)

            For Each KeyValue In dictDiffSource
                intNewRow = DGV.Rows.Add()
                DR = DGV.Rows(intNewRow)
                DR.Cells(0).Value = "Source"
                DR.Cells(1).Value = s(KeyValue.Key).Caption
                DR.Cells(2).Value = KeyValue.Value
            Next

            For Each KeyValue In dictDiffTarget
                intNewRow = DGV.Rows.Add()
                DR = DGV.Rows(intNewRow)
                DR.Cells(0).Value = "Target"
                DR.Cells(1).Value = t(KeyValue.Key).Caption
                DR.Cells(2).Value = KeyValue.Value
            Next

        End If

        Return bMatch

    End Function

    Private Function ColumnCountMatches(DTS As DataTable, DTT As DataTable) As Boolean

        Return DTS.Columns.Count = DTT.Columns.Count

    End Function

    Private Function ColumnsMatch(DGV As DataGridView, DTS As DataTable, DTT As DataTable) As Boolean

        Dim DC As DataGridViewColumn
        Dim dictDiffSource As Dictionary(Of Integer, String)
        Dim dictDiffTarget As Dictionary(Of Integer, String)
        Dim DR As DataGridViewRow
        Dim KeyValue As KeyValuePair(Of Integer, String)
        Dim bMatch As Boolean = False
        Dim intNewRow As Integer
        Dim s As DataColumn()
        Dim strSCaption As String
        Dim intSIndex As Integer = 1
        Dim t As DataColumn()
        Dim strTCaption As String
        Dim intTIndex As Integer = 1

        ReDim s(DTS.Columns.Count)
        ReDim t(DTT.Columns.Count)
        DTS.Columns.CopyTo(s, 1)
        DTT.Columns.CopyTo(t, 1)

        dictDiffSource = New Dictionary(Of Integer, String)
        dictDiffTarget = New Dictionary(Of Integer, String)

        Do
            strSCaption = s(intSIndex).Caption.ToLower
            strTCaption = t(intTIndex).Caption.ToLower
            If strSCaption = strTCaption Then
                If s(intSIndex).DataType <> t(intTIndex).DataType OrElse s(intSIndex).MaxLength <> t(intSIndex).MaxLength Then
                    dictDiffSource.Add(intSIndex, "DataType Mismatch: " & s(intSIndex).DataType.FullName & "(" & s(intSIndex).MaxLength.ToString & ")")
                    dictDiffTarget.Add(intTIndex, "DataType Mismatch: " & t(intTIndex).DataType.FullName & "(" & t(intSIndex).MaxLength.ToString & ")")
                End If
                intSIndex += 1
                intTIndex += 1
            Else
                If strSCaption > strTCaption Then
                    dictDiffTarget.Add(intTIndex, "Missing from Source")
                    intTIndex += 1
                Else 'strSCaption < strTCaption
                    dictDiffSource.Add(intSIndex, "Missing from Target")
                    intSIndex += 1
                End If
            End If
        Loop While intSIndex < s.Count And intTIndex < t.Count


        While intSIndex < s.Count
            dictDiffSource.Add(intSIndex, "Missing from Target")
            intSIndex += 1
        End While

        While intTIndex < t.Count
            dictDiffTarget.Add(intTIndex, "Missing from Source")
            intTIndex += 1
        End While


        If dictDiffSource.Count = 0 And dictDiffTarget.Count = 0 Then
            bMatch = True
        Else

            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Table"
            DC.ReadOnly = True
            DGV.Columns.Add(DC)
            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Column"
            DC.ReadOnly = True
            DGV.Columns.Add(DC)
            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Status"
            DC.ReadOnly = True
            DGV.Columns.Add(DC)

            For Each KeyValue In dictDiffSource
                intNewRow = DGV.Rows.Add()
                DR = DGV.Rows(intNewRow)
                DR.Cells(0).Value = "Source"
                DR.Cells(1).Value = s(KeyValue.Key).Caption
                DR.Cells(2).Value = KeyValue.Value
            Next

            For Each KeyValue In dictDiffTarget
                intNewRow = DGV.Rows.Add()
                DR = DGV.Rows(intNewRow)
                DR.Cells(0).Value = "Target"
                DR.Cells(1).Value = t(KeyValue.Key).Caption
                DR.Cells(2).Value = KeyValue.Value
            Next

        End If

        Return bMatch

    End Function

    Private Function DataMatches(DGV As DataGridView, DTS As DataTable, colS As DataColumnCollection, DTT As DataTable, colT As DataColumnCollection) As Boolean

        Dim DC As DataGridViewColumn
        Dim dictDiffSource As Dictionary(Of String, String)
        Dim dictDiffTarget As Dictionary(Of String, String)
        Dim DR As DataGridViewRow
        Dim DRR As DataRow()
        Dim DRS As DataRow
        Dim DRT As DataRow
        Dim KeyValue As KeyValuePair(Of String, String)
        Dim bMatch As Boolean = False
        Dim intNewRow As Integer
        Dim strPKey As String = String.Empty
        Dim PKeys As Dictionary(Of String, String)
        Dim strPKeyValue As String = String.Empty
        Dim r, rr As Integer


        dictDiffSource = New Dictionary(Of String, String)
        dictDiffTarget = New Dictionary(Of String, String)
        PKeys = New Dictionary(Of String, String)

        For r = 0 To DTS.Rows.Count - 1

            DRS = DTS.Rows(r)
            DRR = DTT.Select(DA.PKeyWhere(DTS.PrimaryKey, DTS, DRS))
            strPKey = DA.PrimaryKey(DTS.PrimaryKey)
            strPKeyValue = DA.PrimaryKeyValue(DTS.PrimaryKey, DTS, DRS)
            PKeys.Add(strPKeyValue, strPKey)
            If DRR.Length = 0 Then
                dictDiffSource.Add(strPKeyValue, "Missing from target")
            ElseIf DRR.Length > 1 Then
                dictDiffTarget.Add(strPKeyValue, "Multiple records in target (" & DRR.Length.ToString & ")")
            Else
                'PKey matches
                For rr = 0 To DTS.Columns.Count - 1
                    If DTS.Rows(r).Item(rr).ToString <> DRR(0).Item(rr).ToString Then
                        dictDiffSource.Add(strPKeyValue & " - " & DTS.Columns(rr).ColumnName, "Values different - " & DTS.Rows(r).Item(rr).ToString)
                        dictDiffTarget.Add(strPKeyValue & " - " & DTT.Columns(rr).ColumnName, "Values different - " & DTT.Rows(r).Item(rr).ToString)
                    End If
                Next

            End If

        Next

        For r = 0 To DTT.Rows.Count - 1

            DRT = DTT.Rows(r)
            strPKeyValue = DA.PrimaryKeyValue(DTT.PrimaryKey, DTT, DRt)
            If Not PKeys.ContainsKey(strPKeyValue) Then
                dictDiffTarget.Add(strPKeyValue, "Missing from target")
            End If

        Next


        If dictDiffSource.Count = 0 And dictDiffTarget.Count = 0 Then
            bMatch = True
        Else

            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Table"
            DC.ReadOnly = True
            DGV.Columns.Add(DC)
            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Column"
            DC.ReadOnly = True
            DGV.Columns.Add(DC)
            DC = New DataGridViewTextBoxColumn
            DC.HeaderText = "Status"
            DC.ReadOnly = True
            DGV.Columns.Add(DC)

            For Each KeyValue In dictDiffSource
                intNewRow = DGV.Rows.Add()
                DR = DGV.Rows(intNewRow)
                DR.Cells(0).Value = "Source"
                DR.Cells(1).Value = KeyValue.Key
                DR.Cells(2).Value = KeyValue.Value
            Next

            For Each KeyValue In dictDiffTarget
                intNewRow = DGV.Rows.Add()
                DR = DGV.Rows(intNewRow)
                DR.Cells(0).Value = "Target"
                DR.Cells(1).Value = KeyValue.Key
                DR.Cells(2).Value = KeyValue.Value
            Next

        End If

        Return bMatch

    End Function

    Private Sub ToggleFilter()

        meFilter = Abs(Not -meFilter)
        ToggleFilter(meFilter)

    End Sub

    Private Sub ToggleFilter(eWhich As eFilterState)

        Select Case eWhich
            Case eFilterState.Hide
                treSource.Top = 0
                treSource.Height += mconFilterMargin
                txtFilterTable.Visible = False

            Case eFilterState.Show
                treSource.Top = mconFilterMargin
                treSource.Height -= mconFilterMargin
                txtFilterTable.Visible = True

        End Select

    End Sub

End Class
