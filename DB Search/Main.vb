Public Class frmMain

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click

        Dim frmAboutbox As frmAbout


        frmAboutbox = New frmAbout
        frmAboutbox.ShowDialog()

    End Sub

    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click

        GetStoredProceduress()
        GetTables()
        GetViews()

    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click

        Dim r As Integer
        Dim strField As String
        Dim strTable As String
        Dim Field As Depot.Tables.Table.Fields.Field
        Dim colFields As Depot.Tables.Table.Fields = Nothing
        Dim strSearch As String
        Dim Table As Depot.Tables.Table
        Dim colTables As Depot.Tables


        tabSearchOptions.SelectedTab = tabSearchOptions.TabPages(0)

        strSearch = LCase(txtString.Text)
        lstResults.Items.Clear()
        txtResults.Text = String.Empty
        colTables = New Depot.Tables

        For r = 0 To lstTables.Items.Count - 1
            strTable = LCase(lstTables.Items(r).ToString)
            Table = New Depot.Tables.Table
            Table.NiceName = strTable
            colTables.Add(Table)
            If InStr(strTable, strSearch) > 0 Then
                AddToResults("Table Name: '" & lstTables.Items(r).ToString & "'")
            End If
        Next

        For r = 0 To lstTables.Items.Count - 1
            Table = colTables(r)
            lstTables.SelectedIndex = r
            lstFields.Items.Clear()
            strTable = Table.NiceName
            DA.GetFields(strTable, colFields)
            For Each Field In colFields
                strField = LCase(Field.Name)
                lstFields.Items.Add(strField)
                If InStr(strField, strSearch) > 0 Then
                    AddToResults("Field Name: '" & Field.Name & "', from Table: '" & lstTables.Items(r).ToString & "'")
                End If
            Next
            lstFields.Refresh()
        Next

    End Sub

    Private Sub btnSearchData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchData.Click

        Dim r As Integer
        Dim strField As String
        Dim strSearch As String
        Dim strTable As String
        Dim Field As Depot.Tables.Table.Fields.Field
        Dim colFields As Depot.Tables.Table.Fields = Nothing


        tabSearchOptions.SelectedTab = tabSearchOptions.TabPages(0)

        strSearch = LCase(txtString.Text)
        lstResults.Items.Clear()
        txtResults.Text = String.Empty

        For r = 0 To lstTables.Items.Count - 1
            lstTables.SelectedIndex = r
            lstFields.Items.Clear()
            strTable = LCase(lstTables.Items(r).ToString)
            DA.GetFields(strTable, colFields)
            For Each Field In colFields
                strField = LCase(Field.Name)
                lstFields.Items.Add(strField)
            Next
            For Each Field In colFields
                strField = LCase(Field.Name)
                lstFields.SelectedIndex = lstFields.Items.Count - 1
                If Field.IsString Then
                    If DA.SearchData(strTable, strField, strSearch) Then
                        AddToResults("Column Name: '" & Field.Name & "', from Table: '" & lstTables.Items(r).ToString & "'")
                    End If
                End If
            Next
            lstFields.Refresh()
        Next

    End Sub

    Private Sub btnSearchSP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchSP.Click

        Dim r As Integer
        Dim strSearch As String
        Dim StoredProcedure As ItemInformation
        Dim lisStoredProcedures As List(Of ItemInformation)


        tabSearchOptions.SelectedTab = tabSearchOptions.TabPages(1)

        strSearch = LCase(txtString.Text)
        lstResults.Items.Clear()
        txtResults.Text = String.Empty
        lisStoredProcedures = GetStoredProceduressDetails()

        For Each StoredProcedure In lisStoredProcedures
            With StoredProcedure
                'hilite the SP we are on.
                For r = 0 To lstSPs.Items.Count - 1
                    If lstSPs.Items(r).ToString.ToLower = .Name.ToLower Then
                        lstSPs.SelectedIndex = r
                        lstSPs.Refresh()
                        Exit For
                    End If
                Next

                If .Name.ToLower.Contains(strSearch) Or _
                            .Definition.ToLower.Contains(strSearch) Then
                    AddToResults("Stored Procedure: '" & .Name & "'")
                End If
            End With
        Next

    End Sub

    Private Sub btnSearchView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchView.Click

        Dim r As Integer
        Dim strSearch As String
        Dim ViewDetail As ItemInformation
        Dim lisViewDetails As List(Of ItemInformation)


        tabSearchOptions.SelectedTab = tabSearchOptions.TabPages(2)

        strSearch = LCase(txtString.Text)
        lstResults.Items.Clear()
        txtResults.Text = String.Empty
        lisViewDetails = DA.GetViewsDetails()

        For Each ViewDetail In lisViewDetails
            With ViewDetail
                'hilite the View we are on.
                For r = 0 To lstViews.Items.Count - 1
                    If lstViews.Items(r).ToString.ToLower = .Name.ToLower Then
                        lstViews.SelectedIndex = r
                        lstViews.Refresh()
                        Exit For
                    End If
                Next

                If .Name.ToLower.Contains(strSearch) Or _
                            .Definition.ToLower.Contains(strSearch) Then
                    AddToResults("View: '" & .Name & "'")
                End If
            End With
        Next

    End Sub

    Private Sub cboDatabase_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDatabase.TextChanged

        Try

            Me.Cursor = Cursors.WaitCursor

            DA.Disconnect()
            DA.Database = cboDatabase.Text

        Catch ex As Exception

            MsgBox("Error in cboDatabase_TextChanged: " & ex.Message)

        Finally

            Me.Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub cboServer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboServer.TextChanged

        Try

            Me.Cursor = Cursors.WaitCursor

            DA.Disconnect()
            DA.Server = cboServer.Text
            FillDBConbo()

        Catch ex As Exception

            MsgBox("Error in cboServer_TextChanged: " & ex.Message)

        Finally

            Me.Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim r As Integer
        Dim strServers As String()

        strServers = GetActiveServers()
        For r = 0 To strServers.GetUpperBound(0)
            cboServer.Items.Add(strServers(r))
        Next

        txtUser.Text = My.Settings.User
        txtPassword.Text = My.Settings.Password
        cboServer.Text = "juliet"
        cboDatabase.Text = "iERP85_Dev"

        DA.DatabaseType = eDatabaseType.SQLServer

    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged

        Try

            DA.Password = txtPassword.Text

        Catch ex As Exception

            MsgBox("Error in txtPassword_TextChanged: " & ex.Message)

        End Try

    End Sub

    Private Sub txtUser_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUser.TextChanged

        Try

            DA.UserID = txtUser.Text

        Catch ex As Exception

            MsgBox("Error in txtUser_TextChanged: " & ex.Message)

        End Try

    End Sub

    Private Sub AddToResults(ByVal strText As String)

        lstResults.Items.Add(strText)
        lstResults.Refresh()
        txtResults.Text &= strText & Environment.NewLine
        txtResults.Refresh()

    End Sub

    Private Sub FillDBConbo()

        Dim intIndex As Integer
        Dim strDB As String()


        Try

            strDB = DA.GetDatabases(cboServer.Text)

            If Not strDB Is Nothing Then

                With cboDatabase
                    .Items.Clear()
                    For intIndex = 0 To strDB.GetUpperBound(0)
                        .Items.Add(strDB(intIndex))
                    Next
                End With

            End If

        Catch ex As Exception

            MsgBox("Error in FillDBConbo: " & ex.Message)

        End Try

    End Sub

    Private Sub GetStoredProceduress()

        Try

            DA.Init()
            If DA.Connected() Then

                With lstSPs

                    .Items.Clear()
                    .Items.AddRange(DA.GetStoredProceduress.ToArray)

                End With

            End If


        Catch ex As Exception

            MsgBox("Error in GetStoredProceduress: " & ex.Message)

        End Try

    End Sub

    Private Sub GetTables()

        Dim r As Integer
        Dim strTables() As String


        Try

            DA.Init()
            If DA.Connected() Then
                strTables = DA.GetTables

                With lstTables

                    .Items.Clear()
                    For r = 0 To strTables.GetUpperBound(0)
                        .Items.Add(strTables(r))
                    Next

                End With

            End If


        Catch ex As Exception

            MsgBox("Error in GetTables: " & ex.Message)

        End Try

    End Sub

    Private Sub GetViews()

        Try

            If DA.Connected() Then

                With lstViews

                    .Items.Clear()
                    .Items.AddRange(DA.GetViews.ToArray)

                End With

            End If


        Catch ex As Exception

            MsgBox("Error in GetViews: " & ex.Message)

        End Try

    End Sub

End Class
