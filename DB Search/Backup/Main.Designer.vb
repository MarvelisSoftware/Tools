<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.lblPassword = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.lblUID = New System.Windows.Forms.Label
        Me.txtUser = New System.Windows.Forms.TextBox
        Me.cboDatabase = New System.Windows.Forms.ComboBox
        Me.cboServer = New System.Windows.Forms.ComboBox
        Me.lblDatabase = New System.Windows.Forms.Label
        Me.lblServer = New System.Windows.Forms.Label
        Me.btnConnect = New System.Windows.Forms.Button
        Me.btnSearch = New System.Windows.Forms.Button
        Me.lblString = New System.Windows.Forms.Label
        Me.txtString = New System.Windows.Forms.TextBox
        Me.btnSearchData = New System.Windows.Forms.Button
        Me.btnSearchSP = New System.Windows.Forms.Button
        Me.btnSearchView = New System.Windows.Forms.Button
        Me.tabSearchOptions = New System.Windows.Forms.TabControl
        Me.tpTablesFields = New System.Windows.Forms.TabPage
        Me.lblFields = New System.Windows.Forms.Label
        Me.lstFields = New System.Windows.Forms.ListBox
        Me.lblTables = New System.Windows.Forms.Label
        Me.lstTables = New System.Windows.Forms.ListBox
        Me.tpStoredProcedures = New System.Windows.Forms.TabPage
        Me.lblSPs = New System.Windows.Forms.Label
        Me.lstSPs = New System.Windows.Forms.ListBox
        Me.tpViews = New System.Windows.Forms.TabPage
        Me.lblViews = New System.Windows.Forms.Label
        Me.lstViews = New System.Windows.Forms.ListBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.lstResults = New System.Windows.Forms.ListBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.txtResults = New System.Windows.Forms.TextBox
        Me.mnuMain = New System.Windows.Forms.MenuStrip
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.tabSearchOptions.SuspendLayout()
        Me.tpTablesFields.SuspendLayout()
        Me.tpStoredProcedures.SuspendLayout()
        Me.tpViews.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.mnuMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Location = New System.Drawing.Point(368, 34)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(53, 13)
        Me.lblPassword.TabIndex = 29
        Me.lblPassword.Text = "Password"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(371, 50)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(164)
        Me.txtPassword.Size = New System.Drawing.Size(99, 20)
        Me.txtPassword.TabIndex = 28
        '
        'lblUID
        '
        Me.lblUID.AutoSize = True
        Me.lblUID.Location = New System.Drawing.Point(255, 34)
        Me.lblUID.Name = "lblUID"
        Me.lblUID.Size = New System.Drawing.Size(29, 13)
        Me.lblUID.TabIndex = 27
        Me.lblUID.Text = "User"
        '
        'txtUser
        '
        Me.txtUser.Location = New System.Drawing.Point(258, 50)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.Size = New System.Drawing.Size(98, 20)
        Me.txtUser.TabIndex = 26
        '
        'cboDatabase
        '
        Me.cboDatabase.FormattingEnabled = True
        Me.cboDatabase.Location = New System.Drawing.Point(136, 49)
        Me.cboDatabase.Name = "cboDatabase"
        Me.cboDatabase.Size = New System.Drawing.Size(99, 21)
        Me.cboDatabase.TabIndex = 25
        '
        'cboServer
        '
        Me.cboServer.FormattingEnabled = True
        Me.cboServer.Location = New System.Drawing.Point(23, 49)
        Me.cboServer.Name = "cboServer"
        Me.cboServer.Size = New System.Drawing.Size(98, 21)
        Me.cboServer.TabIndex = 24
        '
        'lblDatabase
        '
        Me.lblDatabase.AutoSize = True
        Me.lblDatabase.Location = New System.Drawing.Point(133, 32)
        Me.lblDatabase.Name = "lblDatabase"
        Me.lblDatabase.Size = New System.Drawing.Size(53, 13)
        Me.lblDatabase.TabIndex = 23
        Me.lblDatabase.Text = "Database"
        '
        'lblServer
        '
        Me.lblServer.AutoSize = True
        Me.lblServer.Location = New System.Drawing.Point(19, 33)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(38, 13)
        Me.lblServer.TabIndex = 22
        Me.lblServer.Text = "Server"
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(492, 43)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(92, 27)
        Me.btnConnect.TabIndex = 30
        Me.btnConnect.Text = "Connect"
        Me.btnConnect.UseVisualStyleBackColor = True
        '
        'btnSearch
        '
        Me.btnSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSearch.Location = New System.Drawing.Point(616, 43)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(92, 27)
        Me.btnSearch.TabIndex = 37
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'lblString
        '
        Me.lblString.AutoSize = True
        Me.lblString.Location = New System.Drawing.Point(368, 97)
        Me.lblString.Name = "lblString"
        Me.lblString.Size = New System.Drawing.Size(34, 13)
        Me.lblString.TabIndex = 36
        Me.lblString.Text = "String"
        '
        'txtString
        '
        Me.txtString.Location = New System.Drawing.Point(371, 113)
        Me.txtString.Name = "txtString"
        Me.txtString.Size = New System.Drawing.Size(99, 20)
        Me.txtString.TabIndex = 35
        '
        'btnSearchData
        '
        Me.btnSearchData.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSearchData.Location = New System.Drawing.Point(616, 76)
        Me.btnSearchData.Name = "btnSearchData"
        Me.btnSearchData.Size = New System.Drawing.Size(92, 27)
        Me.btnSearchData.TabIndex = 40
        Me.btnSearchData.Text = "Search Data"
        Me.btnSearchData.UseVisualStyleBackColor = True
        '
        'btnSearchSP
        '
        Me.btnSearchSP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSearchSP.Location = New System.Drawing.Point(616, 109)
        Me.btnSearchSP.Name = "btnSearchSP"
        Me.btnSearchSP.Size = New System.Drawing.Size(92, 27)
        Me.btnSearchSP.TabIndex = 41
        Me.btnSearchSP.Text = "Search SP"
        Me.btnSearchSP.UseVisualStyleBackColor = True
        '
        'btnSearchView
        '
        Me.btnSearchView.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSearchView.Location = New System.Drawing.Point(616, 142)
        Me.btnSearchView.Name = "btnSearchView"
        Me.btnSearchView.Size = New System.Drawing.Size(92, 27)
        Me.btnSearchView.TabIndex = 42
        Me.btnSearchView.Text = "Search View"
        Me.btnSearchView.UseVisualStyleBackColor = True
        '
        'tabSearchOptions
        '
        Me.tabSearchOptions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tabSearchOptions.Controls.Add(Me.tpTablesFields)
        Me.tabSearchOptions.Controls.Add(Me.tpStoredProcedures)
        Me.tabSearchOptions.Controls.Add(Me.tpViews)
        Me.tabSearchOptions.Location = New System.Drawing.Point(12, 97)
        Me.tabSearchOptions.Name = "tabSearchOptions"
        Me.tabSearchOptions.SelectedIndex = 0
        Me.tabSearchOptions.Size = New System.Drawing.Size(344, 291)
        Me.tabSearchOptions.TabIndex = 43
        '
        'tpTablesFields
        '
        Me.tpTablesFields.Controls.Add(Me.lblFields)
        Me.tpTablesFields.Controls.Add(Me.lstFields)
        Me.tpTablesFields.Controls.Add(Me.lblTables)
        Me.tpTablesFields.Controls.Add(Me.lstTables)
        Me.tpTablesFields.Location = New System.Drawing.Point(4, 22)
        Me.tpTablesFields.Name = "tpTablesFields"
        Me.tpTablesFields.Padding = New System.Windows.Forms.Padding(3)
        Me.tpTablesFields.Size = New System.Drawing.Size(336, 265)
        Me.tpTablesFields.TabIndex = 0
        Me.tpTablesFields.Text = "Tables / Fields"
        Me.tpTablesFields.UseVisualStyleBackColor = True
        '
        'lblFields
        '
        Me.lblFields.AutoSize = True
        Me.lblFields.Location = New System.Drawing.Point(161, 25)
        Me.lblFields.Name = "lblFields"
        Me.lblFields.Size = New System.Drawing.Size(37, 13)
        Me.lblFields.TabIndex = 38
        Me.lblFields.Text = "Ffields"
        '
        'lstFields
        '
        Me.lstFields.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstFields.FormattingEnabled = True
        Me.lstFields.Location = New System.Drawing.Point(164, 41)
        Me.lstFields.Name = "lstFields"
        Me.lstFields.Size = New System.Drawing.Size(155, 199)
        Me.lstFields.TabIndex = 37
        '
        'lblTables
        '
        Me.lblTables.AutoSize = True
        Me.lblTables.Location = New System.Drawing.Point(20, 25)
        Me.lblTables.Name = "lblTables"
        Me.lblTables.Size = New System.Drawing.Size(39, 13)
        Me.lblTables.TabIndex = 36
        Me.lblTables.Text = "Tables"
        '
        'lstTables
        '
        Me.lstTables.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstTables.FormattingEnabled = True
        Me.lstTables.Location = New System.Drawing.Point(23, 41)
        Me.lstTables.Name = "lstTables"
        Me.lstTables.Size = New System.Drawing.Size(120, 199)
        Me.lstTables.Sorted = True
        Me.lstTables.TabIndex = 35
        '
        'tpStoredProcedures
        '
        Me.tpStoredProcedures.Controls.Add(Me.lblSPs)
        Me.tpStoredProcedures.Controls.Add(Me.lstSPs)
        Me.tpStoredProcedures.Location = New System.Drawing.Point(4, 22)
        Me.tpStoredProcedures.Name = "tpStoredProcedures"
        Me.tpStoredProcedures.Padding = New System.Windows.Forms.Padding(3)
        Me.tpStoredProcedures.Size = New System.Drawing.Size(336, 265)
        Me.tpStoredProcedures.TabIndex = 1
        Me.tpStoredProcedures.Text = "Stored Procedures"
        Me.tpStoredProcedures.UseVisualStyleBackColor = True
        '
        'lblSPs
        '
        Me.lblSPs.AutoSize = True
        Me.lblSPs.Location = New System.Drawing.Point(21, 25)
        Me.lblSPs.Name = "lblSPs"
        Me.lblSPs.Size = New System.Drawing.Size(95, 13)
        Me.lblSPs.TabIndex = 40
        Me.lblSPs.Text = "Stored Procedures"
        '
        'lstSPs
        '
        Me.lstSPs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstSPs.FormattingEnabled = True
        Me.lstSPs.Location = New System.Drawing.Point(24, 41)
        Me.lstSPs.Name = "lstSPs"
        Me.lstSPs.Size = New System.Drawing.Size(285, 199)
        Me.lstSPs.Sorted = True
        Me.lstSPs.TabIndex = 39
        '
        'tpViews
        '
        Me.tpViews.Controls.Add(Me.lblViews)
        Me.tpViews.Controls.Add(Me.lstViews)
        Me.tpViews.Location = New System.Drawing.Point(4, 22)
        Me.tpViews.Name = "tpViews"
        Me.tpViews.Size = New System.Drawing.Size(336, 265)
        Me.tpViews.TabIndex = 2
        Me.tpViews.Text = "Views"
        Me.tpViews.UseVisualStyleBackColor = True
        '
        'lblViews
        '
        Me.lblViews.AutoSize = True
        Me.lblViews.Location = New System.Drawing.Point(24, 25)
        Me.lblViews.Name = "lblViews"
        Me.lblViews.Size = New System.Drawing.Size(35, 13)
        Me.lblViews.TabIndex = 42
        Me.lblViews.Text = "Views"
        '
        'lstViews
        '
        Me.lstViews.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstViews.FormattingEnabled = True
        Me.lstViews.Location = New System.Drawing.Point(27, 41)
        Me.lstViews.Name = "lstViews"
        Me.lstViews.Size = New System.Drawing.Size(285, 199)
        Me.lstViews.Sorted = True
        Me.lstViews.TabIndex = 41
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(391, 160)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(316, 227)
        Me.TabControl1.TabIndex = 44
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.lstResults)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(308, 201)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "List"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'lstResults
        '
        Me.lstResults.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstResults.FormattingEnabled = True
        Me.lstResults.Location = New System.Drawing.Point(6, 9)
        Me.lstResults.Name = "lstResults"
        Me.lstResults.Size = New System.Drawing.Size(283, 186)
        Me.lstResults.TabIndex = 39
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.txtResults)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(308, 201)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Text"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'txtResults
        '
        Me.txtResults.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtResults.Location = New System.Drawing.Point(3, 3)
        Me.txtResults.Multiline = True
        Me.txtResults.Name = "txtResults"
        Me.txtResults.Size = New System.Drawing.Size(302, 195)
        Me.txtResults.TabIndex = 0
        '
        'mnuMain
        '
        Me.mnuMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.mnuMain.Location = New System.Drawing.Point(0, 0)
        Me.mnuMain.Name = "mnuMain"
        Me.mnuMain.Size = New System.Drawing.Size(720, 24)
        Me.mnuMain.TabIndex = 45
        Me.mnuMain.Text = "MenuStrip1"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.AboutToolStripMenuItem.Text = "&About"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(720, 408)
        Me.Controls.Add(Me.btnSearchSP)
        Me.Controls.Add(Me.btnSearchView)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.tabSearchOptions)
        Me.Controls.Add(Me.btnSearchData)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.lblString)
        Me.Controls.Add(Me.txtString)
        Me.Controls.Add(Me.btnConnect)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.lblUID)
        Me.Controls.Add(Me.txtUser)
        Me.Controls.Add(Me.cboDatabase)
        Me.Controls.Add(Me.cboServer)
        Me.Controls.Add(Me.lblDatabase)
        Me.Controls.Add(Me.lblServer)
        Me.Controls.Add(Me.mnuMain)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.mnuMain
        Me.Name = "frmMain"
        Me.Text = "DB Search"
        Me.tabSearchOptions.ResumeLayout(False)
        Me.tpTablesFields.ResumeLayout(False)
        Me.tpTablesFields.PerformLayout()
        Me.tpStoredProcedures.ResumeLayout(False)
        Me.tpStoredProcedures.PerformLayout()
        Me.tpViews.ResumeLayout(False)
        Me.tpViews.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.mnuMain.ResumeLayout(False)
        Me.mnuMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents lblUID As System.Windows.Forms.Label
    Friend WithEvents txtUser As System.Windows.Forms.TextBox
    Friend WithEvents cboDatabase As System.Windows.Forms.ComboBox
    Friend WithEvents cboServer As System.Windows.Forms.ComboBox
    Friend WithEvents lblDatabase As System.Windows.Forms.Label
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents lblString As System.Windows.Forms.Label
    Friend WithEvents txtString As System.Windows.Forms.TextBox
    Friend WithEvents btnSearchData As System.Windows.Forms.Button
    Friend WithEvents btnSearchSP As System.Windows.Forms.Button
    Friend WithEvents btnSearchView As System.Windows.Forms.Button
    Friend WithEvents tabSearchOptions As System.Windows.Forms.TabControl
    Friend WithEvents tpTablesFields As System.Windows.Forms.TabPage
    Friend WithEvents lblFields As System.Windows.Forms.Label
    Friend WithEvents lstFields As System.Windows.Forms.ListBox
    Friend WithEvents lblTables As System.Windows.Forms.Label
    Friend WithEvents lstTables As System.Windows.Forms.ListBox
    Friend WithEvents tpStoredProcedures As System.Windows.Forms.TabPage
    Friend WithEvents lblSPs As System.Windows.Forms.Label
    Friend WithEvents lstSPs As System.Windows.Forms.ListBox
    Friend WithEvents tpViews As System.Windows.Forms.TabPage
    Friend WithEvents lblViews As System.Windows.Forms.Label
    Friend WithEvents lstViews As System.Windows.Forms.ListBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents lstResults As System.Windows.Forms.ListBox
    Friend WithEvents txtResults As System.Windows.Forms.TextBox
    Friend WithEvents mnuMain As System.Windows.Forms.MenuStrip
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
