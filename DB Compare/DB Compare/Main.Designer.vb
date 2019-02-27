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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.tpTablesFields = New System.Windows.Forms.TabPage()
        Me.SplitTables = New System.Windows.Forms.SplitContainer()
        Me.ImagesTree = New System.Windows.Forms.ImageList(Me.components)
        Me.tcTableResults = New System.Windows.Forms.TabControl()
        Me.tabSearchOptions = New System.Windows.Forms.TabControl()
        Me.tpStoredProcedures = New System.Windows.Forms.TabPage()
        Me.lblSPs = New System.Windows.Forms.Label()
        Me.lstSPs = New System.Windows.Forms.ListBox()
        Me.tpViews = New System.Windows.Forms.TabPage()
        Me.lblViews = New System.Windows.Forms.Label()
        Me.lstViews = New System.Windows.Forms.ListBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.grpSource = New System.Windows.Forms.GroupBox()
        Me.lblPasswordSource = New System.Windows.Forms.Label()
        Me.txtPasswordSource = New System.Windows.Forms.TextBox()
        Me.lblUserSource = New System.Windows.Forms.Label()
        Me.txtUserSource = New System.Windows.Forms.TextBox()
        Me.cboDatabaseSource = New System.Windows.Forms.ComboBox()
        Me.cboServerSource = New System.Windows.Forms.ComboBox()
        Me.lblDatabaseSource = New System.Windows.Forms.Label()
        Me.lblServerSource = New System.Windows.Forms.Label()
        Me.grpTarget = New System.Windows.Forms.GroupBox()
        Me.lblPasswordTarget = New System.Windows.Forms.Label()
        Me.txtPasswordTarget = New System.Windows.Forms.TextBox()
        Me.lblUserTarget = New System.Windows.Forms.Label()
        Me.txtUserTarget = New System.Windows.Forms.TextBox()
        Me.cboDatabaseTarget = New System.Windows.Forms.ComboBox()
        Me.cboServerTarget = New System.Windows.Forms.ComboBox()
        Me.lblDatabaseTarget = New System.Windows.Forms.Label()
        Me.lblServerTarget = New System.Windows.Forms.Label()
        Me.btnDiff = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.treSource = New System.Windows.Forms.TreeView()
        Me.txtFilterTable = New System.Windows.Forms.TextBox()
        Me.btnAllSelect = New System.Windows.Forms.Button()
        Me.btnClearAll = New System.Windows.Forms.Button()
        Me.tpTablesFields.SuspendLayout()
        CType(Me.SplitTables, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitTables.Panel1.SuspendLayout()
        Me.SplitTables.Panel2.SuspendLayout()
        Me.SplitTables.SuspendLayout()
        Me.tabSearchOptions.SuspendLayout()
        Me.tpStoredProcedures.SuspendLayout()
        Me.tpViews.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.grpSource.SuspendLayout()
        Me.grpTarget.SuspendLayout()
        Me.SuspendLayout()
        '
        'tpTablesFields
        '
        Me.tpTablesFields.Controls.Add(Me.SplitTables)
        Me.tpTablesFields.Location = New System.Drawing.Point(4, 22)
        Me.tpTablesFields.Name = "tpTablesFields"
        Me.tpTablesFields.Padding = New System.Windows.Forms.Padding(3)
        Me.tpTablesFields.Size = New System.Drawing.Size(721, 230)
        Me.tpTablesFields.TabIndex = 0
        Me.tpTablesFields.Text = "Tables / Fields"
        Me.tpTablesFields.UseVisualStyleBackColor = True
        '
        'SplitTables
        '
        Me.SplitTables.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitTables.Location = New System.Drawing.Point(3, 3)
        Me.SplitTables.Name = "SplitTables"
        '
        'SplitTables.Panel1
        '
        Me.SplitTables.Panel1.Controls.Add(Me.txtFilterTable)
        Me.SplitTables.Panel1.Controls.Add(Me.treSource)
        '
        'SplitTables.Panel2
        '
        Me.SplitTables.Panel2.Controls.Add(Me.tcTableResults)
        Me.SplitTables.Size = New System.Drawing.Size(715, 224)
        Me.SplitTables.SplitterDistance = 324
        Me.SplitTables.TabIndex = 0
        '
        'ImagesTree
        '
        Me.ImagesTree.ImageStream = CType(resources.GetObject("ImagesTree.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImagesTree.TransparentColor = System.Drawing.Color.Transparent
        Me.ImagesTree.Images.SetKeyName(0, "Field")
        Me.ImagesTree.Images.SetKeyName(1, "Table")
        Me.ImagesTree.Images.SetKeyName(2, "FieldSelected")
        Me.ImagesTree.Images.SetKeyName(3, "TableSelected")
        '
        'tcTableResults
        '
        Me.tcTableResults.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tcTableResults.Location = New System.Drawing.Point(0, 0)
        Me.tcTableResults.Name = "tcTableResults"
        Me.tcTableResults.SelectedIndex = 0
        Me.tcTableResults.Size = New System.Drawing.Size(387, 224)
        Me.tcTableResults.TabIndex = 0
        '
        'tabSearchOptions
        '
        Me.tabSearchOptions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabSearchOptions.Controls.Add(Me.tpTablesFields)
        Me.tabSearchOptions.Controls.Add(Me.tpStoredProcedures)
        Me.tabSearchOptions.Controls.Add(Me.tpViews)
        Me.tabSearchOptions.Location = New System.Drawing.Point(21, 218)
        Me.tabSearchOptions.Name = "tabSearchOptions"
        Me.tabSearchOptions.SelectedIndex = 0
        Me.tabSearchOptions.Size = New System.Drawing.Size(729, 256)
        Me.tabSearchOptions.TabIndex = 53
        '
        'tpStoredProcedures
        '
        Me.tpStoredProcedures.Controls.Add(Me.lblSPs)
        Me.tpStoredProcedures.Controls.Add(Me.lstSPs)
        Me.tpStoredProcedures.Location = New System.Drawing.Point(4, 22)
        Me.tpStoredProcedures.Name = "tpStoredProcedures"
        Me.tpStoredProcedures.Padding = New System.Windows.Forms.Padding(3)
        Me.tpStoredProcedures.Size = New System.Drawing.Size(721, 230)
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
        Me.tpViews.Size = New System.Drawing.Size(721, 230)
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
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(762, 24)
        Me.MenuStrip1.TabIndex = 54
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'grpSource
        '
        Me.grpSource.Controls.Add(Me.lblPasswordSource)
        Me.grpSource.Controls.Add(Me.txtPasswordSource)
        Me.grpSource.Controls.Add(Me.lblUserSource)
        Me.grpSource.Controls.Add(Me.txtUserSource)
        Me.grpSource.Controls.Add(Me.cboDatabaseSource)
        Me.grpSource.Controls.Add(Me.cboServerSource)
        Me.grpSource.Controls.Add(Me.lblDatabaseSource)
        Me.grpSource.Controls.Add(Me.lblServerSource)
        Me.grpSource.Location = New System.Drawing.Point(49, 41)
        Me.grpSource.Name = "grpSource"
        Me.grpSource.Size = New System.Drawing.Size(485, 68)
        Me.grpSource.TabIndex = 55
        Me.grpSource.TabStop = False
        Me.grpSource.Text = "Source"
        '
        'lblPasswordSource
        '
        Me.lblPasswordSource.AutoSize = True
        Me.lblPasswordSource.Location = New System.Drawing.Point(366, 17)
        Me.lblPasswordSource.Name = "lblPasswordSource"
        Me.lblPasswordSource.Size = New System.Drawing.Size(53, 13)
        Me.lblPasswordSource.TabIndex = 59
        Me.lblPasswordSource.Text = "Password"
        '
        'txtPasswordSource
        '
        Me.txtPasswordSource.Location = New System.Drawing.Point(369, 33)
        Me.txtPasswordSource.Name = "txtPasswordSource"
        Me.txtPasswordSource.PasswordChar = Global.Microsoft.VisualBasic.ChrW(164)
        Me.txtPasswordSource.Size = New System.Drawing.Size(99, 20)
        Me.txtPasswordSource.TabIndex = 58
        '
        'lblUserSource
        '
        Me.lblUserSource.AutoSize = True
        Me.lblUserSource.Location = New System.Drawing.Point(253, 17)
        Me.lblUserSource.Name = "lblUserSource"
        Me.lblUserSource.Size = New System.Drawing.Size(29, 13)
        Me.lblUserSource.TabIndex = 57
        Me.lblUserSource.Text = "User"
        '
        'txtUserSource
        '
        Me.txtUserSource.Location = New System.Drawing.Point(256, 33)
        Me.txtUserSource.Name = "txtUserSource"
        Me.txtUserSource.Size = New System.Drawing.Size(98, 20)
        Me.txtUserSource.TabIndex = 56
        '
        'cboDatabaseSource
        '
        Me.cboDatabaseSource.FormattingEnabled = True
        Me.cboDatabaseSource.Location = New System.Drawing.Point(134, 32)
        Me.cboDatabaseSource.Name = "cboDatabaseSource"
        Me.cboDatabaseSource.Size = New System.Drawing.Size(99, 21)
        Me.cboDatabaseSource.TabIndex = 55
        '
        'cboServerSource
        '
        Me.cboServerSource.FormattingEnabled = True
        Me.cboServerSource.Location = New System.Drawing.Point(21, 32)
        Me.cboServerSource.Name = "cboServerSource"
        Me.cboServerSource.Size = New System.Drawing.Size(98, 21)
        Me.cboServerSource.TabIndex = 54
        '
        'lblDatabaseSource
        '
        Me.lblDatabaseSource.AutoSize = True
        Me.lblDatabaseSource.Location = New System.Drawing.Point(131, 15)
        Me.lblDatabaseSource.Name = "lblDatabaseSource"
        Me.lblDatabaseSource.Size = New System.Drawing.Size(53, 13)
        Me.lblDatabaseSource.TabIndex = 53
        Me.lblDatabaseSource.Text = "Database"
        '
        'lblServerSource
        '
        Me.lblServerSource.AutoSize = True
        Me.lblServerSource.Location = New System.Drawing.Point(17, 16)
        Me.lblServerSource.Name = "lblServerSource"
        Me.lblServerSource.Size = New System.Drawing.Size(38, 13)
        Me.lblServerSource.TabIndex = 52
        Me.lblServerSource.Text = "Server"
        '
        'grpTarget
        '
        Me.grpTarget.Controls.Add(Me.lblPasswordTarget)
        Me.grpTarget.Controls.Add(Me.txtPasswordTarget)
        Me.grpTarget.Controls.Add(Me.lblUserTarget)
        Me.grpTarget.Controls.Add(Me.txtUserTarget)
        Me.grpTarget.Controls.Add(Me.cboDatabaseTarget)
        Me.grpTarget.Controls.Add(Me.cboServerTarget)
        Me.grpTarget.Controls.Add(Me.lblDatabaseTarget)
        Me.grpTarget.Controls.Add(Me.lblServerTarget)
        Me.grpTarget.Location = New System.Drawing.Point(49, 122)
        Me.grpTarget.Name = "grpTarget"
        Me.grpTarget.Size = New System.Drawing.Size(485, 68)
        Me.grpTarget.TabIndex = 56
        Me.grpTarget.TabStop = False
        Me.grpTarget.Text = "Target"
        '
        'lblPasswordTarget
        '
        Me.lblPasswordTarget.AutoSize = True
        Me.lblPasswordTarget.Location = New System.Drawing.Point(366, 17)
        Me.lblPasswordTarget.Name = "lblPasswordTarget"
        Me.lblPasswordTarget.Size = New System.Drawing.Size(53, 13)
        Me.lblPasswordTarget.TabIndex = 59
        Me.lblPasswordTarget.Text = "Password"
        '
        'txtPasswordTarget
        '
        Me.txtPasswordTarget.Location = New System.Drawing.Point(369, 33)
        Me.txtPasswordTarget.Name = "txtPasswordTarget"
        Me.txtPasswordTarget.PasswordChar = Global.Microsoft.VisualBasic.ChrW(164)
        Me.txtPasswordTarget.Size = New System.Drawing.Size(99, 20)
        Me.txtPasswordTarget.TabIndex = 58
        '
        'lblUserTarget
        '
        Me.lblUserTarget.AutoSize = True
        Me.lblUserTarget.Location = New System.Drawing.Point(253, 17)
        Me.lblUserTarget.Name = "lblUserTarget"
        Me.lblUserTarget.Size = New System.Drawing.Size(29, 13)
        Me.lblUserTarget.TabIndex = 57
        Me.lblUserTarget.Text = "User"
        '
        'txtUserTarget
        '
        Me.txtUserTarget.Location = New System.Drawing.Point(256, 33)
        Me.txtUserTarget.Name = "txtUserTarget"
        Me.txtUserTarget.Size = New System.Drawing.Size(98, 20)
        Me.txtUserTarget.TabIndex = 56
        '
        'cboDatabaseTarget
        '
        Me.cboDatabaseTarget.FormattingEnabled = True
        Me.cboDatabaseTarget.Location = New System.Drawing.Point(134, 32)
        Me.cboDatabaseTarget.Name = "cboDatabaseTarget"
        Me.cboDatabaseTarget.Size = New System.Drawing.Size(99, 21)
        Me.cboDatabaseTarget.TabIndex = 55
        '
        'cboServerTarget
        '
        Me.cboServerTarget.FormattingEnabled = True
        Me.cboServerTarget.Location = New System.Drawing.Point(21, 32)
        Me.cboServerTarget.Name = "cboServerTarget"
        Me.cboServerTarget.Size = New System.Drawing.Size(98, 21)
        Me.cboServerTarget.TabIndex = 54
        '
        'lblDatabaseTarget
        '
        Me.lblDatabaseTarget.AutoSize = True
        Me.lblDatabaseTarget.Location = New System.Drawing.Point(131, 15)
        Me.lblDatabaseTarget.Name = "lblDatabaseTarget"
        Me.lblDatabaseTarget.Size = New System.Drawing.Size(53, 13)
        Me.lblDatabaseTarget.TabIndex = 53
        Me.lblDatabaseTarget.Text = "Database"
        '
        'lblServerTarget
        '
        Me.lblServerTarget.AutoSize = True
        Me.lblServerTarget.Location = New System.Drawing.Point(17, 16)
        Me.lblServerTarget.Name = "lblServerTarget"
        Me.lblServerTarget.Size = New System.Drawing.Size(38, 13)
        Me.lblServerTarget.TabIndex = 52
        Me.lblServerTarget.Text = "Server"
        '
        'btnDiff
        '
        Me.btnDiff.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDiff.Location = New System.Drawing.Point(587, 486)
        Me.btnDiff.Name = "btnDiff"
        Me.btnDiff.Size = New System.Drawing.Size(75, 23)
        Me.btnDiff.TabIndex = 58
        Me.btnDiff.Text = "&Diff"
        Me.btnDiff.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(668, 486)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 59
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(568, 72)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(75, 23)
        Me.btnRefresh.TabIndex = 60
        Me.btnRefresh.Text = "&Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'treSource
        '
        Me.treSource.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.treSource.BackColor = System.Drawing.SystemColors.Window
        Me.treSource.CheckBoxes = True
        Me.treSource.ImageIndex = 0
        Me.treSource.ImageList = Me.ImagesTree
        Me.treSource.Location = New System.Drawing.Point(0, 29)
        Me.treSource.Name = "treSource"
        Me.treSource.SelectedImageIndex = 0
        Me.treSource.Size = New System.Drawing.Size(321, 195)
        Me.treSource.TabIndex = 1
        '
        'txtFilterTable
        '
        Me.txtFilterTable.Location = New System.Drawing.Point(3, 3)
        Me.txtFilterTable.Name = "txtFilterTable"
        Me.txtFilterTable.Size = New System.Drawing.Size(318, 20)
        Me.txtFilterTable.TabIndex = 2
        '
        'btnAllSelect
        '
        Me.btnAllSelect.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAllSelect.Location = New System.Drawing.Point(25, 480)
        Me.btnAllSelect.Name = "btnAllSelect"
        Me.btnAllSelect.Size = New System.Drawing.Size(75, 23)
        Me.btnAllSelect.TabIndex = 61
        Me.btnAllSelect.Text = "&Select All"
        Me.btnAllSelect.UseVisualStyleBackColor = True
        '
        'btnClearAll
        '
        Me.btnClearAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnClearAll.Location = New System.Drawing.Point(106, 480)
        Me.btnClearAll.Name = "btnClearAll"
        Me.btnClearAll.Size = New System.Drawing.Size(75, 23)
        Me.btnClearAll.TabIndex = 62
        Me.btnClearAll.Text = "&Clear All"
        Me.btnClearAll.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(762, 521)
        Me.Controls.Add(Me.btnClearAll)
        Me.Controls.Add(Me.btnAllSelect)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnDiff)
        Me.Controls.Add(Me.grpTarget)
        Me.Controls.Add(Me.grpSource)
        Me.Controls.Add(Me.tabSearchOptions)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmMain"
        Me.Text = "DB Compare"
        Me.tpTablesFields.ResumeLayout(False)
        Me.SplitTables.Panel1.ResumeLayout(False)
        Me.SplitTables.Panel1.PerformLayout()
        Me.SplitTables.Panel2.ResumeLayout(False)
        CType(Me.SplitTables, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitTables.ResumeLayout(False)
        Me.tabSearchOptions.ResumeLayout(False)
        Me.tpStoredProcedures.ResumeLayout(False)
        Me.tpStoredProcedures.PerformLayout()
        Me.tpViews.ResumeLayout(False)
        Me.tpViews.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.grpSource.ResumeLayout(False)
        Me.grpSource.PerformLayout()
        Me.grpTarget.ResumeLayout(False)
        Me.grpTarget.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tpTablesFields As System.Windows.Forms.TabPage
    Friend WithEvents tabSearchOptions As System.Windows.Forms.TabControl
    Friend WithEvents tpStoredProcedures As System.Windows.Forms.TabPage
    Friend WithEvents lblSPs As System.Windows.Forms.Label
    Friend WithEvents lstSPs As System.Windows.Forms.ListBox
    Friend WithEvents tpViews As System.Windows.Forms.TabPage
    Friend WithEvents lblViews As System.Windows.Forms.Label
    Friend WithEvents lstViews As System.Windows.Forms.ListBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents grpSource As System.Windows.Forms.GroupBox
    Friend WithEvents lblPasswordSource As System.Windows.Forms.Label
    Friend WithEvents txtPasswordSource As System.Windows.Forms.TextBox
    Friend WithEvents lblUserSource As System.Windows.Forms.Label
    Friend WithEvents txtUserSource As System.Windows.Forms.TextBox
    Friend WithEvents cboDatabaseSource As System.Windows.Forms.ComboBox
    Friend WithEvents cboServerSource As System.Windows.Forms.ComboBox
    Friend WithEvents lblDatabaseSource As System.Windows.Forms.Label
    Friend WithEvents lblServerSource As System.Windows.Forms.Label
    Friend WithEvents grpTarget As System.Windows.Forms.GroupBox
    Friend WithEvents lblPasswordTarget As System.Windows.Forms.Label
    Friend WithEvents txtPasswordTarget As System.Windows.Forms.TextBox
    Friend WithEvents lblUserTarget As System.Windows.Forms.Label
    Friend WithEvents txtUserTarget As System.Windows.Forms.TextBox
    Friend WithEvents cboDatabaseTarget As System.Windows.Forms.ComboBox
    Friend WithEvents cboServerTarget As System.Windows.Forms.ComboBox
    Friend WithEvents lblDatabaseTarget As System.Windows.Forms.Label
    Friend WithEvents lblServerTarget As System.Windows.Forms.Label
    Friend WithEvents SplitTables As System.Windows.Forms.SplitContainer
    Friend WithEvents btnDiff As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents ImagesTree As System.Windows.Forms.ImageList
    Friend WithEvents tcTableResults As System.Windows.Forms.TabControl
    Friend WithEvents treSource As System.Windows.Forms.TreeView
    Friend WithEvents txtFilterTable As System.Windows.Forms.TextBox
    Friend WithEvents btnAllSelect As System.Windows.Forms.Button
    Friend WithEvents btnClearAll As System.Windows.Forms.Button

End Class
