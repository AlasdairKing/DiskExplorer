<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents mnuFileRunOrOpen As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCommandsDelete As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCommandsOpeninexplorer As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCommandsCreatefolder As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFileExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuNavigateBack As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuNavigateUp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuNavigateDesktop As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuNavigate As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuViewFolders As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuViewFiles As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuViewAll As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuBar1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuViewTouchscreen As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuViewFontsize_0 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuViewFontsize_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuView As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOptionsStartupwithcomputer As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOptions As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelpManual As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelpAbout As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents tmrUpdateIcons As System.Windows.Forms.Timer
    Public WithEvents cmdUp As System.Windows.Forms.Button
	Public WithEvents cmdBack As System.Windows.Forms.Button
	Public WithEvents cmdGo As System.Windows.Forms.Button
	Public WithEvents _staMain_Panel1 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents staMain As System.Windows.Forms.StatusStrip
    Public WithEvents lblFiles As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileRunOrOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCommandsDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCommandsOpeninexplorer = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCommandsCreatefolder = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNavigate = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNavigateBack = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNavigateUp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNavigateDesktop = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuView = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuViewFolders = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuViewFiles = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuViewAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuBar1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuViewTouchscreen = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuViewFontsize_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuViewFontsize_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOptions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOptionsStartupwithcomputer = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpManual = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.tmrUpdateIcons = New System.Windows.Forms.Timer(Me.components)
        Me.cmdUp = New System.Windows.Forms.Button()
        Me.cmdBack = New System.Windows.Forms.Button()
        Me.cmdGo = New System.Windows.Forms.Button()
        Me.staMain = New System.Windows.Forms.StatusStrip()
        Me._staMain_Panel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblFiles = New System.Windows.Forms.Label()
        Me.PanelMain = New System.Windows.Forms.Panel()
        Me.lstDisplay = New System.Windows.Forms.ListBox()
        Me.lstIcons = New System.Windows.Forms.ListBox()
        Me.MainMenu1.SuspendLayout()
        Me.staMain.SuspendLayout()
        Me.PanelMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuNavigate, Me.mnuView, Me.mnuOptions, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(390, 24)
        Me.MainMenu1.TabIndex = 7
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileRunOrOpen, Me.mnuCommandsDelete, Me.mnuCommandsOpeninexplorer, Me.mnuCommandsCreatefolder, Me.mnuFileExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuFileRunOrOpen
        '
        Me.mnuFileRunOrOpen.Name = "mnuFileRunOrOpen"
        Me.mnuFileRunOrOpen.ShortcutKeyDisplayString = "Return"
        Me.mnuFileRunOrOpen.Size = New System.Drawing.Size(242, 22)
        Me.mnuFileRunOrOpen.Text = "&Run or Open"
        '
        'mnuCommandsDelete
        '
        Me.mnuCommandsDelete.Name = "mnuCommandsDelete"
        Me.mnuCommandsDelete.ShortcutKeys = System.Windows.Forms.Keys.Delete
        Me.mnuCommandsDelete.Size = New System.Drawing.Size(242, 22)
        Me.mnuCommandsDelete.Text = "&Delete file or folder"
        '
        'mnuCommandsOpeninexplorer
        '
        Me.mnuCommandsOpeninexplorer.Name = "mnuCommandsOpeninexplorer"
        Me.mnuCommandsOpeninexplorer.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
        Me.mnuCommandsOpeninexplorer.Size = New System.Drawing.Size(242, 22)
        Me.mnuCommandsOpeninexplorer.Text = "Open in &Explorer"
        '
        'mnuCommandsCreatefolder
        '
        Me.mnuCommandsCreatefolder.Name = "mnuCommandsCreatefolder"
        Me.mnuCommandsCreatefolder.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D), System.Windows.Forms.Keys)
        Me.mnuCommandsCreatefolder.Size = New System.Drawing.Size(242, 22)
        Me.mnuCommandsCreatefolder.Text = "&Create folder (directory)"
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Name = "mnuFileExit"
        Me.mnuFileExit.Size = New System.Drawing.Size(242, 22)
        Me.mnuFileExit.Text = "E&xit"
        '
        'mnuNavigate
        '
        Me.mnuNavigate.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNavigateBack, Me.mnuNavigateUp, Me.mnuNavigateDesktop})
        Me.mnuNavigate.Name = "mnuNavigate"
        Me.mnuNavigate.Size = New System.Drawing.Size(66, 20)
        Me.mnuNavigate.Text = "&Navigate"
        '
        'mnuNavigateBack
        '
        Me.mnuNavigateBack.Name = "mnuNavigateBack"
        Me.mnuNavigateBack.ShortcutKeyDisplayString = "Backspace"
        Me.mnuNavigateBack.Size = New System.Drawing.Size(201, 22)
        Me.mnuNavigateBack.Text = "&Back"
        '
        'mnuNavigateUp
        '
        Me.mnuNavigateUp.Name = "mnuNavigateUp"
        Me.mnuNavigateUp.ShortcutKeyDisplayString = "Ctrl+Up"
        Me.mnuNavigateUp.Size = New System.Drawing.Size(201, 22)
        Me.mnuNavigateUp.Text = "&Up"
        '
        'mnuNavigateDesktop
        '
        Me.mnuNavigateDesktop.Name = "mnuNavigateDesktop"
        Me.mnuNavigateDesktop.ShortcutKeyDisplayString = "Home"
        Me.mnuNavigateDesktop.Size = New System.Drawing.Size(201, 22)
        Me.mnuNavigateDesktop.Text = "&Desktop (Home)"
        '
        'mnuView
        '
        Me.mnuView.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuViewFolders, Me.mnuViewFiles, Me.mnuViewAll, Me.mnuBar1, Me.mnuViewTouchscreen, Me._mnuViewFontsize_0, Me._mnuViewFontsize_1})
        Me.mnuView.Name = "mnuView"
        Me.mnuView.Size = New System.Drawing.Size(44, 20)
        Me.mnuView.Text = "&View"
        '
        'mnuViewFolders
        '
        Me.mnuViewFolders.Name = "mnuViewFolders"
        Me.mnuViewFolders.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.mnuViewFolders.Size = New System.Drawing.Size(235, 22)
        Me.mnuViewFolders.Text = "F&olders only"
        '
        'mnuViewFiles
        '
        Me.mnuViewFiles.Name = "mnuViewFiles"
        Me.mnuViewFiles.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.I), System.Windows.Forms.Keys)
        Me.mnuViewFiles.Size = New System.Drawing.Size(235, 22)
        Me.mnuViewFiles.Text = "F&iles only"
        '
        'mnuViewAll
        '
        Me.mnuViewAll.Name = "mnuViewAll"
        Me.mnuViewAll.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.B), System.Windows.Forms.Keys)
        Me.mnuViewAll.Size = New System.Drawing.Size(235, 22)
        Me.mnuViewAll.Text = "&Both files and folders"
        '
        'mnuBar1
        '
        Me.mnuBar1.Name = "mnuBar1"
        Me.mnuBar1.Size = New System.Drawing.Size(232, 6)
        '
        'mnuViewTouchscreen
        '
        Me.mnuViewTouchscreen.Name = "mnuViewTouchscreen"
        Me.mnuViewTouchscreen.Size = New System.Drawing.Size(235, 22)
        Me.mnuViewTouchscreen.Text = "&Touchscreen"
        Me.mnuViewTouchscreen.Visible = False
        '
        'mnuOptions
        '
        Me.mnuOptions.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOptionsStartupwithcomputer})
        Me.mnuOptions.Name = "mnuOptions"
        Me.mnuOptions.Size = New System.Drawing.Size(61, 20)
        Me.mnuOptions.Text = "&Options"
        '
        'mnuOptionsStartupwithcomputer
        '
        Me.mnuOptionsStartupwithcomputer.Name = "mnuOptionsStartupwithcomputer"
        Me.mnuOptionsStartupwithcomputer.Size = New System.Drawing.Size(193, 22)
        Me.mnuOptionsStartupwithcomputer.Text = "&Startup with computer"
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuHelpManual, Me.mnuHelpAbout})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpManual
        '
        Me.mnuHelpManual.Name = "mnuHelpManual"
        Me.mnuHelpManual.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.mnuHelpManual.Size = New System.Drawing.Size(133, 22)
        Me.mnuHelpManual.Text = "&Manual"
        '
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Name = "mnuHelpAbout"
        Me.mnuHelpAbout.Size = New System.Drawing.Size(133, 22)
        Me.mnuHelpAbout.Text = "&About"
        '
        'tmrUpdateIcons
        '
        Me.tmrUpdateIcons.Enabled = True
        Me.tmrUpdateIcons.Interval = 50
        '
        'cmdUp
        '
        Me.cmdUp.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUp.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUp.Location = New System.Drawing.Point(150, 140)
        Me.cmdUp.Name = "cmdUp"
        Me.cmdUp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUp.Size = New System.Drawing.Size(102, 42)
        Me.cmdUp.TabIndex = 5
        Me.cmdUp.Text = "Up"
        Me.cmdUp.UseVisualStyleBackColor = False
        '
        'cmdBack
        '
        Me.cmdBack.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBack.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBack.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBack.Location = New System.Drawing.Point(230, 190)
        Me.cmdBack.Name = "cmdBack"
        Me.cmdBack.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBack.Size = New System.Drawing.Size(102, 42)
        Me.cmdBack.TabIndex = 4
        Me.cmdBack.Text = "Back"
        Me.cmdBack.UseVisualStyleBackColor = False
        '
        'cmdGo
        '
        Me.cmdGo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdGo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdGo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdGo.Location = New System.Drawing.Point(150, 140)
        Me.cmdGo.Name = "cmdGo"
        Me.cmdGo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdGo.Size = New System.Drawing.Size(102, 42)
        Me.cmdGo.TabIndex = 3
        Me.cmdGo.Text = "Go"
        Me.cmdGo.UseVisualStyleBackColor = False
        '
        'staMain
        '
        Me.staMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._staMain_Panel1})
        Me.staMain.Location = New System.Drawing.Point(0, 250)
        Me.staMain.Name = "staMain"
        Me.staMain.Size = New System.Drawing.Size(390, 32)
        Me.staMain.TabIndex = 2
        '
        '_staMain_Panel1
        '
        Me._staMain_Panel1.AutoSize = False
        Me._staMain_Panel1.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me._staMain_Panel1.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me._staMain_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me._staMain_Panel1.Name = "_staMain_Panel1"
        Me._staMain_Panel1.Size = New System.Drawing.Size(120, 32)
        Me._staMain_Panel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFiles
        '
        Me.lblFiles.AutoSize = True
        Me.lblFiles.BackColor = System.Drawing.SystemColors.Control
        Me.lblFiles.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFiles.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFiles.Location = New System.Drawing.Point(0, 30)
        Me.lblFiles.Name = "lblFiles"
        Me.lblFiles.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFiles.Size = New System.Drawing.Size(35, 16)
        Me.lblFiles.TabIndex = 0
        Me.lblFiles.Text = "Files"
        '
        'mnuViewFontsize
        '
        '
        'PanelMain
        '
        Me.PanelMain.Controls.Add(Me.lstDisplay)
        Me.PanelMain.Controls.Add(Me.lstIcons)
        Me.PanelMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelMain.Location = New System.Drawing.Point(0, 24)
        Me.PanelMain.Name = "PanelMain"
        Me.PanelMain.Size = New System.Drawing.Size(390, 226)
        Me.PanelMain.TabIndex = 11
        '
        'lstDisplay
        '
        Me.lstDisplay.BackColor = System.Drawing.SystemColors.Window
        Me.lstDisplay.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstDisplay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstDisplay.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstDisplay.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstDisplay.IntegralHeight = False
        Me.lstDisplay.ItemHeight = 19
        Me.lstDisplay.Location = New System.Drawing.Point(47, 0)
        Me.lstDisplay.Name = "lstDisplay"
        Me.lstDisplay.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstDisplay.Size = New System.Drawing.Size(343, 226)
        Me.lstDisplay.Sorted = True
        Me.lstDisplay.TabIndex = 12
        '
        'lstIcons
        '
        Me.lstIcons.BackColor = System.Drawing.SystemColors.Window
        Me.lstIcons.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstIcons.Dock = System.Windows.Forms.DockStyle.Left
        Me.lstIcons.Enabled = False
        Me.lstIcons.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstIcons.IntegralHeight = False
        Me.lstIcons.ItemHeight = 16
        Me.lstIcons.Location = New System.Drawing.Point(0, 0)
        Me.lstIcons.Name = "lstIcons"
        Me.lstIcons.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstIcons.Size = New System.Drawing.Size(47, 226)
        Me.lstIcons.TabIndex = 11
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(390, 282)
        Me.Controls.Add(Me.PanelMain)
        Me.Controls.Add(Me.cmdUp)
        Me.Controls.Add(Me.cmdBack)
        Me.Controls.Add(Me.cmdGo)
        Me.Controls.Add(Me.staMain)
        Me.Controls.Add(Me.lblFiles)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(19, 68)
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Explorer"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.staMain.ResumeLayout(False)
        Me.staMain.PerformLayout()
        Me.PanelMain.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PanelMain As System.Windows.Forms.Panel
    Public WithEvents lstDisplay As System.Windows.Forms.ListBox
    Public WithEvents lstIcons As System.Windows.Forms.ListBox
#End Region
End Class