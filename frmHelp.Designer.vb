<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmHelp
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		Form_Initialize_renamed()
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
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtHelp As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmHelp))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtHelp = New System.Windows.Forms.TextBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Accessible Radio Help"
		Me.ClientSize = New System.Drawing.Size(390, 279)
		Me.Location = New System.Drawing.Point(5, 29)
		Me.Font = New System.Drawing.Font("Tahoma", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("frmHelp.Icon"), System.Drawing.Icon)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Tag = "frmHelp"
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmHelp"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdOK
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(102, 42)
		Me.cmdOK.Location = New System.Drawing.Point(140, 230)
		Me.cmdOK.TabIndex = 1
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.txtHelp.AutoSize = False
		Me.txtHelp.Size = New System.Drawing.Size(392, 222)
		Me.txtHelp.Location = New System.Drawing.Point(0, 0)
		Me.txtHelp.MultiLine = True
		Me.txtHelp.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtHelp.TabIndex = 0
		Me.txtHelp.AcceptsReturn = True
		Me.txtHelp.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtHelp.BackColor = System.Drawing.SystemColors.Window
		Me.txtHelp.CausesValidation = True
		Me.txtHelp.Enabled = True
		Me.txtHelp.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtHelp.HideSelection = True
		Me.txtHelp.ReadOnly = False
		Me.txtHelp.Maxlength = 0
		Me.txtHelp.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtHelp.TabStop = True
		Me.txtHelp.Visible = True
		Me.txtHelp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtHelp.Name = "txtHelp"
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(txtHelp)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class