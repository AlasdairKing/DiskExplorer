<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLargeFonts
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
	Public WithEvents lblSizer As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLargeFonts))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.lblSizer = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Large Fonts"
		Me.ClientSize = New System.Drawing.Size(388, 255)
		Me.Location = New System.Drawing.Point(5, 47)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
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
		Me.Name = "frmLargeFonts"
		Me.lblSizer.Size = New System.Drawing.Size(4, 18)
		Me.lblSizer.Location = New System.Drawing.Point(10, 10)
		Me.lblSizer.TabIndex = 0
		Me.lblSizer.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblSizer.BackColor = System.Drawing.SystemColors.Control
		Me.lblSizer.Enabled = True
		Me.lblSizer.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblSizer.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSizer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSizer.UseMnemonic = True
		Me.lblSizer.Visible = True
		Me.lblSizer.AutoSize = True
		Me.lblSizer.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSizer.Name = "lblSizer"
		Me.Controls.Add(lblSizer)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class