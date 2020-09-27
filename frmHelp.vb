Option Strict Off
Option Explicit On
Friend Class frmHelp
	Inherits System.Windows.Forms.Form
	'Calendar
	'Copyright Alasdair King, 2010, http://www.alasdairking.me.uk
	'Released under the GNU Public Licence, Version 3.
	
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		On Error Resume Next
		Call Me.Hide()
	End Sub
	
	'UPGRADE_WARNING: Form event frmHelp.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmHelp_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error Resume Next
        txtHelp.Text = "" 'TODO modI18N.helpTopicText(0)
        'TODO Me.Text = modI18N.helpTopicTitle(0)
		Call txtHelp.Focus()
	End Sub
	
	'UPGRADE_NOTE: Form_Initialize was upgraded to Form_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Form_Initialize_Renamed()
		On Error Resume Next
		
	End Sub
	
	Private Sub frmHelp_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error Resume Next
        Me.Icon = frmMain.Icon
	End Sub
	
	'UPGRADE_WARNING: Event frmHelp.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmHelp_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error Resume Next
		If Me.WindowState <> System.Windows.Forms.FormWindowState.Minimized Then
			If VB6.PixelsToTwipsY(Me.Height) > VB6.PixelsToTwipsY(cmdOK.Height) Then
				cmdOK.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - VB6.PixelsToTwipsY(cmdOK.Height) - 90)
				txtHelp.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - VB6.PixelsToTwipsY(cmdOK.Height) - 270)
				txtHelp.Top = VB6.TwipsToPixelsY(90)
			End If
			If VB6.PixelsToTwipsX(Me.Width) > VB6.PixelsToTwipsX(cmdOK.Width) + 180 Then
				txtHelp.Left = VB6.TwipsToPixelsX(90)
				cmdOK.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) / 2 - VB6.PixelsToTwipsX(cmdOK.Width) / 2)
				txtHelp.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - 180)
			End If
		End If
	End Sub
End Class