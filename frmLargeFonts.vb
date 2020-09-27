Option Strict Off
Option Explicit On
Friend Class frmLargeFonts
	Inherits System.Windows.Forms.Form
	'Calendar
	'Copyright Alasdair King, 2010, http://www.alasdairking.me.uk
	'Released under the GNU Public Licence, Version 3.
	
	'Resizes controls to fit contents, typically for large fonts.
	
	
	Public Sub SizeToFit(ByRef c As System.Windows.Forms.Control)
		On Error Resume Next
		Dim newSize As Integer
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Select Case TypeName(c)
			Case "CommandButton"
				lblSizer.Font = c.Font
				'Debug.Print
				'Debug.Print "Size before: " & lblSizer.width
				lblSizer.Text = c.Text
				lblSizer.Refresh()
				'Debug.Print "Size after: " & lblSizer.width
				newSize = VB6.PixelsToTwipsX(lblSizer.Width) * 1.4
				'Debug.Print "Button size: " & c.width
				If VB6.PixelsToTwipsX(c.Width) < newSize Then c.Width = VB6.TwipsToPixelsX(newSize)
				newSize = VB6.PixelsToTwipsY(lblSizer.Height) * 1.8
				If VB6.PixelsToTwipsY(c.Height) < newSize Then c.Height = VB6.TwipsToPixelsY(newSize)
			Case "OptionButton"
				lblSizer.Font = c.Font
				lblSizer.Text = c.Text
				lblSizer.Refresh()
				newSize = VB6.PixelsToTwipsX(lblSizer.Width) * 1 + 600
				If VB6.PixelsToTwipsX(c.Width) < newSize Then
					c.Width = VB6.TwipsToPixelsX(newSize)
				End If
				newSize = VB6.PixelsToTwipsY(lblSizer.Height) * 1.8
				If VB6.PixelsToTwipsY(c.Height) < newSize Then c.Height = VB6.TwipsToPixelsY(newSize)
			Case "CheckBox"
				lblSizer.Font = c.Font
				lblSizer.Text = c.Text
				lblSizer.Refresh()
				newSize = VB6.PixelsToTwipsX(lblSizer.Width) * 1 + 500
				If VB6.PixelsToTwipsX(c.Width) < newSize Then
					c.Width = VB6.TwipsToPixelsX(newSize)
				End If
				newSize = VB6.PixelsToTwipsY(lblSizer.Height) * 1.8
				If VB6.PixelsToTwipsY(c.Height) < newSize Then c.Height = VB6.TwipsToPixelsY(newSize)
			Case "TextBox"
				lblSizer.Font = c.Font
				lblSizer.Text = "Test"
				lblSizer.Refresh()
				newSize = VB6.PixelsToTwipsY(lblSizer.Height) * 1.8
				If VB6.PixelsToTwipsY(c.Height) < newSize Then c.Height = VB6.TwipsToPixelsY(newSize)
			Case Else
				'Debug.Print "TypenName:" & TypeName(c)
		End Select
	End Sub
	
	'UPGRADE_NOTE: Form_Initialize was upgraded to Form_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Form_Initialize_Renamed()
		On Error Resume Next
		Call modXPStyle.InitCommonControlsVB()
	End Sub
End Class