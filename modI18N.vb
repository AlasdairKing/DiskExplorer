Option Strict Off
Option Explicit On
Module modI18N
	'modLanguage
	'Handles the internationalization of all user interface components.
	'Usage: 1 Add to project
	'       2 Ensure that Microsoft Scripting Runtime has a reference in the project
	'       3 Ensure that Microsft XML has a reference in the project
	'       NO LONGER 4 Ensure everything you want to internationalize has a unique .tag attribute
	'       5 Add languages.xml to main program directory and populate with information
	'       6 Change every string "example" to modi18n.GetText("example")
	'       7 Create an instance of the object.
	' Of course, this considerably under-states the work needed for step 5. You can use
	' RegisterForm to help you.
	
	'Copyright (c) 2007, Alasdair King
	'All rights reserved.
	'
	'Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
	'
	'    * Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
	'    * Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
	'    * Neither the name of [Alasdair] nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
	'
	'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
	
	'2 Feb 2007
	'   Made it look up controls without tags: means I don't have to populate the
	'   .tag attribute for every control.
	'14 April 2008
	'   Changed toolbar handling slightly: might need to check tags.
	'7 Dec 2008
	'   Added support for different XP and Vista fonts.
	
	
	Private mTranslator As New Scripting.Dictionary
	Private mstrFontname As String
	Private mintCharset As Short
	Private mLocaleID As Integer
	Private mUserLanguage As String
	Private mLanguageDoc As MSXML2.DOMDocument30 ' The complete language document, holding information about
	' the languages available.
	Private mAppDoc As MSXML2.DOMDocument30 ' The application language document, holding translations
	' for this particular application.
	Private mLanguageFilesNotAvailable As Boolean 'If set then don't do any I18N. Set if fail to find or load application or
	'   language file.
	Private mRightToLeft As Boolean ' indicates this is a right-to-left language, like Hebrew or Arabic. Default it false.
	
	Private Const BLANK As String = ""
	Private Const DEFAULT_LANGUAGE_CODE As String = "en-gb"
	
	'Add the class name of controls with captions or tooltiptext to these strings
	Private Const HAS_CAPTION As String = "*Label*UniLabel*DUniLabel*OptionButton*Menu*CommandButton*Frame*CheckBox*ButtonPlus*FramePlus*SSTab*"
	Private Const HAS_TOOLTIPTEXT As String = "*Label*OptionButton*CommandButton*Frame*CheckBox*DUniLabel*UniLabel*FramePlus*SSTab*TextBox*UniText*UniList*ComboBox*WebBrowser*ProgressBar*ButtonPlus*"
	Private Const HASNOT_FONT As String = "*WebBrowser*Winsock*CommonDialog*Timer*Menu*ProgressBar*Slider*WindowsMediaPlayer*Toolbar*"
	Private Const HAS_FONT As String = "*Label*DUniLabel*UniLabel*ComboBox*TextBox*Frame*CommandButton*CheckBox*OptionButton*ListBox*TabStrip*StatusBar*TreeView*ListView*ImageCombo*DUniText*DUniList*DUniCombo*ButtonPlus*FramePlus*SSTab*"
	
	'To determine language if nothing is defined
	Private Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Short
	'See http://www.codenewsgroups.net/group/microsoft.public.vb.general.discussion/topic2255.aspx
	Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Integer
	Private Declare Function GetLocaleInfo Lib "kernel32"  Alias "GetLocaleInfoA"(ByVal locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer
	Private Declare Function SetLocaleInfo Lib "kernel32"  Alias "SetLocaleInfoA"(ByVal locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String) As Integer
	Private Const LOCALE_USER_DEFAULT As Integer = &H400
	Private Const LOCALE_SYSTEM_DEFAULT As Integer = &H800
	Private Const LOCALE_ILANGUAGE As Integer = &H1 'language id
	Private Const LOCALE_SLANGUAGE As Integer = &H2 'localized name of language
	Private Const LOCALE_SENmodi18n As Integer = &H1001 'English name of language
	Private Const LOCALE_SABBREVLANGNAME As Integer = &H3 'abbreviated language name
	Private Const LOCALE_SNATIVELANGNAME As Integer = &H4 'native name of language
	
	Private Const LOCALE_ICOUNTRY As Integer = &H5 'country code
	Private Const LOCALE_SCOUNTRY As Integer = &H6 'localized name of country
	Private Const LOCALE_SENGCOUNTRY As Integer = &H1002 'English name of country
	Private Const LOCALE_SABBREVCTRYNAME As Integer = &H7 'abbreviated country name
	Private Const LOCALE_SNATIVECTRYNAME As Integer = &H8 'native name of country
	
	Private Const LOCALE_IDEFAULTLANGUAGE As Integer = &H9 'default language id
	Private Const LOCALE_IDEFAULTCOUNTRY As Integer = &HA 'default country code
	Private Const LOCALE_IDEFAULTCODEPAGE As Integer = &HB 'default oem code page
	Private Const LOCALE_IDEFAULTANSICODEPAGE As Integer = &H1004 'default ansi code page
	Private Const LOCALE_IDEFAULTMACCODEPAGE As Integer = &H1011 'default mac code page
	
	Private Const LOCALE_SLIST As Integer = &HC 'list item separator
	Private Const LOCALE_IMEASURE As Integer = &HD '0 = metric, 1 = US
	
	Private Const LOCALE_SDECIMAL As Integer = &HE 'decimal separator
	Private Const LOCALE_STHOUSAND As Integer = &HF 'thousand separator
	Private Const LOCALE_SGROUPING As Integer = &H10 'digit grouping
	Private Const LOCALE_IDIGITS As Integer = &H11 'number of fractional digits
	Private Const LOCALE_ILZERO As Integer = &H12 'leading zeros for decimal
	Private Const LOCALE_INEGNUMBER As Integer = &H1010 'negative number mode
	Private Const LOCALE_SNATIVEDIGITS As Integer = &H13 'native ascii 0-9
	
	Private Const LOCALE_SCURRENCY As Integer = &H14 'local monetary symbol
	Private Const LOCALE_SINTLSYMBOL As Integer = &H15 'intl monetary symbol
	Private Const LOCALE_SMONDECIMALSEP As Integer = &H16 'monetary decimal separator
	Private Const LOCALE_SMONTHOUSANDSEP As Integer = &H17 'monetary thousand separator
	Private Const LOCALE_SMONGROUPING As Integer = &H18 'monetary grouping
	Private Const LOCALE_ICURRDIGITS As Integer = &H19 '# local monetary digits
	Private Const LOCALE_IINTLCURRDIGITS As Integer = &H1A '# intl monetary digits
	Private Const LOCALE_ICURRENCY As Integer = &H1B 'positive currency mode
	Private Const LOCALE_INEGCURR As Integer = &H1C 'negative currency mode
	
	Private Const LOCALE_SDATE As Integer = &H1D 'date separator
	Private Const LOCALE_STIME As Integer = &H1E 'time separator
	Private Const LOCALE_SSHORTDATE As Integer = &H1F 'short date format string
	Private Const LOCALE_SLONGDATE As Integer = &H20 'long date format string
	Private Const LOCALE_STIMEFORMAT As Integer = &H1003 'time format string
	Private Const LOCALE_IDATE As Integer = &H21 'short date format ordering
	Private Const LOCALE_ILDATE As Integer = &H22 'long date format ordering
	Private Const LOCALE_ITIME As Integer = &H23 'time format specifier
	Private Const LOCALE_ITIMEMARKPOSN As Integer = &H1005 'time marker position
	Private Const LOCALE_ICENTURY As Integer = &H24 'century format specifier (short date)
	Private Const LOCALE_ITLZERO As Integer = &H25 'leading zeros in time field
	Private Const LOCALE_IDAYLZERO As Integer = &H26 'leading zeros in day field (short date)
	Private Const LOCALE_IMONLZERO As Integer = &H27 'leading zeros in month field (short date)
	Private Const LOCALE_S1159 As Integer = &H28 'AM designator
	Private Const LOCALE_S2359 As Integer = &H29 'PM designator
	
	Private Const LOCALE_ICALENDARTYPE As Integer = &H1009 'type of calendar specifier
	Private Const LOCALE_IOPTIONALCALENDAR As Integer = &H100B 'additional calendar types specifier
	Private Const LOCALE_IFIRSTDAYOFWEEK As Integer = &H100C 'first day of week specifier
	Private Const LOCALE_IFIRSTWEEKOFYEAR As Integer = &H100D 'first week of year specifier
	
	Private Const LOCALE_SDAYNAME1 As Integer = &H2A 'long name for Monday
	Private Const LOCALE_SDAYNAME2 As Integer = &H2B 'long name for Tuesday
	Private Const LOCALE_SDAYNAME3 As Integer = &H2C 'long name for Wednesday
	Private Const LOCALE_SDAYNAME4 As Integer = &H2D 'long name for Thursday
	Private Const LOCALE_SDAYNAME5 As Integer = &H2E 'long name for Friday
	Private Const LOCALE_SDAYNAME6 As Integer = &H2F 'long name for Saturday
	Private Const LOCALE_SDAYNAME7 As Integer = &H30 'long name for Sunday
	Private Const LOCALE_SABBREVDAYNAME1 As Integer = &H31 'abbreviated name for Monday
	Private Const LOCALE_SABBREVDAYNAME2 As Integer = &H32 'abbreviated name for Tuesday
	Private Const LOCALE_SABBREVDAYNAME3 As Integer = &H33 'abbreviated name for Wednesday
	Private Const LOCALE_SABBREVDAYNAME4 As Integer = &H34 'abbreviated name for Thursday
	Private Const LOCALE_SABBREVDAYNAME5 As Integer = &H35 'abbreviated name for Friday
	Private Const LOCALE_SABBREVDAYNAME6 As Integer = &H36 'abbreviated name for Saturday
	Private Const LOCALE_SABBREVDAYNAME7 As Integer = &H37 'abbreviated name for Sunday
	Private Const LOCALE_SMONTHNAME1 As Integer = &H38 'long name for January
	Private Const LOCALE_SMONTHNAME2 As Integer = &H39 'long name for February
	Private Const LOCALE_SMONTHNAME3 As Integer = &H3A 'long name for March
	Private Const LOCALE_SMONTHNAME4 As Integer = &H3B 'long name for April
	Private Const LOCALE_SMONTHNAME5 As Integer = &H3C 'long name for May
	Private Const LOCALE_SMONTHNAME6 As Integer = &H3D 'long name for June
	Private Const LOCALE_SMONTHNAME7 As Integer = &H3E 'long name for July
	Private Const LOCALE_SMONTHNAME8 As Integer = &H3F 'long name for August
	Private Const LOCALE_SMONTHNAME9 As Integer = &H40 'long name for September
	Private Const LOCALE_SMONTHNAME10 As Integer = &H41 'long name for October
	Private Const LOCALE_SMONTHNAME11 As Integer = &H42 'long name for November
	Private Const LOCALE_SMONTHNAME12 As Integer = &H43 'long name for December
	Private Const LOCALE_SMONTHNAME13 As Integer = &H100E 'long name for 13th month (if exists)
	Private Const LOCALE_SABBREVMONTHNAME1 As Integer = &H44 'abbreviated name for January
	Private Const LOCALE_SABBREVMONTHNAME2 As Integer = &H45 'abbreviated name for February
	Private Const LOCALE_SABBREVMONTHNAME3 As Integer = &H46 'abbreviated name for March
	Private Const LOCALE_SABBREVMONTHNAME4 As Integer = &H47 'abbreviated name for April
	Private Const LOCALE_SABBREVMONTHNAME5 As Integer = &H48 'abbreviated name for May
	Private Const LOCALE_SABBREVMONTHNAME6 As Integer = &H49 'abbreviated name for June
	Private Const LOCALE_SABBREVMONTHNAME7 As Integer = &H4A 'abbreviated name for July
	Private Const LOCALE_SABBREVMONTHNAME8 As Integer = &H4B 'abbreviated name for August
	Private Const LOCALE_SABBREVMONTHNAME9 As Integer = &H4C 'abbreviated name for September
	Private Const LOCALE_SABBREVMONTHNAME10 As Integer = &H4D 'abbreviated name for October
	Private Const LOCALE_SABBREVMONTHNAME11 As Integer = &H4E 'abbreviated name for November
	Private Const LOCALE_SABBREVMONTHNAME12 As Integer = &H4F 'abbreviated name for December
	Private Const LOCALE_SABBREVMONTHNAME13 As Integer = &H100F 'abbreviated name for 13th month (if exists)
	
	Private Const LOCALE_SPOSITIVESIGN As Integer = &H50 'positive sign
	Private Const LOCALE_SNEGATIVESIGN As Integer = &H51 'negative sign
	Private Const LOCALE_IPOSSIGNPOSN As Integer = &H52 'positive sign position
	Private Const LOCALE_INEGSIGNPOSN As Integer = &H53 'negative sign position
	Private Const LOCALE_IPOSSYMPRECEDES As Integer = &H54 'mon sym precedes pos amt
	Private Const LOCALE_IPOSSEPBYSPACE As Integer = &H55 'mon sym sep by space from pos amt
	Private Const LOCALE_INEGSYMPRECEDES As Integer = &H56 'mon sym precedes neg amt
	Private Const LOCALE_INEGSEPBYSPACE As Integer = &H57 'mon sym sep by space from neg amt
	
	Private Const LOCALE_FONTSIGNATURE As Integer = &H58 'font signature
	Private Const LOCALE_SISO639LANGNAME As Integer = &H59 'ISO abbreviated language name
	Private Const LOCALE_SISO3166CTRYNAME As Integer = &H5A 'ISO abbreviated country name
	
	Private Const LOCALE_IDEFAULTEBCDICCODEPAGE As Integer = &H1012 'default ebcdic code page
	Private Const LOCALE_IPAPERSIZE As Integer = &H100A '0 = letter, 1 = a4, 2 = legal, 3 = a3
	Private Const LOCALE_SENGCURRNAME As Integer = &H1007 'english name of currency
	Private Const LOCALE_SNATIVECURRNAME As Integer = &H1008 'native name of currency
	Private Const LOCALE_SYEARMONTH As Integer = &H1006 'year month format string
	Private Const LOCALE_SSORTNAME As Integer = &H1013 'sort name
	Private Const LOCALE_IDIGITSUBSTITUTION As Integer = &H1014 '0 = none, 1 = context, 2 = native digit
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function GetPrivateProfileStrinmodi18n Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function WritePrivateProfileStrinmodi18n Lib "kernel32"  Alias "WritePrivateProfileStringA"(ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
	
	Private Structure OSVERSIONINFO
		Dim dwVersionInfoSize As Integer
		Dim dwMajorVersion As Integer
		Dim dwMinorVersion As Integer
		Dim dwBuildNumber As Integer
		Dim dwPlatformId As Integer
		<VBFixedArray(127)> Dim szCSDVersion() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim szCSDVersion(127)
		End Sub
	End Structure
	'UPGRADE_WARNING: Structure OSVERSIONINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetVersionEx Lib "kernel32"  Alias "GetVersionExA"(ByRef lpVersionInfo As OSVERSIONINFO) As Integer
	
	Private mInitialised As Boolean ' indicates that the initial setup process - determining what language
	'   we're using and loading the language files.
	
	'UPGRADE_WARNING: Lower bound of array DaysOfTheWeek was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public DaysOfTheWeek(7) As String
	
	Public Function rightToLeft() As Boolean
		On Error Resume Next
		rightToLeft = mRightToLeft
	End Function
	
	Public Sub ApplyUILanguageToThisForm(ByRef aForm As System.Windows.Forms.Form, Optional ByRef enforceTooltipSpacing As Boolean = True)
		'goes through the form updating the user components
#If debugging = 0 Then
		On Error Resume Next
#End If
		Dim aControl As System.Windows.Forms.Control
		Dim got As String
		Dim index As String
		Dim controlType As String
		Dim newFont As System.Drawing.Font
		Dim tbButton As Object
		Dim aPanel As Object
		Dim aTab As Object
		
		If Not mInitialised Then Call Initialise()
		If mLanguageFilesNotAvailable Then
			'don't do any translation
		Else
			'First, do the form:
			'do caption
#If debugging Then
			'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression debugging did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
			Open App.Path & "\language.log" For Append As #1
			Write #1, aForm.name & ".Caption"
			Write #1, aForm.Caption
			Close #1
#End If
			'UPGRADE_WARNING: Couldn't resolve default property of object mTranslator.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			got = mTranslator.Item(aForm.Name & ".Caption")
			If Len(got) > 0 Then aForm.Text = got
			'do font if defined in language file
			If Len(mstrFontname) > 0 Then aForm.Font = VB6.FontChangeName(aForm.Font, mstrFontname)
			If mintCharset > -1 Then aForm.Font = VB6.FontChangeGdiCharSet(aForm.Font, mintCharset)
			If mRightToLeft Then aForm.RightToLeft = System.Windows.Forms.RightToLeft.Yes
			
			'Second, do all controls on the form
			'Handle special cases where the control has "sub-controls" - status bar and toolbar and tabstrip
			For	Each aControl In aForm.Controls
				'        Debug.Print "Trying: " & aControl.Tag
				'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If TypeName(aControl) = "StatusBar" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object aControl.Panels. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For	Each aPanel In aControl.Panels
						'UPGRADE_WARNING: Couldn't resolve default property of object aPanel.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Len(aPanel.Tag) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object aPanel.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							index = aPanel.Tag
						Else
							index = ""
							'UPGRADE_WARNING: Couldn't resolve default property of object aPanel.index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							index = aForm.Name & "." & aControl.Name & "." & aPanel.index
							'UPGRADE_WARNING: Couldn't resolve default property of object aPanel.index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If index = "" Then index = aForm.Name & "." & aControl.Name & "." & aPanel.index
						End If
						If mTranslator.Exists(index & ".ToolTipText") Then
							got = ""
							'UPGRADE_WARNING: Couldn't resolve default property of object mTranslator.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							got = mTranslator.Item(index & ".ToolTipText")
							If got <> BLANK Then
								If enforceTooltipSpacing And Len(got) > 0 Then
									got = Trim(got)
									got = " " & got & " "
								End If
								'UPGRADE_WARNING: Couldn't resolve default property of object aPanel.ToolTipText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								aPanel.ToolTipText = got
							End If
						Else
							'No Tooltip in current language.
#If debugging Then
							'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression debugging did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
							Open App.Path & "\language.log" For Append As #1
							Write #1, index & ".ToolTipText"
							Close #1
#End If
						End If
					Next aPanel
					'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				ElseIf TypeName(aControl) = "Toolbar" Then 
					'updates toolbar controls with international text
					'UPGRADE_WARNING: Couldn't resolve default property of object aControl.Buttons. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For	Each tbButton In aControl.Buttons
						'UPGRADE_WARNING: Couldn't resolve default property of object tbButton.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Len(tbButton.Tag) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object tbButton.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							index = tbButton.Tag
						Else
							'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
							If TypeName(tbButton) = "IButton" Then
								'UPGRADE_WARNING: Couldn't resolve default property of object tbButton.key. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								index = aForm.Name & "." & aControl.Name & "." & tbButton.key
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object tbButton.name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								index = aForm.Name & "." & aControl.Name & "." & tbButton.name
							End If
						End If
						If mTranslator.Exists(index & ".Caption") Then
							got = ""
							'UPGRADE_WARNING: Couldn't resolve default property of object mTranslator.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							got = mTranslator.Item(index & ".Caption")
							'UPGRADE_WARNING: Couldn't resolve default property of object tbButton.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If got <> BLANK Then tbButton.Caption = got
						Else
#If debugging Then
							'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression debugging did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
							Open App.Path & "\language.log" For Append As #1
							Write #1, "<item><key>" & index & ".Caption</key><explanation>String</explanation><content language=""en-gb"">" & aForm.Caption & "</content></item>" & vbNewLine
							Close #1
#End If
						End If
						If mTranslator.Exists(index & ".ToolTipText") Then
							got = ""
							'UPGRADE_WARNING: Couldn't resolve default property of object mTranslator.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							got = mTranslator.Item(index & ".ToolTipText")
							If got <> BLANK Then
								If enforceTooltipSpacing And Len(got) > 0 Then
									got = Trim(got)
									got = " " & got & " "
								End If
								'UPGRADE_WARNING: Couldn't resolve default property of object tbButton.ToolTipText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								tbButton.ToolTipText = got
							Else
#If debugging Then
								'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression debugging did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
								Open App.Path & "\language.log" For Append As #1
								Write #1, "<item><key>" & index & ".ToolTipText</key><explanation>String</explanation><content language=""en-gb"">" & tbButton.ToolTipText & "</content></item>" & vbNewLine
								Close #1
#End If
							End If
						End If
					Next tbButton
					'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				ElseIf TypeName(aControl) = "TabStrip" Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object aControl.Tabs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For	Each aTab In aControl.Tabs
						'UPGRADE_WARNING: Couldn't resolve default property of object aTab.key. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						index = aForm.Name & "." & aControl.Name & "." & aTab.key
						If mTranslator.Exists(index & ".Caption") Then
							got = ""
							'UPGRADE_WARNING: Couldn't resolve default property of object mTranslator.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							got = mTranslator.Item(index & ".Caption")
							'UPGRADE_WARNING: Couldn't resolve default property of object aTab.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If got <> BLANK Then aTab.Caption = got
						Else
#If debugging Then
							'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression debugging did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
							Open App.Path & "\language.log" For Append As #1
							Write #1, "<item><key>" & index & ".Caption</key><explanation>String</explanation><content language=""en-gb"">" & aTab.Caption & "</content></item>" & vbNewLine
							Close #1
#End If
						End If
					Next aTab
				End If
				'Work out what this item should be looked up as. Is this defined in the
				'tag for the control?
				If Len(aControl.Tag) > 0 Then
					index = aControl.Tag
				Else
					index = aForm.Name & "." & aControl.Name
				End If
				'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				controlType = "*" & TypeName(aControl) & "*"
				'Caption
				If InStr(1, HAS_CAPTION, controlType) > 0 Then
					'Debug.Print "Tag: " & aControl.Tag
					got = ""
					'UPGRADE_WARNING: Couldn't resolve default property of object mTranslator.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					got = mTranslator.Item(index & ".Caption")
					If Len(got) > 0 Then
						If got <> BLANK Then aControl.Text = got
						'OK, in WebbIE 3.6 the conversion below breaks the correct character encoding: for
						'example, in Polish, the l with a line through it becomes 3 in mnuWebsite in Accessible
						'RSS. But I apparently got it working okay in previous versions. However, I did rework
						'all the I18N in 3.6, so let's assume I broke it and doing a direct assignment (as
						'above) works fine so long as the codepage for the Operating System is set correctly.
						'                If mLocaleID > -1 Then
						'                    aControl.Caption = StrConv(StrConv(got, vbFromUnicode, mLocaleID), vbUnicode)
						'                Else
						'                    aControl.Caption = StrConv(StrConv(got, vbFromUnicode, 0), vbUnicode)
						'                End If
					Else
#If debugging Then
						'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression debugging did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
						Open App.Path & "\language.log" For Append As #1
						Write #1, "<item><key>" & index & ".Caption</key><explanation>String</explanation><content language=""en-gb"">" & aControl.Caption & "</content></item>" & vbNewLine
						Close #1
#End If
					End If
				Else
					'Does not have .Caption property
					'            Debug.Print "No caption for: " & index & " " & controlType
				End If
				'Tooltip
				If InStr(1, HAS_TOOLTIPTEXT, controlType) > 0 Then
					got = ""
					'UPGRADE_WARNING: Couldn't resolve default property of object mTranslator.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					got = mTranslator.Item(index & ".ToolTipText")
					If got <> BLANK Then
						If enforceTooltipSpacing And Len(got) > 0 Then
							got = Trim(got)
							got = " " & got & " "
						End If
						'UPGRADE_ISSUE: Control property aControl.ToolTipText was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						aControl.ToolTipText = got
					End If
				Else
					'Debug.Print "No Tooltiptext for: " & index & " " & controlType
				End If
				'Font object, if defined
				If mintCharset > -1 Or Len(mstrFontname) > 0 Then
					If InStr(1, HAS_FONT, controlType) > 0 Then
						newFont = aControl.Font
						'WebbIE 3.8.0: need to do fontname BEFORE charset, or it won't change for
						'CommandButtons.
						If Len(mstrFontname) > 0 Then newFont = VB6.FontChangeName(newFont, mstrFontname)
						If mintCharset > -1 Then newFont = VB6.FontChangeGdiCharSet(newFont, mintCharset)
						aControl.Font = newFont
						'Debug.Print "Applying font to: " & aForm.name & "." & controlType
					Else
						'Debug.Print "No font: " & controlType
					End If
				End If
			Next aControl
		End If
	End Sub
	
	'Public Sub RegisterForm(aForm As Form)
	''saves the information about the form
	'    on error resume next
	'    Dim aControl As Control
	'    Dim parentName As String
	'    Dim fso As New FileSystemObject
	'    Dim ts As TextStream
	'    Dim item As String
	'
	'    parentName = aForm.name
	'    Set ts = fso.OpenTextFile(modPath.settingsPath & "\allforms.txt", ForAppending, True, TristateFalse)
	'    item = vbTab & vbTab & "<item>" & vbNewLine & "<key>" & parentName & ".Caption</key><explanation/>"
	'    item = item & "<content language=""en-gb"">" & aForm.Caption & "</content>" & vbNewLine & "</item>" & vbNewLine
	'    Call ts.Write(item)
	'    For Each aControl In aForm.Controls
	'        If Len(aControl.Tag) > 0 Then
	'            'got a tag: record details
	'            'Caption
	'            If TypeOf aControl Is Label Or TypeOf aControl Is OptionButton Or _
	''            TypeOf aControl Is Menu Or TypeOf aControl Is CommandButton Or _
	''            TypeOf aControl Is Frame Or TypeOf aControl Is CheckBox Then
	'                item = vbTab & vbTab & "<item><key>" & aControl.Tag & ".Caption" & "</key><explanation/><content language=""en-gb"">" & aControl.Caption & "</content></item>" & vbNewLine
	'                Call ts.Write(item)
	'            End If
	'        End If
	'    Next aControl
	'    Call ts.Close
	'End Sub
	
	Public Function GetText(ByRef key As String, Optional ByRef enforceTooltipSpacing As Boolean = False) As String
		'returns the internationalised string for the given key.
#If debugging = 0 Then
		On Error Resume Next
#End If
		Dim got As String
		
		If Not mInitialised Then Call Initialise()
		If mLanguageFilesNotAvailable Then
			GetText = key
		ElseIf key = "" Then 
		Else
			If mTranslator.Exists(key) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mTranslator.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				got = mTranslator.Item(key)
				If got <> BLANK Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mTranslator.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetText = mTranslator.Item(key)
					If enforceTooltipSpacing And Len(got) > 0 Then
						GetText = Trim(GetText)
						GetText = " " & GetText & " "
					End If
				End If
			Else
#If debugging Then
				'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression debugging did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
				Open App.Path & "\language.log" For Append As #1
				Print #1, "<item><key>" & key & "</key><explanation>String</explanation><content language=""en-gb"">" & key & "</content></item>"
				Close #1
#End If
				GetText = key
			End If
		End If
	End Function
	
	
	Private Sub Initialise()
		'Loads language/I18N information ready for subsequent translation work.
		'In order of preference:
		'1 An external INI file called EXEName.Language.ini. This gets loaded and parsed into the en-gb language
		'   and en-gb is set as the default language. The result is an XML document.
		'2 An external XML file called EXEName.Language.xml. This replaces the embedded XML file directly.
		'3 The internal XML file for language.
		'Having parsed the XML file into mLanguageDoc, then try to work out what language to use:
		'1 If an external ini file, use en-gb (default)
		'2 Read from common WebbIE files > WebbIE 3.ini
		'3 Read from common WebbIE files > WebbIE3.ini
		'4 Read from registry in HKCU to support older WebbIE versions.
		'5 Read from EXEName.ini
		'6 Try to work it out from the Windows Locale.
		'7 Default to en-gb.
		'The language is then loaded, and font, charset and localeID set. But translating the UI doesn't happen
		'until prompted.
		
#If debugging = 0 Then
		On Error Resume Next
#End If
		Dim result As String
		Dim locale As Integer
		Dim LCType As Integer
		Dim lpLCData As String
		Dim cchData As Integer
		Dim languageIterator As MSXML2.IXMLDOMNode
		Dim s As String
		
		mInitialised = True
		'Set some defaults
		If IsVistaOrAbove Then
			mstrFontname = "Segoe UI"
		Else
			mstrFontname = "Tahoma"
		End If
		
#If debugging Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression debugging did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		If Dir(App.Path & "\language.log", vbNormal) <> "" Then
		Call Kill(App.Path & "\language.log")
		End If
#End If
		mLanguageDoc = New MSXML2.DOMDocument30
		mAppDoc = New MSXML2.DOMDocument30
		'First, load language information. This is either internal to the application (as a resource) or external
		'(as an XML file or ini file)
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Dir(GetAppPath & "\" & My.Application.Info.AssemblyName & ".Language.ini", FileAttribute.Normal) <> "" Then
			'Found ini file. Load this in preference to XML.
			'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			Call LoadXMLDocFromIni(GetAppPath & "\" & My.Application.Info.AssemblyName & ".Language.ini")
			'Always set language to "local" = the default for ini use.
			mUserLanguage = "local"
			Call SetLanguage("local")
		Else
			'nope, no Ini file. XML?
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Dir(modPath.GetAppPath & "\Languages.xml", FileAttribute.Normal) <> "" Then
				'found external XML. Try loading it.
				mLanguageDoc.async = False
				mLanguageDoc.preserveWhiteSpace = False
				mLanguageDoc.validateOnParse = False
				mLanguageDoc.resolveExternals = False
				Call mLanguageDoc.Load(modPath.GetAppPath & "\Languages.xml")
			End If
			'Now load the application-specific translations.
			mAppDoc.async = False
			mAppDoc.preserveWhiteSpace = False
			mAppDoc.validateOnParse = False
			mAppDoc.resolveExternals = False
			'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			Call mAppDoc.Load(modPath.GetAppPath & "\" & My.Application.Info.AssemblyName & ".Language.xml")
			'Assertion: loaded the language file by now.
			
			If mAppDoc.parseError.errorCode <> 0 Or mLanguageDoc.parseError.errorCode <> 0 Then
				'One or both failed to load.
				mLanguageFilesNotAvailable = True
			Else
				'Now, what language shall we use?
				'First thing to check is for a Claro-based I18N setting which overrides the WebbIE
				'setting. If you're reading this under the Open-source licence: this is to allow for distribution
				'by Claro Software of Accessible PDF using their I18N mechanism. Feel free to ignore.
				result = modPath.ReadAppEXEIni("I18N", "ClaroLanguage", "")
				
				If Len(result) > 0 Then
					mUserLanguage = result
				Else
					'Let's check to see if the user has set the
					'language in one of the various places we might have saved it.
					'First, the new official place to save it.
					result = ReadIniFileLanguage(modPath.commonSettingsPath & "\WebbIE 3.ini", "Internationalisation", "Language")
					If Len(result) = 0 Then result = ReadIniFileLanguage(modPath.commonSettingsPath & "\WebbIE3.ini", "Internationalisation", "Language")
					If Len(result) > 0 Then
						'Good, got something from there.
						mUserLanguage = result
					Else
						'Nope. Let's try the old place to save language info, the registry
						mUserLanguage = GetSetting("WebbIE 3", "User Settings", "Language", "Nothing recorded")
						If mUserLanguage = "Nothing recorded" Then
							'Nope. Let's see if installation has written a default language to
							'program files.
							result = ReadIniFileLanguage(GetAppPath & "\language.ini", "Internationalisation", "Language")
							If Len(result) = 0 Then result = modPath.ReadAppEXEIni("Internationalisation", "Language", "")
							If Len(result) > 0 Then
								'right, got a result from the default language file.
								mUserLanguage = result
							Else
								'okay, we don't have a defined language: use Windows information
								'to make a stab at it
								mUserLanguage = ""
								lpLCData = Space(255) & Chr(0)
								cchData = Len(lpLCData)
								LCType = LOCALE_SCOUNTRY
								locale = LOCALE_USER_DEFAULT
								Call GetLocaleInfo(locale, LCType, lpLCData, cchData)
								'locale now contains the locale ID for the current user or system.
								'Go through the language file looking for matching languages.
								If mLanguageDoc.parseError.errorCode = 0 Then
									For	Each languageIterator In mLanguageDoc.documentElement.selectNodes("languages/language")
										If CInt(languageIterator.Attributes.getNamedItem("localeID").text) = locale Then
											'Windows locale matches this language: use it
											mUserLanguage = languageIterator.Attributes.getNamedItem("id").text
											Exit For
										End If
									Next languageIterator
								End If
								If mUserLanguage = "" Then
									'okay, haven't got anything at all. Default to British English.
									mUserLanguage = DEFAULT_LANGUAGE_CODE
								End If
							End If
						End If
					End If
				End If
				Call SetLanguage(mUserLanguage)
			End If
		End If
		DaysOfTheWeek(1) = GetText("Sunday")
		DaysOfTheWeek(2) = GetText("Monday")
		DaysOfTheWeek(3) = GetText("Tuesday")
		DaysOfTheWeek(4) = GetText("Wednesday")
		DaysOfTheWeek(5) = GetText("Thursday")
		DaysOfTheWeek(6) = GetText("Friday")
		DaysOfTheWeek(7) = GetText("Saturday")
	End Sub
	
	'Private Sub Class_Terminate()
	'    on error resume next
	'    'I'm no longer (28 Jan 2008) going to write the language chosen by a particular application, because
	'    'it is too easy to have a program than fails to read the correct language and then defaults back
	'    'to English. So leave this to the installer or LanguageSelector.exe
	'    'Call WritePrivateProfileString("Internationalisation", "Language", CStr(mUserLanguage), modPath.commonSettingsPath & "\WebbIE 3.ini")
	'    'also write to registry to support old installed apps
	'    If modPath.runningLocal Then
	'        'don't store in registry
	'    Else
	'        'running in installed machine: No, don't do this. Take language
	'        'out of WebbIE.
	'        'Call SaveSetting("WebbIE 3", "User Settings", "Language", CStr(mUserLanguage))
	'    End If
	'End Sub
	
	Public Function GetLanguage() As String
#If debugging = 0 Then
		On Error Resume Next
#End If
		If Not mInitialised Then Call Initialise()
		GetLanguage = mUserLanguage
	End Function
	
	Public Sub SetLanguage(ByRef newLanguage As String)
		'loads and parses the language requested, or uses English if not found
#If debugging = 0 Then
		On Error Resume Next
#End If
		Dim itemIterator As MSXML2.IXMLDOMNode
		Dim languageNode As MSXML2.IXMLDOMNode
		Dim contentNode As MSXML2.IXMLDOMNode
		Dim i As Short
		Dim s As String
		Dim attNode As MSXML2.IXMLDOMAttribute
		If Not mInitialised Then Call Initialise()
		
		If mLanguageFilesNotAvailable Then
			'Don't do anything, failed to load language files
		Else
			'Assertion: have already loaded language file in the Class_Initialize event, so
			'no need to load it now.
			'Unless of course we failed to load it.
			If mLanguageDoc.parseError.errorCode = 0 And Len(mLanguageDoc.xml) > 0 Then
				mRightToLeft = False ' default
				mUserLanguage = newLanguage ' we should really check this is a valid language, but what the heck.
				If mLanguageDoc.documentElement.selectSingleNode("languages/language[@id='" & mUserLanguage & "']") Is Nothing Then
					mstrFontname = ""
				Else
					mstrFontname = mLanguageDoc.documentElement.selectSingleNode("languages/language[@id='" & mUserLanguage & "']").Attributes.getNamedItem("font").text
					'Support for a Vista font: supply a "vistaFont" attribute.
					attNode = mLanguageDoc.documentElement.selectSingleNode("languages/language[@id='" & mUserLanguage & "']").Attributes.getNamedItem("vistaFont")
					If Not (attNode Is Nothing) And IsVistaOrAbove Then
						If Len(attNode.text) > 0 Then
							mstrFontname = attNode.text
						End If
					End If
				End If
				If mLanguageDoc.documentElement.selectSingleNode("languages/language[@id='" & mUserLanguage & "']") Is Nothing Then
					mintCharset = -1
				Else
					mintCharset = CShort(mLanguageDoc.documentElement.selectSingleNode("languages/language[@id='" & mUserLanguage & "']").Attributes.getNamedItem("charset").text)
				End If
				If mLanguageDoc.documentElement.selectSingleNode("languages/language[@id='" & mUserLanguage & "']") Is Nothing Then
					mLocaleID = -1
				Else
					languageNode = mLanguageDoc.documentElement.selectSingleNode("languages/language[@id='" & mUserLanguage & "']")
					mstrFontname = languageNode.Attributes.getNamedItem("font").text
					mintCharset = CShort(languageNode.Attributes.getNamedItem("charset").text)
					mLocaleID = CInt(languageNode.Attributes.getNamedItem("localeID").text)
					If Not languageNode.Attributes.getNamedItem("rightToLeft") Is Nothing Then
						mRightToLeft = CBool(languageNode.Attributes.getNamedItem("rightToLeft").text)
					End If
				End If
				Call mTranslator.removeAll()
				languageNode = mAppDoc.documentElement.selectSingleNode("contents")
				For	Each itemIterator In languageNode.selectNodes("item")
					contentNode = itemIterator.selectSingleNode("content[@language='" & mUserLanguage & "']")
					If contentNode Is Nothing Then
						'didn't find desired content: use English (or whatever default is)
						contentNode = itemIterator.selectSingleNode("content[@language='" & DEFAULT_LANGUAGE_CODE & "']")
					End If
					'now add to translation dictionary
					'        Debug.Print "Adding: " & itemIterator.selectSingleNode("key").Text
					'        Debug.Print "With: " & contentNode.Text
#If debugging Then
					'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression debugging did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
					If mTranslator.Exists(itemIterator.selectSingleNode("key").text) Then
					Call languageNode.removeChild(itemIterator)
					Set itemIterator = Nothing
					End If
#End If
					If Not itemIterator Is Nothing Then
						If contentNode.text = "" Then
							Call mTranslator.Add(itemIterator.selectSingleNode("key").text, BLANK)
						Else
							Call mTranslator.Add(itemIterator.selectSingleNode("key").text, contentNode.text)
						End If
					End If
				Next itemIterator
			End If
#If debugging Then
			'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression debugging did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
			Call mAppDoc.save(mAppDoc.url & ".bak")
#End If
		End If
ErrorInitialise: 
	End Sub
	
	Public Sub ApplyUILanguageToAllForms()
#If debugging = 0 Then
		On Error Resume Next
#End If
		'processes the whole application updating every user interface component on every loaded form
		Dim aForm As System.Windows.Forms.Form
		If Not mInitialised Then Call Initialise()
		If mLanguageFilesNotAvailable Then
		Else
			For	Each aForm In My.Application.OpenForms
				'        Debug.Print "Doing: " & aForm.name
				Call ApplyUILanguageToThisForm(aForm)
			Next aForm
		End If
	End Sub
	
	Private Sub LoadXMLDocFromIni(ByRef Path As String)
#If debugging = 0 Then
		On Error Resume Next
#End If
		'loads the ini file indicated by path into the mlanguagedoc file
		Dim fso As Scripting.FileSystemObject
		fso = New Scripting.FileSystemObject
		Dim ts As Scripting.TextStream
		Dim got As String
		Dim newItem As MSXML2.IXMLDOMNode
		Dim parts() As String
		
		mLanguageDoc = New MSXML2.DOMDocument30
		mAppDoc = New MSXML2.DOMDocument30
		mLanguageDoc.async = False
		mAppDoc.async = False
		'load the xml frameworks
		Call mLanguageDoc.loadXML("<language version=""3""><languages><language id=""local"" name=""Locally-defined Language"" font=""Tahoma"" charset=""0"" localeID=""0""><localname>Locally-defined Langauge</localname></language></languages><contents/></language>")
		Call mAppDoc.loadXML("<language version=""3""><contents/><help/><popupHelp/></language>")
		'parse the ini file to extract language information
		ts = fso.OpenTextFile(Path, Scripting.IOMode.ForReading, False, Scripting.Tristate.TristateTrue)
		While Not ts.AtEndOfStream
			got = ts.ReadLine
			
			If Left(got, 1) = "[" Or Len(Trim(got)) = 0 Or InStr(1, got, "=", CompareMethod.Text) = 0 Then
				'ignore this line
			ElseIf StrComp(Left(got, 9), "FontName=", CompareMethod.Text) = 0 Then 
				'got a specific special font instruction
				mLanguageDoc.documentElement.selectSingleNode("languages/language").Attributes.getNamedItem("font").text = Trim(Replace(got, "FontName=", "", CompareMethod.Text))
			ElseIf StrComp(Left(got, 8), "Charset=", CompareMethod.Text) = 0 Then 
				'got a charset instruction
				mLanguageDoc.documentElement.selectSingleNode("languages/language").Attributes.getNamedItem("charset").text = Trim(Replace(got, "Charset=", "", CompareMethod.Text))
			ElseIf StrComp(Left(got, 9), "LocaleID=", CompareMethod.Text) = 0 Then 
				'got a locale ID instruction
				mLanguageDoc.documentElement.selectSingleNode("languages/language").Attributes.getNamedItem("localeID").text = Trim(Replace(got, "LocaleID=", "", CompareMethod.Text))
			Else
				'got an entry to add
				parts = Split(got, "=", 2)
				If UBound(parts) < 1 Then
					'don't add, not got all the bits!
				Else
					'Add to loaded document.
					newItem = mAppDoc.createNode(MSXML2.tagDOMNodeType.NODE_ELEMENT, "item", "")
					Call newItem.appendChild(mAppDoc.createNode(MSXML2.tagDOMNodeType.NODE_ELEMENT, "key", ""))
					newItem.selectSingleNode("key").text = parts(0)
					Call newItem.appendChild(mAppDoc.createNode(MSXML2.tagDOMNodeType.NODE_ELEMENT, "content", ""))
					'UPGRADE_WARNING: Couldn't resolve default property of object mAppDoc.createAttribute(language). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call newItem.selectSingleNode("content").Attributes.setNamedItem(mAppDoc.createAttribute("language"))
					newItem.selectSingleNode("content").Attributes.getNamedItem("language").text = "local"
					parts(1) = Trim(parts(1))
					If Left(parts(1), 1) = """" And Right(parts(1), 1) = """" Then
						parts(1) = Left(parts(1), Len(parts(1)) - 1)
						parts(1) = Right(parts(1), Len(parts(1)) - 1)
					End If
					newItem.selectSingleNode("content").text = parts(1)
					Call mAppDoc.documentElement.selectSingleNode("contents").appendChild(newItem)
				End If
			End If
		End While
		Call ts.Close()
		'Debug.Print mLanguageDoc.xml
	End Sub
	
	Public Function fontDefined() As Boolean
#If debugging = 0 Then
		On Error Resume Next
#End If
		If Not mInitialised Then Call Initialise()
		fontDefined = (Len(mstrFontname) > 0)
	End Function
	
	Public Function helpTopicCount() As Short
#If debugging = 0 Then
		On Error Resume Next
#End If
		Dim helpNode As MSXML2.IXMLDOMNode
		
		If Not mInitialised Then Call Initialise()
		If mLanguageFilesNotAvailable Then
		Else
			'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Dir(GetAppPath & "\" & My.Application.Info.AssemblyName & ".Help.txt") <> "" Then
				'Got ini file to use as help
				helpTopicCount = 1
			Else
				helpNode = mAppDoc.documentElement.selectSingleNode("help")
				If helpNode Is Nothing Then
					helpTopicCount = 0
				Else
					helpTopicCount = helpNode.selectNodes("topic").length
				End If
			End If
		End If
	End Function
	
	Public Function helpTopicTitle(ByRef index As Short) As String
#If debugging = 0 Then
		On Error Resume Next
#End If
		Dim got As String
		
		If Not mInitialised Then Call Initialise()
		If mLanguageFilesNotAvailable Then
		Else
			got = helpTopicText(index)
			If Len(got) > 0 Then helpTopicTitle = Left(got, InStr(1, got, vbNewLine) - 1)
		End If
	End Function
	
	Public Function helpTopicText(ByRef index As Short) As String
#If debugging = 0 Then
		On Error Resume Next
#End If
		Dim fso As Scripting.FileSystemObject
		Dim ts As Scripting.TextStream
		Dim topicNode As MSXML2.IXMLDOMNode
		Dim topics As MSXML2.IXMLDOMNodeList
		Dim contentNode As MSXML2.IXMLDOMNode
		
		If Not mInitialised Then Call Initialise()
		
		If mLanguageFilesNotAvailable Then
		Else
			'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Dir(GetAppPath & "\" & My.Application.Info.AssemblyName & ".Help.txt") <> "" Then
				'local Unicode help file to use
				fso = New Scripting.FileSystemObject
				'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				ts = fso.OpenTextFile(GetAppPath & "\" & My.Application.Info.AssemblyName & ".Help.txt", Scripting.IOMode.ForReading, False, Scripting.Tristate.TristateTrue)
				helpTopicText = ts.ReadAll
				Call ts.Close()
				'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				fso = Nothing
			Else
				'Use XML if valid
				topics = mAppDoc.documentElement.selectNodes("help/topic")
				If topics Is Nothing Then
				ElseIf topics.length = 0 Then 
				ElseIf index < 0 Or index > topics.length - 1 Then 
				Else
					topicNode = topics.Item(index)
					contentNode = topicNode.selectSingleNode("content[@language=""" & mUserLanguage & """]")
					If contentNode Is Nothing Then contentNode = topicNode.selectSingleNode("content[@language=""en-gb""]")
					If contentNode Is Nothing Then
						'No topic!
					Else
						helpTopicText = ProcessHelpNode(contentNode)
					End If
				End If
			End If
		End If
	End Function
	
	Private Function ProcessHelpNode(ByRef helpNode As MSXML2.IXMLDOMNode) As String
#If debugging = 0 Then
		On Error Resume Next
#End If
		'iterate through the helpNode inserting newlines where indicated by
		'<p> nodes
		Dim paragraphNode As MSXML2.IXMLDOMNode
		Dim output As String
		
		output = helpNode.Attributes.getNamedItem("title").text & vbNewLine & vbNewLine
		For	Each paragraphNode In helpNode.selectNodes("p")
			output = output & paragraphNode.text & vbNewLine
		Next paragraphNode
		ProcessHelpNode = output
	End Function
	
	Private Function ReadIniFileLanguage(ByVal strIniFile As String, ByVal strSECTION As String, ByVal strKey As String) As String
#If debugging = 0 Then
		On Error Resume Next
#End If
		Dim strBuffer As String
		Dim strNull As String
		
		strBuffer = Space(256)
		If GetPrivateProfileStrinmodi18n(strSECTION, strKey, strNull, strBuffer, Len(strBuffer) - 1, strIniFile) > 0 Then
			ReadIniFileLanguage = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
		Else
			ReadIniFileLanguage = ""
		End If
	End Function
	
	Public Function popupHelp(ByRef key As String) As String
#If debugging = 0 Then
		On Error Resume Next
#End If
		'look up popup help
		Dim resultNode As MSXML2.IXMLDOMNode
		Dim textNode As MSXML2.IXMLDOMNode
		
		If mLanguageFilesNotAvailable Then
		Else
			resultNode = mAppDoc.documentElement.selectSingleNode("popupHelp/item[key=""" & key & """]")
			If resultNode Is Nothing Then
			Else
				textNode = resultNode.selectSingleNode("content[@language=""" & mUserLanguage & """]")
				If textNode Is Nothing Then textNode = resultNode.selectSingleNode("content[@language=""en-gb""]")
				If textNode Is Nothing Then
					'No result found.
				Else
					popupHelp = textNode.text
				End If
			End If
		End If
	End Function
	
	Private Function IsVistaOrAbove() As Boolean
#If debugging = 0 Then
		On Error Resume Next
#End If
		'UPGRADE_WARNING: Arrays in structure tOSV may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim tOSV As OSVERSIONINFO
		tOSV.dwVersionInfoSize = Len(tOSV)
		If GetVersionEx(tOSV) > 0 Then
			If (tOSV.dwMajorVersion > 5) Then
				IsVistaOrAbove = True
			End If
		End If
	End Function
End Module