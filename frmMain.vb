Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
	
	'1.0.1
	'   Fixed bug with XP Style and Large Fonts.
	'1.0.2
	'   Fixed always maximising
	'1.0.3
	'   I18N code added.
	'1.1.0
	'   Can now scale font in and out using Ctrl+ and Ctrl-
	'   Left-hand window shows file type, sort-of.
	'   Redid how startup works - no longer done in installer, which bugs the crap out of people
	
	Dim mFSO As Scripting.FileSystemObject
	Dim mFol As Scripting.Folder
	Dim mItems As Collection
	Dim mHistory(20) As String
	Private mFontChangedByUser As Boolean
	Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Integer, ByVal wBar As Integer, ByVal bShow As Integer) As Integer
	
	'Constants
	Private Const SB_HORZ As Short = 0 'Horizontal Scrollbar
	Private Const SB_VERT As Short = 1 'Vertical Scrollbbar
	Private Const SB_BOTH As Short = 3 'Both ScrollBars
	
	
	Private Sub cmdBack_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBack.Click
		On Error Resume Next
		Call mnuNavigateBack_Click(mnuNavigateBack, New System.EventArgs())
	End Sub
	
	Private Sub cmdGo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGo.Click
		On Error Resume Next
		Call lstDisplay_KeyPress(lstDisplay, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Return)))
	End Sub
	
	Private Sub cmdUp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUp.Click
		On Error Resume Next
		Call mnuNavigateUp_Click(mnuNavigateUp, New System.EventArgs())
	End Sub
	
	Private Sub frmMain_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		On Error Resume Next
		Dim path As String
		
		If KeyCode = System.Windows.Forms.Keys.Back Then
			KeyCode = 0
			path = GetFromHistory
			If path = "" Then
				Call Beep()
			Else
				mFol = mFSO.GetFolder(path)
				Call Display()
			End If
		ElseIf KeyCode = System.Windows.Forms.Keys.Up And (Shift And VB6.ShiftConstants.CtrlMask) = VB6.ShiftConstants.CtrlMask Then 
			KeyCode = 0
			Call mnuNavigateUp_Click(mnuNavigateUp, New System.EventArgs())
		ElseIf KeyCode = System.Windows.Forms.Keys.Home And (Shift And VB6.ShiftConstants.CtrlMask) = VB6.ShiftConstants.CtrlMask Then 
			KeyCode = 0
			Call mnuNavigateDesktop_Click(mnuNavigateDesktop, New System.EventArgs())
		End If
	End Sub
	
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error Resume Next
		Dim path As String
		
		Call modPath.DetermineSettingsPath("WebbIE", "Disk Explorer", "1")

		mFontChangedByUser = CBool(GetSettingIni("Disk Explorer", "Font", "ChangedByUser", "False"))
		If mFontChangedByUser Then
			lstDisplay.Font = VB6.FontChangeSize(lstDisplay.Font, CDec(GetSettingIni("Disk Explorer", "Font", "Size", CStr(lstDisplay.Font.SizeInPoints))))
		Else
            'TODO Call modLargeFonts.ApplySystemSettingsToForm(Me)
		End If
        mnuFileRunOrOpen.Text = mnuFileRunOrOpen.Text
        Me.mnuNavigateBack.Text = Me.mnuNavigateBack.Text
        Me.mnuNavigateUp.Text = Me.mnuNavigateUp.Text
        Me.mnuViewFontsize(0).Text = mnuViewFontsize(0).Text
        Me.mnuViewFontsize(1).Text = mnuViewFontsize(1).Text
		
		mnuOptionsStartupwithcomputer.Checked = CBool(modPath.GetSettingIni("Disk Explorer", "Program", "Startup", CStr(False)))
		Call ApplyStartup()
		
		mFSO = New Scripting.FileSystemObject
		If Len(VB.Command()) > 0 Then
			path = VB.Command()
		Else
			'Debug.Print modPath.GetSpecialFolderPath(modPath.CSIDL_PERSONAL)
			path = modPath.GetSettingIni("Disk Explorer", "History", "LastPath", modPath.GetSpecialFolderPath((modPath.CSIDL_PERSONAL)))
		End If
		mFol = mFSO.GetFolder(path)
		mnuViewAll.Checked = True
		Call Display()
	End Sub
	
	'UPGRADE_WARNING: Event frmMain.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmMain_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error Resume Next
		lstDisplay.Left = 0
		lstDisplay.Top = 0
		lstIcons.Width = VB6.TwipsToPixelsX(25 * lstDisplay.Font.SizeInPoints)
		lstIcons.Left = 0
		lstDisplay.Left = lstIcons.Width
		lstIcons.Top = 0
		lstDisplay.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - VB6.PixelsToTwipsX(lstIcons.Width))
		If Me.mnuViewTouchscreen.Checked Then
			cmdGo.Visible = True
			cmdBack.Visible = True
			cmdUp.Visible = True
			lstDisplay.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - VB6.PixelsToTwipsY(staMain.Height) - VB6.PixelsToTwipsY(cmdGo.Height))
			cmdGo.Left = 0
			cmdBack.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdGo.Left) + VB6.PixelsToTwipsX(cmdGo.Width))
			cmdUp.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdBack.Left) + VB6.PixelsToTwipsX(cmdBack.Width))
			cmdGo.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lstDisplay.Top) + VB6.PixelsToTwipsY(lstDisplay.Height))
			cmdBack.Top = cmdGo.Top
			cmdUp.Top = cmdGo.Top
		Else
			cmdGo.Visible = False
			cmdBack.Visible = False
			cmdUp.Visible = False
			lstDisplay.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - VB6.PixelsToTwipsY(staMain.Height))
		End If
		lstIcons.Height = lstDisplay.Height
	End Sub
	
	Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		On Error Resume Next
		Dim f As System.Windows.Forms.Form
		
        For Each f In My.Application.OpenForms
            If Not f Is Me Then
                Call f.Close()
            End If
        Next f
		If mFontChangedByUser Then
			Call SaveSettingIni("Disk Explorer", "Font", "Size", CStr(lstDisplay.Font.SizeInPoints))
		End If
		Call SaveSettingIni("Disk Explorer", "Program", "Startup", CStr(mnuOptionsStartupwithcomputer.Checked))
	End Sub
	
	Private Sub Display(Optional ByRef maintainListPosition As Boolean = False)
		On Error Resume Next
		Dim fo As Scripting.Folder
		Dim fi As Scripting.File
		Dim newItem As clsItem
		'UPGRADE_NOTE: name was upgraded to name_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim name_Renamed As String
		Dim d As Scripting.Drive
		Dim at As String
		Dim i As Integer
		Dim returnTo As Integer
		Dim fileCount As Integer
		Dim folderCount As Integer
        Dim addedTo As Long

		at = lstDisplay.Text
		Call lstDisplay.Items.Clear()
		Call lstIcons.Items.Clear()
		mItems = New Collection
		lstDisplay.Visible = False
        lstIcons.Visible = False
        lstDisplay.Sorted = True
		If mnuViewFolders.Checked Or mnuViewAll.Checked Then
			For	Each fo In mFol.SubFolders
				'Debug.Print fo.name & " " & fo.Attributes
				If (fo.Attributes And Scripting.__MIDL___MIDL_itf_scrrun_0000_0000_0001.Hidden) = 0 Then
                    folderCount = folderCount + 1
                    addedTo = lstDisplay.Items.Add(fo.Name & " " & GetText("Folder"))
                    Call lstIcons.Items.Insert(addedTo, "[]")
					newItem = New clsItem
					newItem.path = fo.path
                    newItem.itemType_Renamed = clsItem.itemType.FOLDER_TYPE
                    Call mItems.Add(newItem)
                    VB6.SetItemData(lstDisplay, addedTo, mItems.Count())
				End If
			Next fo
		End If
		If mnuViewFiles.Checked Or mnuViewAll.Checked Then
			For	Each fi In mFol.Files
				If (fi.Attributes And Scripting.__MIDL___MIDL_itf_scrrun_0000_0000_0001.Hidden) = 0 Then
					name_Renamed = fi.name
					fileCount = fileCount + 1
					If LCase(VB.Right(name_Renamed, 4)) = ".lnk" Then
						name_Renamed = Replace(name_Renamed, ".lnk", " - " & GetText("Shortcut"),  ,  , CompareMethod.Text)
						'                ElseIf LCase(Right(name, 4)) = ".txt" Then
						'                    name = Replace(name, ".txt", " - " & GetText("Text file"), , , vbTextCompare)
						'                ElseIf LCase(Right(name, 4)) = ".xls" Then
						'                    name = Replace(name, ".xls", " - " & GetText("Excel spreadsheet"), , , vbTextCompare)
						'                ElseIf LCase(Right(name, 4)) = ".xls" Or LCase(Right(name, 5)) = ".xlsx" Then
						'                    name = Replace(name, ".xlsx", " - " & GetText("Excel spreadsheet"), , , vbTextCompare)
						'                    name = Replace(name, ".xls", " - " & GetText("Excel spreadsheet"), , , vbTextCompare)
						'                ElseIf LCase(Right(name, 4)) = ".doc" Or LCase(Right(name, 5)) = ".docx" Then
						'                    name = Replace(name, ".docx", " - " & GetText("Word document"), , , vbTextCompare)
						'                    name = Replace(name, ".doc", " - " & GetText("Word document"), , , vbTextCompare)
					End If
                    addedTo = lstDisplay.Items.Add(name_Renamed)
                    Call lstIcons.Items.Insert(addedTo, " ")
					newItem = New clsItem
                    newItem.itemType_Renamed = clsItem.itemType.FILE_TYPE
					newItem.path = fi.path
                    Call mItems.Add(newItem)
					VB6.SetItemData(lstDisplay, addedTo, mItems.Count())
				End If
			Next fi
        End If
        lstDisplay.Sorted = False
		If mnuViewFolders.Checked Or mnuViewAll.Checked Then
			If mFol.ParentFolder Is Nothing Then
			Else
                Call lstDisplay.Items.Insert(0, GetText("Up a Folder"))
				Call lstIcons.Items.Insert(0, "^")
				newItem = New clsItem
                newItem.itemType_Renamed = clsItem.itemType.UP_FOLDER
				Call mItems.Add(newItem)
                VB6.SetItemData(lstDisplay, 0, mItems.Count())
			End If
		End If
		'Display current location.
		If mFol.name = "" Then
			Call lstDisplay.Items.Insert(0, GetText("In: Root (top) of") & " " & mFol.Drive.DriveLetter & " drive.")
			Call lstIcons.Items.Insert(0, ">")
		Else
			Call lstDisplay.Items.Insert(0, GetText("In:") & " " & mFol.name)
			Call lstIcons.Items.Insert(0, ">")
		End If
		newItem = New clsItem
        newItem.itemType_Renamed = clsItem.itemType.CURRENT_FOLDER
		Call mItems.Add(newItem)
        VB6.SetItemData(lstDisplay, 0, mItems.Count())
        'Desktop
        addedTo = lstDisplay.Items.Count
        Call lstDisplay.Items.Insert(lstDisplay.Items.Count, GetText("Go to Desktop"))
		Call lstIcons.Items.Insert(lstDisplay.Items.Count - 1, ">")
		newItem = New clsItem
        newItem.itemType_Renamed = clsItem.itemType.DESKTOP
		Call mItems.Add(newItem)
        VB6.SetItemData(lstDisplay, addedTo, mItems.Count())
        'My Documents
        addedTo = lstDisplay.Items.Count
		Call lstDisplay.Items.Insert(lstDisplay.Items.Count, GetText("Go to Documents"))
		Call lstIcons.Items.Insert(lstDisplay.Items.Count - 1, ">")
		newItem = New clsItem
        newItem.itemType_Renamed = clsItem.itemType.MY_DOCUMENTS
		Call mItems.Add(newItem)
		'UPGRADE_ISSUE: ListBox property lstDisplay.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
        VB6.SetItemData(lstDisplay, addedTo, mItems.Count())
		'Drive
		For	Each d In mFSO.Drives
			If d.DriveType = Scripting.__MIDL___MIDL_itf_scrrun_0001_0000_0001.Fixed Or d.DriveType = Scripting.__MIDL___MIDL_itf_scrrun_0001_0000_0001.CDRom Or d.DriveType = Scripting.__MIDL___MIDL_itf_scrrun_0001_0000_0001.Removable Then
				name_Renamed = GetText("Go to") & " " & d.DriveLetter & " drive "
				newItem = New clsItem
				newItem.path = d.path
				If d.DriveType = Scripting.__MIDL___MIDL_itf_scrrun_0001_0000_0001.Fixed Then
                    newItem.itemType_Renamed = clsItem.itemType.MAIN_DRIVE
					name_Renamed = name_Renamed & GetText("(hard disk)")
				ElseIf d.DriveType = Scripting.__MIDL___MIDL_itf_scrrun_0001_0000_0001.CDRom Then 
                    newItem.itemType_Renamed = clsItem.itemType.CD_DRIVE
					name_Renamed = name_Renamed & GetText("(CD drive)")
				ElseIf d.DriveType = Scripting.__MIDL___MIDL_itf_scrrun_0001_0000_0001.Removable Then 
                    newItem.itemType_Renamed = clsItem.itemType.USB_DRIVE
					name_Renamed = name_Renamed & GetText("(USB drive)")
				End If
                Call mItems.Add(newItem)
                addedTo = lstDisplay.Items.Count
				Call lstDisplay.Items.Insert(lstDisplay.Items.Count, name_Renamed)
				Call lstIcons.Items.Insert(lstDisplay.Items.Count - 1, "D")
				'UPGRADE_ISSUE: ListBox property lstDisplay.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
                VB6.SetItemData(lstDisplay, addedTo, mItems.Count())
			End If
		Next d
		
		If mFol.name = "" Then
			Me.Text = GetText("Disk Explorer") & " - " & mFol.Drive.Path & " " & GetText("drive") & " - " & fileCount & " " & IIf(fileCount <> 1, GetText("files"), GetText("file")) & ", " & folderCount & " " & IIf(folderCount <> 1, GetText("folders"), GetText("folder")) & "."
		Else
			Me.Text = GetText("Disk Explorer") & " - " & mFol.name & " - " & fileCount & " " & GetText("files,") & " " & folderCount & " " & GetText("folders.")
		End If
		lstDisplay.Visible = True
		lstIcons.Visible = True
		Me.staMain.Text = mFol.path
		If maintainListPosition Then
			returnTo = -1
			For i = 0 To lstDisplay.Items.Count - 1
				If VB6.GetItemString(lstDisplay, i) = at Then
					returnTo = i
					Exit For
				End If
			Next i
			If returnTo > -1 Then
				lstDisplay.SelectedIndex = returnTo
			Else
				lstDisplay.SelectedIndex = 0
			End If
		Else
			lstDisplay.SelectedIndex = 0
		End If
		If Not Me.Visible Then Call Me.Show()
		Call AddToHistory((mFol.path))
		lstDisplay.Focus()
	End Sub
	
	Private Sub AddToHistory(ByRef path As String)
		On Error Resume Next
		Dim i As Integer
		
		For i = UBound(mHistory) To 1 Step -1
			mHistory(i) = mHistory(i - 1)
		Next i
		mHistory(0) = path
		
		Debug.Print("ADD")
		For i = 0 To 20
			Debug.Print(i & " " & mHistory(i))
		Next i
	End Sub
	
	Private Function GetFromHistory() As String
		On Error Resume Next
		Dim i As Integer
		
		GetFromHistory = mHistory(1)
		For i = 1 To UBound(mHistory)
			mHistory(i - 1) = mHistory(i)
		Next i
		For i = 1 To UBound(mHistory)
			mHistory(i - 1) = mHistory(i)
		Next i
		mHistory(UBound(mHistory) - 1) = ""
		mHistory(UBound(mHistory)) = ""
		
		Debug.Print("GET")
		For i = 0 To 20
			Debug.Print(i & " " & mHistory(i))
		Next i
	End Function
	
	Private Sub LaunchWrite(ByRef path As String)
		On Error Resume Next
		Dim writePath As String
		
		writePath = modPath.GetSpecialFolderPath((modPath.CSIDL_WINDOWS))
	End Sub
	
	Public Sub mnuCommandsCreatefolder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCommandsCreatefolder.Click
		On Error Resume Next
		'UPGRADE_NOTE: name was upgraded to name_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim name_Renamed As String
		
		name_Renamed = InputBox(GetText("Enter the name of the new folder or directory:"))
		If Len(name_Renamed) > 0 Then
			On Error GoTo CreationError
			Call mFSO.CreateFolder(mFol.path & "\" & name_Renamed)
		End If
		Call Display(True)
		Exit Sub
CreationError: 
		MsgBox(GetText("Failed to create folder:") & " " & Err.Description, MsgBoxStyle.Exclamation)
		On Error Resume Next
	End Sub
	
	Public Sub mnuCommandsDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCommandsDelete.Click
		On Error Resume Next
		Dim item As clsItem
		Dim fo As Scripting.Folder
		Dim fi As Scripting.File
		Dim result As Integer
		
		item = mItems.Item(VB6.GetItemData(lstDisplay, lstDisplay.SelectedIndex))
        Select Case item.itemType_Renamed
            Case clsItem.itemType.FILE_TYPE
                fi = mFSO.GetFile(item.path)
                If (fi.Attributes And Scripting.__MIDL___MIDL_itf_scrrun_0000_0000_0001.System) > 0 Then
                    'System file
                    MsgBox(GetText("You cannot delete system files. Use Windows Explorer instead."), MsgBoxStyle.Exclamation)
                Else
                    fo = fi.ParentFolder
                    If (fo.Attributes And Scripting.__MIDL___MIDL_itf_scrrun_0000_0000_0001.System) > 0 Then
                        MsgBox(GetText("You cannot delete files in system folders. Use Windows Explorer instead."), MsgBoxStyle.Exclamation)
                    Else
                        result = DeleteFileToRecycleBin((item.path))
                        If result = ERROR_OK Then
                            Call Display(True)
                        Else
                            MsgBox(GetText("Sorry, failed to delete file for unknown reason (error code") & " " & result & ")", MsgBoxStyle.Exclamation)
                            Call Display(True)
                        End If
                    End If
                End If
            Case clsItem.itemType.FOLDER_TYPE
                fo = mFSO.GetFolder(item.path)
                If (fo.Attributes And Scripting.__MIDL___MIDL_itf_scrrun_0000_0000_0001.System) > 0 Then
                    'System file
                    MsgBox(GetText("You cannot delete system folders. Use Windows Explorer instead."), MsgBoxStyle.Exclamation)
                Else
                    result = DeleteFileToRecycleBin((item.path))
                    If result = ERROR_OK Then
                        Call Display(True)
                    Else
                        MsgBox(GetText("Sorry, failed to delete file for unknown reason (error code") & " " & result & ")", MsgBoxStyle.Exclamation)
                        Call Display(True)
                    End If
                End If
            Case Else
                Call Beep()
        End Select
		
	End Sub
	
	Public Sub mnuCommandsOpeninexplorer_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCommandsOpeninexplorer.Click
		On Error Resume Next
		Call Shell("explorer """ & mFol.path & """", AppWinStyle.NormalFocus)
	End Sub
	
	Public Sub mnuFileExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileExit.Click
		On Error Resume Next
		Call Me.Close()
	End Sub
	
	Public Sub mnuFileRunOrOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileRunOrOpen.Click
		On Error Resume Next
		Call lstDisplay_KeyPress(lstDisplay, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Return)))
	End Sub
	
	Public Sub mnuHelpAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpAbout.Click
		On Error Resume Next
		MsgBox(My.Application.Info.Title & vbTab & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision & vbNewLine & "Package Version" & vbTab & modVersion.GetPackageVersion & vbNewLine & "Alasdair King, http://www.webbie.org.uk", MsgBoxStyle.Information + MsgBoxStyle.OKOnly)
	End Sub
	
	Public Sub mnuHelpManual_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpManual.Click
		On Error Resume Next
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
        frmHelp = New frmHelp
		frmHelp.Icon = Me.Icon
		Call VB6.ShowForm(frmHelp, VB6.FormShowConstants.Modal, Me)
	End Sub
	
	Public Sub mnuNavigateBack_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateBack.Click
		On Error Resume Next
		Call frmMain_KeyDown(Me, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Back Or 0 * &H10000))
	End Sub
	
	Public Sub mnuNavigateDesktop_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateDesktop.Click
		On Error Resume Next
		Dim path As String
		path = modPath.GetSpecialFolderPath((modPath.CSIDL_DESKTOP))
		mFol = mFSO.GetFolder(path)
		Call Display()
	End Sub
	
	Public Sub mnuNavigateUp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateUp.Click
		On Error Resume Next
		If mFol.ParentFolder Is Nothing Then
			Call Beep()
		Else
			mFol = mFol.ParentFolder
			Call Display()
		End If
	End Sub
	
	Public Sub mnuOptionsStartupwithcomputer_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOptionsStartupwithcomputer.Click
		mnuOptionsStartupwithcomputer.Checked = Not mnuOptionsStartupwithcomputer.Checked
		Call ApplyStartup()
	End Sub
	
    Private Sub ApplyStartup()
        Dim regKey As Microsoft.Win32.RegistryKey
        regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
        If mnuOptionsStartupwithcomputer.Checked Then
            'Turn on.
            regKey.SetValue("WebbIE Disk Explorer", My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".exe")
        Else
            'Turn off.
            regKey.DeleteValue("WebbIE Disk Explorer")
        End If
        regKey.Close()
    End Sub
	
	Public Sub mnuViewAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewAll.Click
		On Error Resume Next
		mnuViewFolders.Checked = False
		mnuViewFiles.Checked = False
		mnuViewAll.Checked = True
		Call Display(True)
	End Sub
	
	Public Sub mnuViewFiles_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewFiles.Click
		On Error Resume Next
		mnuViewFolders.Checked = False
		mnuViewFiles.Checked = True
		mnuViewAll.Checked = False
		Call Display(True)
	End Sub
	
	Public Sub mnuViewFolders_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewFolders.Click
		On Error Resume Next
		mnuViewFolders.Checked = True
		mnuViewFiles.Checked = False
		mnuViewAll.Checked = False
		Call Display(True)
	End Sub
	
	Public Sub mnuViewFontsize_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewFontsize.Click
		Dim index As Short = mnuViewFontsize.GetIndex(eventSender)
		On Error Resume Next
		If index = 0 Then
			'Increase font size
			lstDisplay.Font = VB6.FontChangeSize(lstDisplay.Font, lstDisplay.Font.SizeInPoints + 1)
			Call SaveSettingIni("Disk Explorer", "Font", "ChangedByUser", CStr(True))
		Else
			'Decrease font size
			lstDisplay.Font = VB6.FontChangeSize(lstDisplay.Font, lstDisplay.Font.SizeInPoints - 1)
			Call SaveSettingIni("Disk Explorer", "Font", "ChangedByUser", CStr(True))
		End If
	End Sub
	
	Public Sub mnuViewTouchscreen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewTouchscreen.Click
		On Error Resume Next
		mnuViewTouchscreen.Checked = Not mnuViewTouchscreen.Checked
		Call frmMain_Resize(Me, New System.EventArgs())
	End Sub
	
	Private Sub tmrUpdateIcons_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrUpdateIcons.Tick
        On Error Resume Next
        If lstIcons.Font.Size <> lstDisplay.Font.Size Then
            lstIcons.Font = VB6.FontChangeSize(lstIcons.Font, lstDisplay.Font.SizeInPoints)
        End If
        If lstIcons.TopIndex <> lstDisplay.TopIndex Then
            lstIcons.TopIndex = lstDisplay.TopIndex
        End If
        ShowScrollBar(lstIcons.Handle.ToInt32, SB_VERT, False)
    End Sub

    Private Sub lstDisplay_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstDisplay.DoubleClick
        On Error Resume Next
        Call lstDisplay_KeyPress(lstDisplay, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Return)))
    End Sub

    Private Sub lstDisplay_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lstDisplay.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        On Error Resume Next
        If ((KeyCode = System.Windows.Forms.Keys.Down) Or (KeyCode = System.Windows.Forms.Keys.PageDown)) And lstDisplay.SelectedIndex = lstDisplay.Items.Count - 1 Then
            Call Beep()
        ElseIf ((KeyCode = System.Windows.Forms.Keys.Up) Or (KeyCode = System.Windows.Forms.Keys.PageUp)) And lstDisplay.SelectedIndex = 0 Then
            Call Beep()
        End If
    End Sub

    Private Sub lstDisplay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles lstDisplay.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error Resume Next
        Dim path As String
        Dim extension As String
        Dim item As clsItem
        Dim f As Scripting.File

        If KeyAscii = System.Windows.Forms.Keys.Return Then
            item = mItems.Item(VB6.GetItemData(lstDisplay, lstDisplay.SelectedIndex))
            Select Case item.itemType_Renamed
                Case clsItem.itemType.CURRENT_FOLDER
                    MsgBox(mFol.Path, MsgBoxStyle.Information, GetText("Current Folder"))
                Case clsItem.itemType.UP_FOLDER
                    If mFol.ParentFolder Is Nothing Then
                        Call Beep()
                    Else
                        mFol = mFol.ParentFolder
                        Call Display()
                    End If
                Case clsItem.itemType.FOLDER_TYPE
                    mFol = mFSO.GetFolder(item.path)
                    Call Display()
                Case clsItem.itemType.FILE_TYPE
                    path = item.path
                    f = mFSO.GetFile(path)
                    Call modShellExec.ShellExecute(0, "open", path, "", f.ParentFolder.Path, modShellExec.SW_NORMAL)
                Case clsItem.itemType.DESKTOP
                    path = modPath.GetSpecialFolderPath((modPath.CSIDL_DESKTOP))
                    mFol = mFSO.GetFolder(path)
                    Call Display()
                Case clsItem.itemType.MY_DOCUMENTS
                    path = modPath.GetSpecialFolderPath((modPath.CSIDL_PERSONAL))
                    mFol = mFSO.GetFolder(path)
                    Call Display()
                Case clsItem.itemType.MAIN_DRIVE
                    mFol = mFSO.GetFolder(item.path)
                    Call Display()
                Case clsItem.itemType.CD_DRIVE
                    On Error GoTo NoDiskError
                    mFol = mFSO.GetFolder(item.path)
                    On Error Resume Next
                    Call Display()
                Case Else
                    MsgBox("Not implemented yet")
            End Select
        End If
        GoTo EventExitSub
NoDiskError:
        MsgBox(GetText("There is no CD-ROM or audio CD in the CD drive."), MsgBoxStyle.Exclamation)
        Resume Next
EventExitSub:
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub lstDisplay_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lstDisplay.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        If (Shift And VB6.ShiftConstants.CtrlMask) > 0 Then
            If KeyCode = 187 Or KeyCode = System.Windows.Forms.Keys.Add Then
                Call mnuViewFontsize_Click(mnuViewFontsize.Item(0), New System.EventArgs())
            ElseIf KeyCode = 189 Or KeyCode = System.Windows.Forms.Keys.Subtract Then
                Call mnuViewFontsize_Click(mnuViewFontsize.Item(1), New System.EventArgs())
            End If
        ElseIf (Shift And VB6.ShiftConstants.AltMask) > 0 Then
            If KeyCode = System.Windows.Forms.Keys.Up Then
                Call mnuNavigateUp_Click(mnuNavigateUp, New System.EventArgs())
            ElseIf KeyCode = System.Windows.Forms.Keys.Left Then
                Call mnuNavigateBack_Click(mnuNavigateBack, New System.EventArgs())
            End If
        End If
    End Sub
End Class