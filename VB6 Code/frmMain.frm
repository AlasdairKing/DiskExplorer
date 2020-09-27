VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Explorer"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   810
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdateIcons 
      Interval        =   50
      Left            =   1800
      Top             =   1320
   End
   Begin VB.ListBox lstIcons 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1800
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2715
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstDisplay 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      Caption         =   "Files"
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   405
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileRunOrOpen 
         Caption         =   "&Run or Open"
      End
      Begin VB.Menu mnuCommandsDelete 
         Caption         =   "&Delete file or folder"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuCommandsOpeninexplorer 
         Caption         =   "Open in &Explorer"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuCommandsCreatefolder 
         Caption         =   "&Create folder (directory)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuNavigate 
      Caption         =   "&Navigate"
      Begin VB.Menu mnuNavigateBack 
         Caption         =   "&Back"
      End
      Begin VB.Menu mnuNavigateUp 
         Caption         =   "&Up"
      End
      Begin VB.Menu mnuNavigateDesktop 
         Caption         =   "&Desktop (Home)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewFolders 
         Caption         =   "F&olders only"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuViewFiles 
         Caption         =   "F&iles only"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuViewAll 
         Caption         =   "&Both files and folders"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTouchscreen 
         Caption         =   "&Touchscreen"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewFontsize 
         Caption         =   "Increase font size"
         Index           =   0
      End
      Begin VB.Menu mnuViewFontsize 
         Caption         =   "Decrease font size"
         Index           =   1
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsStartupwithcomputer 
         Caption         =   "&Startup with computer"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpManual 
         Caption         =   "&Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

'Constants
Private Const SB_HORZ = 0 'Horizontal Scrollbar
Private Const SB_VERT = 1 'Vertical Scrollbbar
Private Const SB_BOTH = 3 'Both ScrollBars


Private Sub cmdBack_Click()
    On Error Resume Next
    Call mnuNavigateBack_Click
End Sub

Private Sub cmdGo_Click()
    On Error Resume Next
    Call lstDisplay_KeyPress(vbKeyReturn)
End Sub

Private Sub cmdUp_Click()
    On Error Resume Next
    Call mnuNavigateUp_Click
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim path As String
    
    If KeyCode = vbKeyBack Then
        KeyCode = 0
        path = GetFromHistory
        If path = "" Then
            Call Beep
        Else
            Set mFol = mFSO.GetFolder(path)
            Call Display
        End If
    ElseIf KeyCode = vbKeyUp And (Shift And vbCtrlMask) = vbCtrlMask Then
        KeyCode = 0
        Call mnuNavigateUp_Click
    ElseIf KeyCode = vbKeyHome And (Shift And vbCtrlMask) = vbCtrlMask Then
        KeyCode = 0
        Call mnuNavigateDesktop_Click
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim path As String
    
    Call modPath.DetermineSettingsPath("WebbIE", "Disk Explorer", "1")
    Call modRememberPosition.LoadPosition(Me)
    Call modI18N.ApplyUILanguageToThisForm(Me)
    
    mFontChangedByUser = CBool(GetSettingIni("Disk Explorer", "Font", "ChangedByUser", "False"))
    If mFontChangedByUser Then
        lstDisplay.FontSize = GetSettingIni("Disk Explorer", "Font", "Size", lstDisplay.FontSize)
    Else
        Call modLargeFonts.ApplySystemSettingsToForm(Me)
    End If
    mnuFileRunOrOpen.Caption = mnuFileRunOrOpen.Caption & vbTab & GetText("Enter")
    Me.mnuNavigateBack.Caption = Me.mnuNavigateBack.Caption & vbTab & GetText("Backspace")
    Me.mnuNavigateUp.Caption = Me.mnuNavigateUp.Caption & vbTab & GetText("Ctrl+Up")
    Me.mnuViewFontsize(0).Caption = mnuViewFontsize(0).Caption & vbTab & "Ctrl+Plus"
    Me.mnuViewFontsize(1).Caption = mnuViewFontsize(1).Caption & vbTab & "Ctrl+Minus"
    
    mnuOptionsStartupwithcomputer.Checked = CBool(modPath.GetSettingIni("Disk Explorer", "Program", "Startup", False))
    Call ApplyStartup
    
    Set mFSO = New FileSystemObject
    If Len(Command$) > 0 Then
        path = Command$
    Else
        'Debug.Print modPath.GetSpecialFolderPath(modPath.CSIDL_PERSONAL)
        path = modPath.GetSettingIni("Disk Explorer", "History", "LastPath", modPath.GetSpecialFolderPath(modPath.CSIDL_PERSONAL))
    End If
    Set mFol = mFSO.GetFolder(path)
    mnuViewAll.Checked = True
    Call Display
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstDisplay.Left = 0
    lstDisplay.Top = 0
    lstIcons.Width = 25 * lstDisplay.FontSize
    lstIcons.Left = 0
    lstDisplay.Left = lstIcons.Width
    lstIcons.Top = 0
    lstDisplay.Width = Me.ScaleWidth - lstIcons.Width
    If Me.mnuViewTouchscreen.Checked Then
        cmdGo.Visible = True
        cmdBack.Visible = True
        cmdUp.Visible = True
        lstDisplay.Height = Me.ScaleHeight - staMain.Height - cmdGo.Height
        cmdGo.Left = 0
        cmdBack.Left = cmdGo.Left + cmdGo.Width
        cmdUp.Left = cmdBack.Left + cmdBack.Width
        cmdGo.Top = lstDisplay.Top + lstDisplay.Height
        cmdBack.Top = cmdGo.Top
        cmdUp.Top = cmdGo.Top
    Else
        cmdGo.Visible = False
        cmdBack.Visible = False
        cmdUp.Visible = False
        lstDisplay.Height = Me.ScaleHeight - staMain.Height
    End If
    lstIcons.Height = lstDisplay.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim f As Form
    
    Call modRememberPosition.SavePosition(Me)
    For Each f In Forms
        If Not f Is Me Then
            Call Unload(f)
        End If
    Next f
    If mFontChangedByUser Then
        Call SaveSettingIni("Disk Explorer", "Font", "Size", lstDisplay.FontSize)
    End If
    Call SaveSettingIni("Disk Explorer", "Program", "Startup", mnuOptionsStartupwithcomputer.Checked)
End Sub

Private Sub Display(Optional maintainListPosition As Boolean = False)
    On Error Resume Next
    Dim fo As Folder
    Dim fi As File
    Dim newItem As clsItem
    Dim name As String
    Dim d As Drive
    Dim at As String
    Dim i As Long
    Dim returnTo As Long
    Dim fileCount As Long
    Dim folderCount As Long
    
    at = lstDisplay.Text
    Call lstDisplay.Clear
    Call lstIcons.Clear
    Set mItems = New Collection
    lstDisplay.Visible = False
    lstIcons.Visible = False
    If mnuViewFolders.Checked Or mnuViewAll.Checked Then
        For Each fo In mFol.SubFolders
            'Debug.Print fo.name & " " & fo.Attributes
            If (fo.Attributes And Hidden) = 0 Then
                folderCount = folderCount + 1
                Call lstDisplay.AddItem(fo.name & " " & GetText("Folder"))
                Call lstIcons.AddItem("[]", lstDisplay.NewIndex)
                Set newItem = New clsItem
                newItem.path = fo.path
                newItem.itemType = FOLDER_TYPE
                Call mItems.Add(newItem)
                lstDisplay.ItemData(lstDisplay.NewIndex) = mItems.Count
            End If
        Next fo
    End If
    If mnuViewFiles.Checked Or mnuViewAll.Checked Then
        For Each fi In mFol.Files
            If (fi.Attributes And Hidden) = 0 Then
                name = fi.name
                fileCount = fileCount + 1
                If LCase(Right(name, 4)) = ".lnk" Then
                    name = Replace(name, ".lnk", " - " & GetText("Shortcut"), , , vbTextCompare)
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
                Call lstDisplay.AddItem(name)
                Call lstIcons.AddItem(" ", lstDisplay.NewIndex)
                Set newItem = New clsItem
                newItem.itemType = FILE_TYPE
                newItem.path = fi.path
                Call mItems.Add(newItem)
                lstDisplay.ItemData(lstDisplay.NewIndex) = mItems.Count
            End If
        Next fi
    End If
    If mnuViewFolders.Checked Or mnuViewAll.Checked Then
        If mFol.ParentFolder Is Nothing Then
        Else
            Call lstDisplay.AddItem(GetText("Up a Folder"), 0)
            Call lstIcons.AddItem("^", 0)
            Set newItem = New clsItem
            newItem.itemType = UP_FOLDER
            Call mItems.Add(newItem)
            lstDisplay.ItemData(lstDisplay.NewIndex) = mItems.Count
        End If
    End If
    'Display current location.
    If mFol.name = "" Then
        Call lstDisplay.AddItem(GetText("In: Root (top) of") & " " & mFol.Drive.DriveLetter & " drive.", 0)
        Call lstIcons.AddItem(">", 0)
    Else
        Call lstDisplay.AddItem(GetText("In:") & " " & mFol.name, 0)
        Call lstIcons.AddItem(">", 0)
    End If
    Set newItem = New clsItem
    newItem.itemType = CURRENT_FOLDER
    Call mItems.Add(newItem)
    lstDisplay.ItemData(lstDisplay.NewIndex) = mItems.Count
    'Desktop
    Call lstDisplay.AddItem(GetText("Go to Desktop"), lstDisplay.ListCount)
    Call lstIcons.AddItem(">", lstDisplay.ListCount - 1)
    Set newItem = New clsItem
    newItem.itemType = DESKTOP
    Call mItems.Add(newItem)
    lstDisplay.ItemData(lstDisplay.NewIndex) = mItems.Count
    'My Documents
    Call lstDisplay.AddItem(GetText("Go to Documents"), lstDisplay.ListCount)
    Call lstIcons.AddItem(">", lstDisplay.ListCount - 1)
    Set newItem = New clsItem
    newItem.itemType = MY_DOCUMENTS
    Call mItems.Add(newItem)
    lstDisplay.ItemData(lstDisplay.NewIndex) = mItems.Count
    'Drive
    For Each d In mFSO.Drives
        If d.DriveType = Fixed Or d.DriveType = CDRom Or d.DriveType = Removable Then
            name = GetText("Go to") & " " & d.DriveLetter & " drive "
            Set newItem = New clsItem
            newItem.path = d.path
            If d.DriveType = Fixed Then
                newItem.itemType = MAIN_DRIVE
                name = name & GetText("(hard disk)")
            ElseIf d.DriveType = CDRom Then
                newItem.itemType = CD_DRIVE
                name = name & GetText("(CD drive)")
            ElseIf d.DriveType = Removable Then
                newItem.itemType = USB_DRIVE
                name = name & GetText("(USB drive)")
            End If
            Call mItems.Add(newItem)
            Call lstDisplay.AddItem(name, lstDisplay.ListCount)
            Call lstIcons.AddItem("D", lstDisplay.ListCount - 1)
            lstDisplay.ItemData(lstDisplay.NewIndex) = mItems.Count
        End If
    Next d
    
    If mFol.name = "" Then
        Me.Caption = GetText("Disk Explorer") & " - " & mFol.Drive & " " & GetText("drive") & " - " & fileCount & _
            " " & IIf(fileCount <> 1, GetText("files"), GetText("file")) & ", " & folderCount & _
            " " & IIf(folderCount <> 1, GetText("folders"), GetText("folder")) & "."
    Else
        Me.Caption = GetText("Disk Explorer") & " - " & mFol.name & " - " & fileCount & " " & GetText("files,") & " " & folderCount & " " & GetText("folders.")
    End If
    lstDisplay.Visible = True
    lstIcons.Visible = True
    Me.staMain.SimpleText = mFol.path
    If maintainListPosition Then
        returnTo = -1
        For i = 0 To lstDisplay.ListCount - 1
            If lstDisplay.List(i) = at Then
                returnTo = i
                Exit For
            End If
        Next i
        If returnTo > -1 Then
            lstDisplay.ListIndex = returnTo
        Else
            lstDisplay.ListIndex = 0
        End If
    Else
        lstDisplay.ListIndex = 0
    End If
    If Not Me.Visible Then Call Me.Show
    Call AddToHistory(mFol.path)
    lstDisplay.SetFocus
End Sub

Private Sub AddToHistory(path As String)
    On Error Resume Next
    Dim i As Long
    
    For i = UBound(mHistory) To 1 Step -1
        mHistory(i) = mHistory(i - 1)
    Next i
    mHistory(0) = path
    
    Debug.Print "ADD"
    For i = 0 To 20
        Debug.Print i & " " & mHistory(i)
    Next i
End Sub

Private Function GetFromHistory() As String
    On Error Resume Next
    Dim i As Long
    
    GetFromHistory = mHistory(1)
    For i = 1 To UBound(mHistory)
        mHistory(i - 1) = mHistory(i)
    Next i
    For i = 1 To UBound(mHistory)
        mHistory(i - 1) = mHistory(i)
    Next i
    mHistory(UBound(mHistory) - 1) = ""
    mHistory(UBound(mHistory)) = ""
    
    Debug.Print "GET"
    For i = 0 To 20
        Debug.Print i & " " & mHistory(i)
    Next i
End Function

Private Sub lstDisplay_DblClick()
    On Error Resume Next
    Call lstDisplay_KeyPress(vbKeyReturn)
End Sub

Private Sub lstDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If ((KeyCode = vbKeyDown) Or (KeyCode = vbKeyPageDown)) And lstDisplay.ListIndex = lstDisplay.ListCount - 1 Then
        Call Beep
    ElseIf ((KeyCode = vbKeyUp) Or (KeyCode = vbKeyPageUp)) And lstDisplay.ListIndex = 0 Then
        Call Beep
    End If
End Sub

Private Sub lstDisplay_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim path As String
    Dim extension As String
    Dim item As clsItem
    Dim f As File
    
    If KeyAscii = vbKeyReturn Then
        Set item = mItems.item(lstDisplay.ItemData(lstDisplay.ListIndex))
        Select Case item.itemType
            Case CURRENT_FOLDER
                MsgBox mFol.path, vbInformation, GetText("Current Folder")
            Case UP_FOLDER
                If mFol.ParentFolder Is Nothing Then
                    Call Beep
                Else
                    Set mFol = mFol.ParentFolder
                    Call Display
                End If
            Case FOLDER_TYPE
                Set mFol = mFSO.GetFolder(item.path)
                Call Display
            Case FILE_TYPE
                path = item.path
                Set f = mFSO.GetFile(path)
                Call modShellExec.ShellExecute(0&, "open", path, "", f.ParentFolder, modShellExec.SW_NORMAL)
            Case DESKTOP
                path = modPath.GetSpecialFolderPath(modPath.CSIDL_DESKTOP)
                Set mFol = mFSO.GetFolder(path)
                Call Display
            Case MY_DOCUMENTS
                path = modPath.GetSpecialFolderPath(modPath.CSIDL_PERSONAL)
                Set mFol = mFSO.GetFolder(path)
                Call Display
            Case MAIN_DRIVE
                Set mFol = mFSO.GetFolder(item.path)
                Call Display
            Case CD_DRIVE
                On Error GoTo NoDiskError:
                Set mFol = mFSO.GetFolder(item.path)
                On Error Resume Next
                Call Display
            Case Else
                MsgBox "Not implemented yet"
        End Select
    End If
    Exit Sub
NoDiskError:
    MsgBox GetText("There is no CD-ROM or audio CD in the CD drive."), vbExclamation
    Resume Next
End Sub

Private Sub LaunchWrite(path As String)
    On Error Resume Next
    Dim writePath As String
    
    writePath = modPath.GetSpecialFolderPath(modPath.CSIDL_WINDOWS)
End Sub

Private Sub lstDisplay_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        If KeyCode = 187 Or KeyCode = vbKeyAdd Then
            Call mnuViewFontsize_Click(0)
        ElseIf KeyCode = 189 Or KeyCode = vbKeySubtract Then
            Call mnuViewFontsize_Click(1)
        End If
    ElseIf (Shift And vbAltMask) > 0 Then
        If KeyCode = vbKeyUp Then
            Call mnuNavigateUp_Click
        ElseIf KeyCode = vbKeyLeft Then
            Call mnuNavigateBack_Click
        End If
    End If
End Sub

Private Sub mnuCommandsCreatefolder_Click()
    On Error Resume Next
    Dim name As String
    
    name = InputBox(GetText("Enter the name of the new folder or directory:"))
    If Len(name) > 0 Then
        On Error GoTo CreationError
        Call mFSO.CreateFolder(mFol.path & "\" & name)
    End If
    Call Display(True)
    Exit Sub
CreationError:
    MsgBox GetText("Failed to create folder:") & " " & Err.Description, vbExclamation
    On Error Resume Next
End Sub

Private Sub mnuCommandsDelete_Click()
    On Error Resume Next
    Dim item As clsItem
    Dim fo As Folder
    Dim fi As File
    Dim result As Long
    
    Set item = mItems.item(lstDisplay.ItemData(lstDisplay.ListIndex))
    Select Case item.itemType
        Case FILE_TYPE
            Set fi = mFSO.GetFile(item.path)
            If (fi.Attributes And System) > 0 Then
                'System file
                MsgBox GetText("You cannot delete system files. Use Windows Explorer instead."), vbExclamation
            Else
                Set fo = fi.ParentFolder
                If (fo.Attributes And System) > 0 Then
                    MsgBox GetText("You cannot delete files in system folders. Use Windows Explorer instead."), vbExclamation
                Else
                    result = DeleteFileToRecycleBin(item.path)
                    If result = ERROR_OK Then
                        Call Display(True)
                    Else
                        MsgBox GetText("Sorry, failed to delete file for unknown reason (error code") & " " & result & ")", vbExclamation
                        Call Display(True)
                    End If
                End If
            End If
        Case FOLDER_TYPE
            Set fo = mFSO.GetFolder(item.path)
            If (fo.Attributes And System) > 0 Then
                'System file
                MsgBox GetText("You cannot delete system folders. Use Windows Explorer instead."), vbExclamation
            Else
                result = DeleteFileToRecycleBin(item.path)
                If result = ERROR_OK Then
                    Call Display(True)
                Else
                    MsgBox GetText("Sorry, failed to delete file for unknown reason (error code") & " " & result & ")", vbExclamation
                    Call Display(True)
                End If
            End If
        Case Else
            Call Beep
    End Select
        
End Sub

Private Sub mnuCommandsOpeninexplorer_Click()
    On Error Resume Next
    Call Shell("explorer """ & mFol.path & """", vbNormalFocus)
End Sub

Private Sub mnuFileExit_Click()
    On Error Resume Next
    Call Unload(Me)
End Sub

Private Sub mnuFileRunOrOpen_Click()
    On Error Resume Next
    Call lstDisplay_KeyPress(vbKeyReturn)
End Sub

Private Sub mnuHelpAbout_Click()
    On Error Resume Next
    MsgBox App.Title & vbTab & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & "Package Version" & vbTab & modVersion.GetPackageVersion & vbNewLine & "Alasdair King, http://www.webbie.org.uk", vbInformation + vbOKOnly
End Sub

Private Sub mnuHelpManual_Click()
    On Error Resume Next
    Call Load(frmHelp)
    frmHelp.Icon = Me.Icon
    Call frmHelp.Show(vbModal, Me)
End Sub

Private Sub mnuNavigateBack_Click()
    On Error Resume Next
    Call Form_KeyDown(vbKeyBack, 0)
End Sub

Private Sub mnuNavigateDesktop_Click()
    On Error Resume Next
    Dim path As String
    path = modPath.GetSpecialFolderPath(modPath.CSIDL_DESKTOP)
    Set mFol = mFSO.GetFolder(path)
    Call Display
End Sub

Private Sub mnuNavigateUp_Click()
    On Error Resume Next
    If mFol.ParentFolder Is Nothing Then
        Call Beep
    Else
        Set mFol = mFol.ParentFolder
        Call Display
    End If
End Sub

Private Sub mnuOptionsStartupwithcomputer_Click()
    mnuOptionsStartupwithcomputer.Checked = Not mnuOptionsStartupwithcomputer.Checked
    Call ApplyStartup
End Sub

Private Sub ApplyStartup()
    If mnuOptionsStartupwithcomputer.Checked Then
        'Turn on.
        Call modIni.SetRegValue(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WebbIE Disk Explorer", App.path & "\" & App.EXEName & ".exe")
    Else
        'Turn off.
        Call modIni.SetRegValue(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WebbIE Disk Explorer", "")
    End If
End Sub

Private Sub mnuViewAll_Click()
    On Error Resume Next
    mnuViewFolders.Checked = False
    mnuViewFiles.Checked = False
    mnuViewAll.Checked = True
    Call Display(True)
End Sub

Private Sub mnuViewFiles_Click()
    On Error Resume Next
    mnuViewFolders.Checked = False
    mnuViewFiles.Checked = True
    mnuViewAll.Checked = False
    Call Display(True)
End Sub

Private Sub mnuViewFolders_Click()
    On Error Resume Next
    mnuViewFolders.Checked = True
    mnuViewFiles.Checked = False
    mnuViewAll.Checked = False
    Call Display(True)
End Sub

Private Sub mnuViewFontsize_Click(index As Integer)
    On Error Resume Next
    If index = 0 Then
        'Increase font size
        lstDisplay.Font.size = lstDisplay.Font.size + 1
        Call SaveSettingIni("Disk Explorer", "Font", "ChangedByUser", True)
    Else
        'Decrease font size
        lstDisplay.Font.size = lstDisplay.Font.size - 1
        Call SaveSettingIni("Disk Explorer", "Font", "ChangedByUser", True)
    End If
End Sub

Private Sub mnuViewTouchscreen_Click()
    On Error Resume Next
    mnuViewTouchscreen.Checked = Not mnuViewTouchscreen.Checked
    Call Form_Resize
End Sub

Private Sub tmrUpdateIcons_Timer()
    On Error Resume Next
    lstIcons.FontSize = lstDisplay.FontSize
    lstIcons.TopIndex = lstDisplay.TopIndex
    ShowScrollBar lstIcons.hWnd, SB_VERT, False
End Sub
