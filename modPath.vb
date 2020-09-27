Option Strict Off
Option Explicit On
Module modPath
	
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
	
	'Changelog
	'       28 Aug 2009             Made GetAppPath return String explicitly.
	
	Public settingsPath As String ' the settings directory for the application
	Public runningLocal As Boolean ' whether we're running off the local folder
	'only
	Public commonSettingsPath As String ' the settings directory for the WebbIE applications
	Public nonRoamingSettingsPath As String ' the settings directory for non-roaming data
	'(e.g. Local Settings)
	
	'SHGetSpecialFolderLocation
	'Returns the Folder ID of the user's My Documents folder (or another folder indicated
	'by CSIDL)
	Private Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwnd As Integer, ByVal nFolder As Integer, ByRef ppidl As Integer) As Integer
	'SHGetPathFromIDList
	'Returns the path (string) from the folder ID obtained by SHGetSpecialFolderLocation
	Public Declare Function SHGetPathFromIDList Lib "Shell32"  Alias "SHGetPathFromIDListA"(ByVal Pidl As Integer, ByVal pszPath As String) As Integer
	
	' constants for Shell.NameSpace method -- these are the "special folders"
	'contained in the Windows shell
	Public Const CSIDL_DESKTOP As Integer = &H0 ' Desktop
	Public Const CSIDL_INTERNET As Integer = &H1 ' The internet
	Public Const CSIDL_PROGRAMS As Integer = &H2 ' Shortcuts in the Programs menu
	Public Const CSIDL_CONTROLS As Integer = &H3 ' Control Panel
	Public Const CSIDL_PRINTERS As Integer = &H4 ' Printers
	Public Const CSIDL_PERSONAL As Integer = &H5 ' Shortcuts to Personal files
	Public Const CSIDL_FAVORITES As Integer = &H6 ' Shortcuts to favorite folders
	Public Const CSIDL_STARTUP As Integer = &H7 ' Shortcuts to apps that start at boot Time
	Public Const CSIDL_RECENT As Integer = &H8 ' Shortcuts to recently used docs
	Public Const CSIDL_SENDTO As Integer = &H9 ' Shortcuts for the SendTo menu
	Public Const CSIDL_BITBUCKET As Integer = &HA ' Recycle Bin
	Public Const CSIDL_STARTMENU As Integer = &HB ' User-defined items in Start Menu
	Public Const CSIDL_DESKTOPDIRECTORY As Integer = &H10 ' Directory with all the desktop shortcuts
	Public Const CSIDL_DRIVES As Integer = &H11 ' My Computer
	Public Const CSIDL_NETWORK As Integer = &H12 ' Network Neighborhood virtual folder
	Public Const CSIDL_NETHOOD As Integer = &H13 ' Directory containing objects in the network neighborhood
	Public Const CSIDL_FONTS As Integer = &H14 ' Installed fonts
	Public Const CSIDL_TEMPLATES As Integer = &H15 ' Shortcuts to document templates
	Public Const CSIDL_COMMON_STARTMENU As Integer = &H16 ' Directory with items in the Start menu for all users
	Public Const CSIDL_COMMON_PROGRAMS As Integer = &H17 ' Directory with items in the Programs menu for all users
	Public Const CSIDL_COMMON_STARTUP As Integer = &H18 ' Directory with items in the StartUp submenu for all users
	Public Const CSIDL_COMMON_DESKTOPDIRECTORY As Integer = &H19 ' Directory with items on the desktop of all users
	Public Const CSIDL_APPDATA As Integer = &H1A ' Folder for application-specific data
	Public Const CSIDL_PRINTHOOD As Integer = &H1B ' Directory with references to printer links
	Public Const CSIDL_LOCAL_APPDATA As Integer = &H1C '{user name}\Local Settings\Application Data (non roaming)
	Public Const CSIDL_ALTSTARTUP As Integer = &H1D ' (DBCS) Directory corresponding to user 's nonlocalized Startup program group
	Public Const CSIDL_COMMON_ALTSTARTUP As Integer = &H1E ' (DBCS) Directory with Startup items for all users
	Public Const CSIDL_COMMON_FAVORITES As Integer = &H1F ' Directory with all user's favorit items
	Public Const CSIDL_INTERNET_CACHE As Integer = &H20 ' Directory for temporary internet Files
	Public Const CSIDL_COOKIES As Integer = &H21 ' Directory for Internet cookies
	Public Const CSIDL_HISTORY As Integer = &H22 ' Directory for Internet history items
	Public Const CSIDL_COMMON_APPDATA As Integer = &H23 'All Users\Application Data
	Public Const CSIDL_WINDOWS As Integer = &H24 'GetWindowsDirectory()
	Public Const CSIDL_SYSTEM As Integer = &H25 'GetSystemDirectory()
	Public Const CSIDL_PROGRAM_FILES As Integer = &H26 'C:\Program Files
	Public Const CSIDL_MYPICTURES As Integer = &H27 'C:\Program Files\My Pictures
	Public Const CSIDL_PROFILE As Integer = &H28 'USERPROFILE
	Public Const CSIDL_SYSTEMX86 As Integer = &H29 'x86 system directory on RISC
	Public Const CSIDL_PROGRAM_FILESX86 As Integer = &H2A 'x86 C:\Program Files on RISC
	Public Const CSIDL_PROGRAM_FILES_COMMON As Integer = &H2B 'C:\Program Files\Common
	Public Const CSIDL_PROGRAM_FILES_COMMONX86 As Integer = &H2C 'x86 Program Files\Common on RISC
	Public Const CSIDL_COMMON_TEMPLATES As Integer = &H2D 'All Users\Templates
	Public Const CSIDL_COMMON_DOCUMENTS As Integer = &H2E 'All Users\Documents
	Public Const CSIDL_COMMON_ADMINTOOLS As Integer = &H2F 'All Users\Start Menu\Programs\Administrative Tools
	Public Const CSIDL_ADMINTOOLS As Integer = &H30 '{user name}\Start Menu\Programs\Administrative Tools
	
	Public Const CSIDL_FLAG_CREATE As Integer = &H8000 'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
	Public Const CSIDL_FLAG_DONT_VERIFY As Integer = &H4000 'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
	Public Const CSIDL_FLAG_MASK As Integer = &HFF00 'mask for all possible flag values
	
	Public Sub DetermineSettingsPath(Optional ByVal companyName As String = "", Optional ByVal productName As String = "", Optional ByVal version As String = "")
		'works out whether we are running on a memory stick/standalone or as an
		'installed application.
		On Error Resume Next
		Dim fso As Scripting.FileSystemObject
		Dim key As String
		Dim section As String
		Dim Path As String
        Dim got As String

		'Get any override values for company, version and application name from the program .ini file, if any.
		'Otherwise use the values passed.
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        companyName = modIniFile.GetString("ApplicationDataPath", "Company", companyName, GetAppPath() & "\" & My.Application.Info.AssemblyName & ".ini")
		If Len(companyName) = 0 Then
            companyName = My.Application.Info.CompanyName
        End If
        'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        productName = modIniFile.GetString("ApplicationDataPath", "ApplicationName", productName, GetAppPath() & "\" & My.Application.Info.AssemblyName & ".ini")
        If Len(productName) = 0 Then
            productName = My.Application.Info.ProductName
        End If
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        version = modIniFile.GetString("ApplicationDataPath", "Version", version, GetAppPath() & "\" & My.Application.Info.AssemblyName & ".ini")
        If Len(version) = 0 Then
            version = CStr(My.Application.Info.Version.Major)
        End If
		
		fso = New Scripting.FileSystemObject
		If fso.FileExists(GetAppPath & "\installed.ini") Then
			runningLocal = False
		Else
			'try checking for local INI file to indicate not running from stick
			key = "RunAsInstalled" & Chr(0)
			section = "Program" & Chr(0)
			'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			Path = GetAppPath & "\" & My.Application.Info.AssemblyName & ".ini" & Chr(0)
            got = modIniFile.GetString(section, key, "0" & Chr(0), Path)
            If CBool(got) Then
                'ini file indicates we should run as installed version
                runningLocal = False
            Else
                'running from a memory stick or other non-installed location
                runningLocal = True
            End If
		End If
		If runningLocal Then
			'running from a memory stick or other non-installed location
			settingsPath = GetAppPath & "\Settings"
			nonRoamingSettingsPath = settingsPath
			'need to create
			If Not fso.FolderExists(settingsPath) Then Call fso.CreateFolder(settingsPath)
		Else
			'run as installed version
			settingsPath = GetSpecialFolderPath(CSIDL_APPDATA)
			'3.12.2 Can't access Local Appdata - it's per-machine.
			nonRoamingSettingsPath = GetSpecialFolderPath(CSIDL_APPDATA) ' GetSpecialFolderPath(CSIDL_LOCAL_APPDATA)
		End If
		settingsPath = settingsPath & "\" & companyName
		If Not fso.FolderExists(settingsPath) Then Call fso.CreateFolder(settingsPath)
		nonRoamingSettingsPath = nonRoamingSettingsPath & "\" & companyName
		If Not fso.FolderExists(nonRoamingSettingsPath) Then Call fso.CreateFolder(nonRoamingSettingsPath)
		commonSettingsPath = settingsPath & "\Common"
		If Not fso.FolderExists(commonSettingsPath) Then Call fso.CreateFolder(commonSettingsPath)
		settingsPath = settingsPath & "\" & productName
		If Not fso.FolderExists(settingsPath) Then Call fso.CreateFolder(settingsPath)
		nonRoamingSettingsPath = nonRoamingSettingsPath & "\" & productName
		If Not fso.FolderExists(nonRoamingSettingsPath) Then Call fso.CreateFolder(nonRoamingSettingsPath)
		settingsPath = settingsPath & "\" & version
		If Not fso.FolderExists(settingsPath) Then Call fso.CreateFolder(settingsPath)
		nonRoamingSettingsPath = nonRoamingSettingsPath & "\" & version
		If Not fso.FolderExists(nonRoamingSettingsPath) Then Call fso.CreateFolder(nonRoamingSettingsPath)
        fso = Nothing
	End Sub
	
	Public Function GetSpecialFolderPath(ByRef CSIDL As Integer) As String
		'returns the special folder indicated by the CSIDL constant
		On Error Resume Next
		Dim Path As String
		Dim result As Integer
		Dim referenceID As Integer
		
		'work out where to save them to:
		Path = Space(260)
		result = SHGetSpecialFolderLocation(0, CSIDL, referenceID)
		result = SHGetPathFromIDList(referenceID, Path)
		'assertion: path now contains path to special folder
		Path = Trim(Path)
		'take off final null character which trim has left behind
		Path = Replace(Path, Chr(0), "")
		'return
		GetSpecialFolderPath = Path
	End Function
	
	Public Function GetAppPath() As String
		On Error Resume Next
		'work out some paths for use
		GetAppPath = My.Application.Info.DirectoryPath
		If Right(GetAppPath, 1) = "\" Then
			GetAppPath = Left(GetAppPath, Len(GetAppPath) - 1)
		End If
	End Function
	
	'UPGRADE_NOTE: default was upgraded to default_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetSettingIni(ByRef appName As String, ByRef section As String, ByRef key As String, ByRef default_Renamed As String) As String
		On Error Resume Next
		Dim value As String
		Dim size As Integer
		Dim nullTerminatedDefault As String
		Dim Path As String
		
		nullTerminatedDefault = default_Renamed & Chr(0)
		section = section & Chr(0)
		key = key & Chr(0)
		value = Space(256) & Chr(0)
		Path = settingsPath & "\" & appName & ".ini" & Chr(0)
        GetSettingIni = modIniFile.GetString(section, key, nullTerminatedDefault, Path)
    End Function
	
	Public Sub SaveSettingIni(ByRef appName As String, ByRef section As String, ByRef key As String, ByRef value As String)
		On Error Resume Next
		Dim Path As String
		
		section = section & Chr(0)
		key = key & Chr(0)
		value = value & Chr(0)
		Path = settingsPath & "\" & appName & ".ini"
		
        Call modIniFile.WriteString(section, key, value, Path)
	End Sub
	
	'UPGRADE_NOTE: default was upgraded to default_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function ReadAppEXEIni(ByRef section As String, ByRef key As String, ByRef default_Renamed As String) As String
		On Error Resume Next
		Dim iniFile As String
		Dim got As String
		
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		iniFile = GetAppPath & "\" & My.Application.Info.AssemblyName & ".ini" & Chr(0)
		section = section & Chr(0)
		key = key & Chr(0)
		default_Renamed = default_Renamed & Chr(0)
		got = Space(255) & Chr(0)
		
        Return modIniFile.GetString(section, key, default_Renamed, iniFile)
    End Function
End Module