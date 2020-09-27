Option Strict On
Option Explicit On
Module modDeleteFile
	
	Public Const ERROR_OK As Integer = 0
	Private Structure SHFILEOPTSTRUCT
		Dim hWnd As Integer
		Dim wFunc As Integer
		Dim pFrom As String
		Dim pTo As String
		Dim fFlags As Short
		Dim fAnyOperationsAborted As Integer
		Dim hNameMappings As Integer
		Dim lpszProgressTitle As Integer
	End Structure
	'UPGRADE_WARNING: Structure SHFILEOPTSTRUCT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function SHFileOperation Lib "Shell32.dll"  Alias "SHFileOperationA"(ByRef lpFileOp As SHFILEOPTSTRUCT) As Integer
	Private Const FO_DELETE As Integer = &H3
	Private Const FOF_ALLOWUNDO As Integer = &H40
	
	Public Function DeleteFileToRecycleBin(ByRef Filename As String) As Integer
		On Error Resume Next
        My.Computer.FileSystem.DeleteFile(Filename, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)
    End Function
End Module