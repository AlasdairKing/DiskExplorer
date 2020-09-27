Attribute VB_Name = "modDeleteFile"
Option Explicit

Public Const ERROR_OK As Long = 0
Private Type SHFILEOPTSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type
Private Declare Function SHFileOperation Lib "Shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPTSTRUCT) As Long
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40

Public Function DeleteFileToRecycleBin(Filename As String) As Long
    On Error Resume Next
    Dim fop As SHFILEOPTSTRUCT
    
    With fop
        .wFunc = FO_DELETE
        .pFrom = Filename
        .fFlags = FOF_ALLOWUNDO
    End With
    DeleteFileToRecycleBin = SHFileOperation(fop)
End Function

