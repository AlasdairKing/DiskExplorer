Attribute VB_Name = "modShellExec"
Option Explicit

'see http://www.nsftools.com/tips/ShellExec.lss
'** ShellExecute will open a file using the registered file association on the computer.
'** If it returns a value of greater than 32 then the call was successful; otherwise
'** it should return one of the error codes below. The parameters are:
'**     hwnd = an active window handle, or 0
'**     operation = "edit", "explore", "find", "open", or "print"
'**     fileName = a file or directory name
'**     parameters = if fileName is an executable file, the command line parameters
'**                         to pass when launching the application, or "" if no parameters
'**                         are necessary
'**     directory = the default directory to use, or "" if you don't care
'**     displayType = one of the displayType constants listed below
Public Declare Function ShellExecute Lib "Shell32" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal operation As String, ByVal filename As String, _
ByVal parameters As String, ByVal directory As String, ByVal displayType As Long) As Long

'** FindExecutable will determine the executable file that is set up to open a particular
'** file based on the file associations on this computer. If it returns a value of greater than
'** 32 then the call was successful; otherwise it should return one of the error codes
'** below. The parameters are:
'**     fileName = the full path to the file you are trying to find the association for
'**     directory = the default directory to use, or "" if you don't care
'**     retAssociation = the associated executable will be returned as this parameter,
'**                         with a maximum string length of 255 characters (you will want
'**                         to pass a String that's 256 characters long and trim the
'**                         null-terminated result)
Public Declare Function FindExecutable Lib "Shell32" Alias "FindExecutableA" _
(ByVal filename As String, ByVal directory As String, ByVal retAssociation As String) As Long

'** constants for the displayType parameter
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10
'** possible errors returned by ShellExecute
Public Const ERROR_OUT_OF_MEMORY = 0       'The operating system is out of memory or resources.
Public Const ERROR_FILE_NOT_FOUND = 2      'The specified file was not found.
Public Const ERROR_PATH_NOT_FOUND = 3  'The specified path was not found.
Public Const ERROR_BAD_FORMAT = 11         'The .exe file is invalid (non-Microsoft Win32® .exe or error in .exe image).
Public Const SE_ERR_FNF = 2                            'The specified file was not found.
Public Const SE_ERR_PNF = 3                        'The specified path was not found.
Public Const SE_ERR_ACCESSDENIED = 5       'The operating system denied access to the specified file.
Public Const SE_ERR_OOM = 8                        'There was not enough memory to complete the operation.
Public Const SE_ERR_SHARE = 26                 'A sharing violation occurred.
Public Const SE_ERR_ASSOCINCOMPLETE = 27   'The file name association is incomplete or invalid.
Public Const SE_ERR_DDETIMEOUT = 28            'The DDE transaction could not be completed because the request timed out.
Public Const SE_ERR_DDEFAIL = 29               'The DDE transaction failed.
Public Const SE_ERR_DDEBUSY = 30               'The Dynamic Data Exchange (DDE) transaction could not be completed because other DDE transactions were being processed.
Public Const SE_ERR_NOASSOC = 31               'There is no application associated with the given file name extension. This error will also be returned if you attempt to print a file that is not printable.
Public Const SE_ERR_DLLNOTFOUND = 32       'The specified dynamic-link library (DLL) was not found.

