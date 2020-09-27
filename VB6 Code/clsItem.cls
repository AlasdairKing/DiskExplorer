VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPath As String
Private mType As itemType

Public Enum itemType
    FILE_TYPE = 0
    FOLDER_TYPE = 1
    UP_FOLDER = 2
    DESKTOP = 3
    MY_DOCUMENTS = 4
    MAIN_DRIVE
    CD_DRIVE
    USB_DRIVE
    CURRENT_FOLDER
End Enum

Public Property Get path() As String
    On Error Resume Next
    path = mPath
End Property

Public Property Let path(newPath As String)
    On Error Resume Next
'''    On Error GoTo PathInvalid
'''    If InStr(1, newPath, "\") > 0 Then
'''        If Dir(newPath, vbDirectory) <> "" Or Dir(newPath, vbNormal) <> "" Then
'''            mPath = newPath
'''        Else
'''            'Call Err.Raise(1000, , "Invalid path/file")
'''        End If
'''    Else
'''        'It's a drive folder, e.g. "C:"
        mPath = newPath
'''    End If
'''    Exit Property
'''PathInvalid:
'''    'Don't change path
'''    If mPath = "" Then mPath = "C:\"
End Property

Public Property Get itemType() As itemType
    On Error Resume Next
    itemType = mType
End Property

Public Property Let itemType(newType As itemType)
    On Error Resume Next
    mType = newType
End Property
