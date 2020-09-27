Option Strict Off
Option Explicit On
Friend Class clsItem
	
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
	
	
	Public Property path() As String
		Get
			On Error Resume Next
			path = mPath
		End Get
		Set(ByVal Value As String)
			On Error Resume Next
			mPath = Value
		End Set
	End Property
	
	
	Public Property itemType_Renamed() As itemType
		Get
			On Error Resume Next
            itemType_Renamed = mType
		End Get
		Set(ByVal Value As itemType)
			On Error Resume Next
			mType = Value
		End Set
	End Property
End Class