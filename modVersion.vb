Option Strict Off
Option Explicit On
Module modVersion
	'Calendar
	'Copyright Alasdair King, 2010, http://www.alasdairking.me.uk
	'Released under the GNU Public Licence, Version 3.
	
	'modVersion
	
	'Finds out the installation version of the suite by reading version.ini

	Public Function GetPackageVersion() As String
		On Error Resume Next
		Dim s As String
		s = Space(255)
        Return modIniFile.GetString("Package", "Version", "", My.Application.Info.DirectoryPath & "\version.ini")
    End Function
End Module