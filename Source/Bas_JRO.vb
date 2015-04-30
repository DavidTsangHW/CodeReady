Option Strict Off
Option Explicit On
Module Bas_JRO
	Public Function CompactDatabase(ByRef Filepath As String) As Boolean
		Dim Source As Object
		
		Dim Connstr As String
		Dim Jr As New JRO.JetEngine
		
		Dim Destination As String
		
		On Error Resume Next
		
		CompactDatabase = True
		Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Filepath & ";Persist Security Info=False"
		
		Destination = AppPath(FSO.GetTempName)
		
		Jr.CompactDatabase(Connstr, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Destination & " ;Jet OLEDB:Engine Type=4")
		
		If Not Err.Number = 0 Then
			MsgBox(Err.Description, MsgBoxStyle.Critical)
			CompactDatabase = False
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Source. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileCopy(Destination, Source)
		End If
		
		Kill(Destination)
		
	End Function
End Module