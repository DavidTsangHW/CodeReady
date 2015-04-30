Option Strict Off
Option Explicit On
Module Bas_OpenURLByDefaultBrowser
	Public Declare Function ShellExecute Lib "shell32.dll"  Alias "ShellExecuteA"(ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
	
	Public Declare Function FindExecutable Lib "shell32.dll"  Alias "FindExecutableA"(ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Integer
	
	Public Sub OpenBrowser(ByRef URL As String)
		Dim hwnd As Object
		
		'This function open an URL by the default browser of the machine.
		'Aware, there will be an void browser appeared while passing a mailto: url.
		
		Dim FileName As String
		Dim Dummy As String
		Dim BrowserExec As New VB6.FixedLengthString(255)
		Dim RetVal As Integer
		Dim FileNumber As Short
		
		Const SW_SHOW As Short = 5 ' Displays Window in its current size
		' and position
		Const SW_SHOWNORMAL As Short = 1 ' Restores Window if Minimized or
		' Maximized
		
		' First, create a known, temporary HTML file
		BrowserExec.Value = Space(255)
		FileName = "C:\temphtm.HTM"
		FileNumber = FreeFile ' Get unused file number
		FileOpen(FileNumber, FileName, OpenMode.Output) ' Create temp HTML file
		WriteLine(FileNumber, "<HTML> <\HTML>") ' Output text
		FileClose(FileNumber) ' Close file
		' Then find the application associated with it
		RetVal = FindExecutable(FileName, Dummy, BrowserExec.Value)
		BrowserExec.Value = Trim(BrowserExec.Value)
		' If an application is found, launch it!
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If RetVal <= 32 Or IsNothing(BrowserExec.Value) Then ' Error
			MsgBox("Could not find associated Browser", MsgBoxStyle.Exclamation, "Browser Not Found")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object hwnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RetVal = ShellExecute(hwnd, "open", BrowserExec.Value, URL, Dummy, SW_SHOWNORMAL)
			If RetVal <= 32 Then ' Error
				MsgBox("Web Page not Opened", MsgBoxStyle.Exclamation, "URL Failed")
			End If
		End If
		Kill(FileName)
		'delete temp HTML file
	End Sub
End Module