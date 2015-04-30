Option Strict Off
Option Explicit On
Module Bas_License
	Public Const ModuleName As String = "Bas_License"
	
	Public RegisterName As String
	Public LicenseKey As String
	
	Public IsLicensed As Boolean
	
	Public Sub SaveLicense(ByVal RegisterNameString As String, ByVal UserCodeKey As String)
		
		
		Dim Ts As Scripting.TextStream
		
		Ts = Fso.OpenTextFile(AppPath("license.dat"), Scripting.IOMode.ForWriting, True)
		
		Ts.WriteLine("Version: " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision)
		
		Ts.WriteLine("Date: " & VB6.Format(Now, "YYYY/MM/DD"))
		
		Ts.WriteLine("Register Name: " & RegisterNameString)
		
		Ts.WriteLine("License Key: " & UserCodeKey)
		
		Ts.Close()
		
	End Sub
	
	Public Function ValidateLicense(ByVal RegisterNameString As String, ByVal UserCodeKey As String) As Boolean
		Dim Key As Object
		
		Dim Code(16) As Short
		Dim RegisterName(16) As Short
		
		Dim Pointer(7) As Boolean
		Dim Pointers As String
		
		Dim idx As Short
		
		Dim CodeString As String
		Dim KeyString As String
		Dim KeyLength As String
		
		Dim UserKey As String
		
		If Len(UserCodeKey) < 7 Then
			
			Exit Function
			
		End If
		
		CodeString = Mid(UserCodeKey, 1, 7)
		
		For idx = 0 To 6
			
			Code(idx) = Asc(Mid(CodeString, idx + 1, 1))
			
		Next 
		
		For idx = 0 To 8
			
			RegisterName(idx) = 32
			
			If Len(RegisterNameString) > idx Then
				
				Select Case Asc(Mid(UCase(RegisterNameString), idx + 1, 1))
					
					Case 48 To 57
						
						RegisterName(idx) = Asc(Mid(UCase(RegisterNameString), idx + 1, 1))
						
					Case 65 To 90
						
						RegisterName(idx) = Asc(Mid(UCase(RegisterNameString), idx + 1, 1))
						
					Case Else
						
						RegisterName(idx) = 32
						
				End Select
				
			End If
			
		Next 
		
		Pointers = DecimalToBinary(CInt(RegisterName(3)), 7)
		
		For idx = 0 To 6
			
			Select Case CShort(Mid(Pointers, idx + 1, 1))
				
				Case 0
					
					Pointer(idx) = True
					
				Case 1
					
					Pointer(idx) = False
					
			End Select
			
		Next 
		
		
		For idx = 0 To 6
			
			If Pointer(idx) = True Then
				
				Select Case Code(idx) + idx
					
					Case 0 To 47
						
						'UPGRADE_WARNING: Couldn't resolve default property of object Key. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Key = Key & Chr(49)
						
					Case 58 To 64
						
						'UPGRADE_WARNING: Couldn't resolve default property of object Key. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Key = Key & Chr(Code(idx) + idx + 7)
						
					Case Is > 90
						
						'UPGRADE_WARNING: Couldn't resolve default property of object Key. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Key = Key & Chr(Code(idx) + idx - 90 + 65)
						
					Case Else
						
						'UPGRADE_WARNING: Couldn't resolve default property of object Key. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Key = Key & Chr(Code(idx) + idx)
						
				End Select
				
			End If
			
		Next 
		
		KeyLength = CStr(Asc(Right(UserCodeKey, 1)))
		
		KeyLength = CStr(CShort(Right(CStr(KeyLength), 1)))
		
		UserKey = Right(UserCodeKey, CDbl(KeyLength) + 1)
		
		UserKey = Left(UserKey, CInt(KeyLength))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Key. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Key = UserKey Then
			
			ValidateLicense = True
			
		End If
		
	End Function
	
	Public Sub ReadLicenseFile(ByRef RegisterName As String, ByRef LicenseKey As String)
		
		Dim Ts As Scripting.TextStream
		Dim ReadLine As String
		
		If Fso.FileExists(AppPath("license.dat")) = False Then
			
			Exit Sub
			
		End If
		
		Ts = Fso.OpenTextFile(AppPath("license.dat"), Scripting.IOMode.ForReading, False)
		
		Do Until Ts.AtEndOfStream
			
			ReadLine = Ts.ReadLine
			
			Select Case True
				
				
				Case InStr(1, " " & ReadLine, "Register Name:") > 0
					
					RegisterName = Mid(ReadLine, Len("Register Name: ") + 1)
					
				Case InStr(1, " " & ReadLine, "License Key: ") > 0
					
					
					LicenseKey = Mid(ReadLine, Len("License Key: ") + 1)
					
			End Select
			
			
		Loop 
		
		
		
		
		
	End Sub
End Module