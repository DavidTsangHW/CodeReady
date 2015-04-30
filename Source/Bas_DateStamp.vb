Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module Bas_DateStamp
	Private Const ModuleName As String = "Bas_InstallationDate"
	Public Const TrialPeriod As Short = 7
	Public InstallationDate As Date
	Public TrialRemaining As Short
	Public Const AccessLogFile As String = "cid.dat"
	
	Public Const InstallationID_Code As Short = 8
	Public Const AccessID_Code As Short = 3
	
	'UPGRADE_NOTE: DateString was upgraded to DateString_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetMaskedDate(ByVal Code As Short, ByRef DateString_Renamed As String) As String
		
		Dim BaseId(32) As Short
		Dim BaseIdString As String
		
		Dim MaskedDate(6) As Short
		Dim MaskedDateString As String
		
		Dim Idx As Short
		
		BaseIdString = GetRandomId(16)
		
		For Idx = 1 To Len(BaseIdString)
			
			BaseId(Idx) = Asc(Mid(BaseIdString, Idx, 1))
			
		Next 
		
		For Idx = 1 To Len(DateString_Renamed)
			
			MaskedDate(Idx) = CShort(Mid(DateString_Renamed, Idx, 1)) + (BaseId(Code + Idx - 1) Mod (Code * 2)) + 65
			MaskedDateString = MaskedDateString & Chr(MaskedDate(Idx))
			
		Next 
		
		GetMaskedDate = BaseIdString & MaskedDateString
		
	End Function
	
	Public Function GetRandomId(ByVal length As Short) As String
		Dim MousePointer As Object
		
		Dim Key As String
		Dim Idx As Short
		Dim ThisKey As Short
		Dim LastKey As Short
		Dim IsAlpha As Short
		Dim Pass As Boolean
		
		Dim Ascii(255) As Short
		
		Call Randomize()
		Pass = False
		Key = ""
		LastKey = 256
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MousePointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MousePointer = System.Windows.Forms.Cursors.WaitCursor
		
		Do Until Pass = True
			
			For Idx = 0 To length
				
				IsAlpha = Rnd()
				
				Do While (True = True)
					
					If IsAlpha > 0.45 Then
						ThisKey = 65 + Int(65 * Rnd())
					Else
						ThisKey = 49 + Int(49 * Rnd())
					End If
					
					If ThisKey <> LastKey Then
						
						Select Case ThisKey
							
							Case 49 To 57
								Ascii(ThisKey) = Ascii(ThisKey) + 1
								Exit Do
								
							Case 65 To 78
								Ascii(ThisKey) = Ascii(ThisKey) + 1
								Exit Do
								
							Case 80 To 90
								Ascii(ThisKey) = Ascii(ThisKey) + 1
								Exit Do
								
						End Select
						
					End If
					
				Loop 
				
				LastKey = ThisKey
				Key = Key & Chr(ThisKey)
				
			Next 
			
			Pass = True
			
			For Idx = LBound(Ascii) To UBound(Ascii)
				
				If Ascii(Idx) >= 3 Then
					Pass = False
					Key = ""
				End If
				Ascii(Idx) = 0
				
			Next 
			
		Loop 
		
		GetRandomId = Key
		
	End Function
	
	Public Function GetSerialId() As String
		Dim GetMaskedString As Object
		
		Dim MaskedString As String
		
		MaskedString = GetSetting(My.Application.Info.Title, "License", "Serial Id")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object GetMaskedString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetMaskedString = MaskedString
		
	End Function
	
	Public Function GetInstallationId_Registry() As String
		
		Dim MaskedString As String
		
		MaskedString = GetSetting(My.Application.Info.Title, "License", "Installation Id")
		
		GetInstallationId_Registry = MaskedString
		
	End Function
	
	Public Function GetInstallationId_File() As String
		
		Dim Ts As Scripting.TextStream
		
		Dim ReadLine As String
		
		On Error GoTo ErrorHandler
		
		Ts = Fso.OpenTextFile(Environ("WINDIR") & "\" & AccessLogFile)
		
		ReadLine = Ts.ReadLine
		
		GetInstallationId_File = Mid(ReadLine, 1, 24)
		
		Ts.Close()
		
		'UPGRADE_NOTE: Object Ts may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Ts = Nothing
		
		Exit Function
		
ErrorHandler: 
		
		Exit Function
		
	End Function
	
	Public Function GetAccessId_File() As String
		
		Dim Ts As Scripting.TextStream
		
		Dim ReadLine As String
		
		On Error GoTo ErrorHandler
		
		Ts = Fso.OpenTextFile(Environ("WINDIR") & "\" & AccessLogFile)
		
		ReadLine = Ts.ReadLine
		
		GetAccessId_File = Mid(ReadLine, 25)
		
		Ts.Close()
		
		'UPGRADE_NOTE: Object Ts may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Ts = Nothing
		
		Exit Function
		
ErrorHandler: 
		
		Exit Function
		
	End Function
	
	
	Public Function GetAccessId_Registry() As String
		
		Dim MaskedString As String
		
		MaskedString = GetSetting(My.Application.Info.Title, "License", "Access Id")
		
		GetAccessId_Registry = MaskedString
		
	End Function
	
	Public Function IsExpired() As Boolean
		Dim vbSbhortDate As Object
		
		Dim InstallationID As String
		Dim AccessID As String
		
		Dim InstallationID_File As String
		Dim AccessID_File As String
		
		Dim AccessDate As Date
		Dim InstallationDate As Date
		
		Dim AccessDate_File As Date
		Dim InstallationDate_File As Date
		
		Dim CurrentDate As Date
        Dim ValidDate As Date

        IsExpired = False

        Exit Function

		
		InstallationID = GetInstallationId_Registry
		AccessID = GetAccessId_Registry
		
		InstallationID_File = GetInstallationId_File
		AccessID_File = GetAccessId_File
		
		On Error GoTo ErrorHandler
		
		IsExpired = True
		
		If Len(InstallationID) = 0 And Len(InstallationID_File) = 0 Then
			
			InstallationDate = CDate(FormatDateTime(Now, DateFormat.ShortDate))
			
			InstallationID = GenInstallationID(InstallationDate)
			
			AccessID = GenAccessID
			
			Call SaveInstallationID_Registry(InstallationID)
			
			Call SaveAccessID_Registry(AccessID)
			
			Call SaveAccessLogFile(GenInstallationID(InstallationDate), GenAccessID)
			
			TrialRemaining = TrialPeriod
			
			IsExpired = False
			
			Exit Function
			
		End If
		
		InstallationDate = UnmaskId(InstallationID, InstallationID_Code)
		'UPGRADE_WARNING: Couldn't resolve default property of object vbSbhortDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallationDate = CDate(FormatDateTime(InstallationDate, vbSbhortDate))
		
		AccessDate = UnmaskId(AccessID, AccessID_Code)
		AccessDate = CDate(FormatDateTime(AccessDate, DateFormat.ShortDate))
		
		InstallationDate_File = UnmaskId(InstallationID_File, InstallationID_Code)
		'UPGRADE_WARNING: Couldn't resolve default property of object vbSbhortDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallationDate_File = CDate(FormatDateTime(InstallationDate_File, vbSbhortDate))
		
		AccessDate_File = UnmaskId(AccessID_File, AccessID_Code)
		AccessDate_File = CDate(FormatDateTime(AccessDate_File, DateFormat.ShortDate))
		
		CurrentDate = CDate(FormatDateTime(Now, DateFormat.ShortDate))
		
		ValidDate = CurrentDate
		
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		If DateDiff(Microsoft.VisualBasic.DateInterval.Day, AccessDate, CurrentDate) < 0 Then
			
			ValidDate = AccessDate
			
		Else
			
			AccessID = GenAccessID
			
			InstallationID = GenInstallationID(InstallationDate)
			
			Call SaveInstallationID_Registry(InstallationID)
			
			Call SaveAccessID_Registry(AccessID)
			
			Call SaveAccessLogFile(GenInstallationID(InstallationDate), GenAccessID)
			
		End If
		
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		TrialRemaining = TrialPeriod - DateDiff(Microsoft.VisualBasic.DateInterval.Day, InstallationDate, ValidDate)
		
		If TrialRemaining < 0 Or TrialRemaining > TrialPeriod Then
			
			IsExpired = True
			
		Else
			
			IsExpired = False
			
		End If
		
		Exit Function
		
ErrorHandler: 
		
		IsExpired = True
		
	End Function
	
	Public Sub SaveInstallationID_Registry(ByVal InstallationID As String)
		
		On Error GoTo ErrorHandler
		
		Call SaveSetting(My.Application.Info.Title, "License", "Installation ID", InstallationID)
		
		Exit Sub
		
ErrorHandler: 
		
		MsgBox(Err.Number & ": " & Err.Number)
		
	End Sub
	
	Public Function GenInstallationID(ByVal InstallationDate As Date) As String
		
		'UPGRADE_NOTE: DateString was upgraded to DateString_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim DateString_Renamed As String
		
		DateString_Renamed = leading(CStr(VB.Day(InstallationDate)), 2, "0") & leading(CStr(Month(InstallationDate)), 2, "0") & Mid(CStr(Year(InstallationDate)), 3, 2)
		
		GenInstallationID = GetMaskedDate(InstallationID_Code, DateString_Renamed)
		
	End Function
	
	Public Sub SaveAccessID_Registry(ByVal AccessID As String)
		
		On Error GoTo ErrorHandler
		
		Call SaveSetting(My.Application.Info.Title, "License", "Access ID", AccessID)
		
		Exit Sub
		
ErrorHandler: 
		
		MsgBox(Err.Number & ": " & Err.Description)
		
	End Sub
	
	Public Sub SaveAccessLogFile(ByVal InstallationID As String, ByVal AccessID As String)
		
		Dim MaskedString As String
		
		On Error GoTo ErrorHandler
		
		MaskedString = InstallationID & AccessID
		
		Call WriteNewFile(Environ("WINDIR") & "\" & AccessLogFile, MaskedString)
		
		Exit Sub
		
ErrorHandler: 
		
		MsgBox(Err.Number & ": " & Err.Description)
		
	End Sub
	
	
	Public Function GenAccessID() As String
		
		'UPGRADE_NOTE: DateString was upgraded to DateString_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim DateString_Renamed As String
		Dim AccessDate As String
		
		AccessDate = CStr(Now)
		
		DateString_Renamed = leading(CStr(VB.Day(CDate(AccessDate))), 2, "0") & leading(CStr(Month(CDate(AccessDate))), 2, "0") & Mid(CStr(Year(CDate(AccessDate))), 3, 2)
		
		GenAccessID = GetMaskedDate(AccessID_Code, DateString_Renamed)
		
	End Function
	
	Public Function GetMinInstallationDate(ByVal InstallationID As String, ByVal InstallationID2 As String) As Date
		
		Dim InstallationDate1 As Date
		Dim InstallationDate2 As Date
		Dim MinInstallationDate As Date
		
		
		
		
		
		
		
		
		
	End Function
	
	Public Function GetMaxAccessDate(ByVal AccessID As String, ByVal AccessID2 As String) As Date
		
	End Function
	
	Public Function UnmaskId(ByVal MaskedString As String, ByVal Code As Short) As Date
		Dim Idx As Object
		
		Dim BaseId(32) As Short
		Dim BaseIdString As String
		
		Dim MaskedDate(6) As Short
		Dim MaskedDateString As String
		
		'UPGRADE_NOTE: DateString was upgraded to DateString_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim DateString_Renamed As String
		
		BaseIdString = Mid(MaskedString, Code, 6)
		
		MaskedDateString = Mid(MaskedString, Len(MaskedString) - 6 + 1, 6)
		
		For Idx = 1 To Len(BaseIdString)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Idx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BaseId(Idx) = Asc(Mid(BaseIdString, Idx, 1))
			
		Next 
		
		For Idx = 1 To Len(MaskedDateString)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Idx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MaskedDate(Idx) = Asc(Mid(MaskedDateString, Idx, 1))
			
		Next 
		
		For Idx = 1 To Len(MaskedDateString)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Idx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DateString_Renamed = DateString_Renamed & MaskedDate(Idx) - BaseId(Idx) Mod (Code * 2) - 65
			
		Next 
		
		UnmaskId = ToSystemDateFormat(leading(Mid(DateString_Renamed, 1, 2), 2, "0"), leading(Mid(DateString_Renamed, 3, 2), 2, "0"), Mid(DateString_Renamed, 5, 2))
		
	End Function
	
	Public Function GetMaskedSystemSerialId() As String
		
		Dim MaskedDiskSerialId(32) As Short
		
		Dim DiskSerialId As String
		Dim MaskedDiskSerialIdString As String
		
		Dim Idx As Short
		
		DiskSerialId = CStr(Fso.Drives(Left(My.Application.Info.DirectoryPath, 1)).SerialNumber)
		
		For Idx = 1 To Len(DiskSerialId)
			
			MaskedDiskSerialId(Idx) = Asc(Mid(DiskSerialId, Idx, 1)) Mod 10 + 65
			MaskedDiskSerialIdString = MaskedDiskSerialIdString & Chr(MaskedDiskSerialId(Idx))
			
		Next 
		
		GetMaskedSystemSerialId = MaskedDiskSerialIdString
		
	End Function
	
	Public Function ToSystemDateFormat(ByVal D As String, ByVal M As String, ByVal Y As String) As Date
		
		Dim ShortDateString As String
		Dim CurrentDate As String
		
		'UPGRADE_NOTE: DateString was upgraded to DateString_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim DateString_Renamed As String
		
		DateString_Renamed = VB6.Format(Now, "DD/MM/YYYY")
		
		ShortDateString = FormatDateTime(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -VB.Day(Now) + 1, Now), DateFormat.ShortDate)
		
		DateString_Renamed = "1" & Mid(DateString_Renamed, 3)
		
		If ShortDateString = DateString_Renamed Then
			
			ToSystemDateFormat = CDate(D & "/" & M & "/" & Y)
			
		Else
			
			ToSystemDateFormat = CDate(M & "/" & D & "/" & Y)
			
		End If
		
	End Function
End Module