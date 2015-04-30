Option Strict Off
Option Explicit On
Module Bas_ADO
	Private Const ModuleName As String = "Bas_ADO"
	
	Public Function ADO_TestConnection(ByRef ConnectionString As String) As Boolean
		
		Dim Cn As New ADODB.Connection
		
		On Error GoTo ErrorHandler
		
		Cn.Open(ConnectionString)
		
		Cn.Close()
		
		'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Cn = Nothing
		
		ADO_TestConnection = True
		
		Exit Function
		
ErrorHandler: 
		ADO_TestConnection = False
		
	End Function
	
	Public Function ADOFieldSum(ByVal Rs As ADODB.Recordset, Optional ByRef Fieldidx As Short = 0) As Single
		
		Dim Sum As Single
		
		Dim CloneRs As New ADODB.Recordset
		
		If Rs.RecordCount <= 0 Then
			Exit Function
		End If
		
		CloneRs = Rs.Clone
		
		With CloneRs
			
			Sum = 0
			
			.MoveFirst()
			
			Do Until .EOF
				If .Fields(Fieldidx).Value > 0 Then
					Sum = Sum + .Fields(Fieldidx).Value
				End If
				.MoveNext()
			Loop 
			
		End With
		
		ADOFieldSum = Sum
		
	End Function
	
	Public Function GetRsAbsolutePosition(ByVal Rs As ADODB.Recordset, ByVal Position As Double) As ADODB.Recordset
		
		If Position < 0 Then
			Exit Function
		End If
		
		If Rs.RecordCount = 0 Then
			Exit Function
		End If
		
		Rs.MoveFirst()
		Rs.Move(Position)
		
		GetRsAbsolutePosition = Rs
		
	End Function
	
	Public Function Get_OLEDB_4_0_ConnectionString(ByVal FilePath As String) As String
		
		Dim connstr As String
		
		connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
		connstr = connstr & FilePath
		connstr = connstr & ";Persist Security Info=False"
		
		Get_OLEDB_4_0_ConnectionString = connstr
		
	End Function
	
	Public Function Get_SQLOLEDB_ConnectionString(ByVal ServerName As String) As String
		
		Dim connstr As String
		
		connstr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=ACTUARY;Data Source="
		connstr = connstr & ServerName
		
		Get_SQLOLEDB_ConnectionString = connstr
		
		
	End Function
	
	Public Function ADO_DefaultValue(ByVal FieldType As Short) As String
		
		Select Case FieldType
			
			' Listindex
			'0  Text
			'1  Integer
			'2  Real
			'3  Letter
			'4  Boolean
			'5  Not Supportted
			'6  Unknown
			
			Case 20 'An 8 -byte signed integer
				ADO_DefaultValue = CStr(0)
				
			Case 128 'A binary value
				ADO_DefaultValue = CStr(0)
				
			Case 11 ' A Boolean value
				ADO_DefaultValue = CStr(0)
				
			Case 8 'A null-terminated character string
				ADO_DefaultValue = " "
				
			Case 136 'A chapter type, indicating a child recordset
				ADO_DefaultValue = " "
				
			Case 129 'A String Value
				ADO_DefaultValue = " "
				
			Case 6 'A currency value, An 8-byte signed integer scaled by 10,000
				'with 4 digits to the right of the decimal point
				ADO_DefaultValue = CStr(0)
				
				
			Case 7 'A Date value. A Double where the whole partis the number of dayssince december 30 1899, and the fractional part is a fraction of the day
				ADO_DefaultValue = CStr(Now)
				
			Case 133 ' A date value (yyyymmdd)
				ADO_DefaultValue = CStr(Now)
				
			Case 137 ' A database field time
				ADO_DefaultValue = CStr(Now)
				
			Case 134 ' A time value
				ADO_DefaultValue = CStr(Now)
				
			Case 135 'A date-time stamp (yyyymmddhhmmss plus a fractional in billionths)
				ADO_DefaultValue = CStr(Now)
				
			Case 14 ' An exact numeric value with fixed precision and scale
				ADO_DefaultValue = CStr(0)
				
			Case 5 'A double-precision floating point value
				ADO_DefaultValue = CStr(0)
				
			Case 0 'Unspecifid
				ADO_DefaultValue = " "
				
			Case 10 'A 32-bit error code
				ADO_DefaultValue = CStr(0)
				
			Case 64 'A DOS/WIN32 file time. The number of 100-nanosecond intervals since Jan 1 1601
				ADO_DefaultValue = CStr(0)
				
			Case 72 'A globally unique number
				ADO_DefaultValue = CStr(0)
				
			Case 9 'A pointer to an IDispatch interface on an OLE object
				ADO_DefaultValue = CStr(0)
				
			Case 3 'A 4-byte signed integer
				ADO_DefaultValue = CStr(0)
				
			Case 13 'A pointer to an IUnknown interface on an OLE object
				ADO_DefaultValue = CStr(0)
				
			Case 205 'A long binary value
				ADO_DefaultValue = CStr(0)
				
			Case 201 'A long string value
				ADO_DefaultValue = " "
				
			Case 203 'A null-terminated string value
				ADO_DefaultValue = CStr(0)
				
			Case 131 'An exact numeric value with a fixed precision and scale
				ADO_DefaultValue = CStr(0)
				
			Case 138 'A variant that is noot equivalent to an Automation variant
				ADO_DefaultValue = CStr(0)
				
			Case 4 'A single-precision floating point value
				ADO_DefaultValue = CStr(0)
				
			Case 2 'A 2-byte signed integer
				ADO_DefaultValue = CStr(0)
				
			Case 16 'A 1-byte signed integer
				ADO_DefaultValue = CStr(0)
				
			Case 21 'An 8 byte unsigned integer
				ADO_DefaultValue = CStr(0)
				
			Case 19 'An 4-byte unsigned integer
				ADO_DefaultValue = CStr(0)
				
			Case 18 'An 2-byte unsigned integer
				ADO_DefaultValue = CStr(0)
				
			Case 17 'An 1-byte unsigned integer
				ADO_DefaultValue = CStr(0)
				
			Case 132 'A user-defined variable
				ADO_DefaultValue = CStr(0)
				
			Case 204 'A binary value
				ADO_DefaultValue = CStr(0)
				
			Case 200 'A string value
				ADO_DefaultValue = " "
				
			Case 12 'An Automation Variant
				ADO_DefaultValue = CStr(0)
				
			Case 139 'A variable width exact numeric, with a signed scale value
				ADO_DefaultValue = CStr(0)
				
			Case 202 'A null-terminated Unicode character string
				ADO_DefaultValue = " "
				
			Case 130 'A null-terminated Unicode character string
				ADO_DefaultValue = " "
				
			Case Else 'Unknown type
				ADO_DefaultValue = " "
				
		End Select
		
	End Function
	
	Public Function ADO_FieldDelimiter(ByVal FieldType As Short) As String
		
        Dim ReturnValue As String = ""
		
		Select Case ADO_DefaultValue(FieldType)
			
			Case CStr(Now)
				ReturnValue = "#"
				
			Case "0"
				ReturnValue = ""
				
			Case " "
				ReturnValue = "'"
				
		End Select
		
		ADO_FieldDelimiter = ReturnValue
		
	End Function
	
	
	Public Sub SetTxtFieldText(ByVal Rs As ADODB.Recordset, ByRef Frm As System.Windows.Forms.Form, ByVal RsStart As Short, ByVal TxtStart As Short, Optional ByVal OffSet As Short = 0)
		
        '        Dim RsIdx As Short
        '		Dim TxtFieldIdx As Short

        '		Dim Off_Set As Short

        '		On Error GoTo ErrorHandler

        '		Off_Set = OffSet

        '		If Off_Set = -1 Then
        '			'UPGRADE_ISSUE: Control txtfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '            Off_Set = Frm.txtfield.UBound - TxtStart
        '		End If

        '		If Off_Set > (Rs.Fields.Count - 1) - RsStart Then
        '			Off_Set = (Rs.Fields.Count - 1) - RsStart
        '		End If

        '		If Off_Set = -1 Then
        '			Off_Set = 0
        '		End If

        '		RsIdx = RsStart
        '		TxtFieldIdx = TxtStart

        '		For RsIdx = RsStart To (RsStart + Off_Set)

        '			'UPGRADE_ISSUE: Control txtfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '			With Frm.txtfield(TxtFieldIdx)

        '				If Len(Rs.Fields(RsIdx).Value) > 0 Then
        '					'UPGRADE_ISSUE: Control txtfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '					.Text = CStr(Rs.Fields(RsIdx).Value)
        '				Else
        '					'UPGRADE_ISSUE: Control txtfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '					.Text = ""
        '				End If

        '			End With

        '			TxtFieldIdx = TxtFieldIdx + 1

        '		Next 

        '		Exit Sub

        'ErrorHandler: 
        '		Call LogFormError(ModuleName, "SetTxtField", Err.Description)
        '		Resume Next
		
	End Sub
	
	Public Sub SetTxtField(ByVal Rs As ADODB.Recordset, ByRef Frm As System.Windows.Forms.Form, ByVal RsStart As Short, ByVal TxtStart As Short, Optional ByVal OffSet As Short = 0)
		
        '		Dim RsIdx As Short
        '		Dim TxtFieldIdx As Short

        '		Dim Off_Set As Short

        '		On Error GoTo ErrorHandler

        '		Off_Set = OffSet

        '		If Off_Set = -1 Then
        '			'UPGRADE_ISSUE: Control txtfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '			Off_Set = Frm.txtfield.UBound - TxtStart
        '		End If

        '		If Off_Set > (Rs.Fields.Count - 1) - RsStart Then
        '			Off_Set = (Rs.Fields.Count - 1) - RsStart
        '		End If

        '		If Off_Set = -1 Then
        '			Off_Set = 0
        '		End If

        '		RsIdx = RsStart
        '		TxtFieldIdx = TxtStart

        '		For RsIdx = RsStart To (RsStart + Off_Set)

        '			'UPGRADE_ISSUE: Control txtfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '			With Frm.txtfield(TxtFieldIdx)

        '				'UPGRADE_ISSUE: Control txtfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '				.DataSource = Rs
        '				'UPGRADE_ISSUE: Control txtfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '				.DataField = Rs.Fields(RsIdx).Name

        '				'set text field max length, If field type is text
        '				If ADO_FieldDelimiter(Rs.Fields(RsIdx).Type) = "'" Then
        '					'UPGRADE_ISSUE: Control txtfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '					.MaxLength = Rs.Fields(RsIdx).DefinedSize
        '				End If

        '			End With

        '			TxtFieldIdx = TxtFieldIdx + 1

        '		Next 

        '		Exit Sub

        'ErrorHandler: 
        '		Call LogFormError(ModuleName, "SetTxtField", Err.Description)
        '		Resume Next
		
	End Sub
	
	Public Function RsFind(ByVal Rs As ADODB.Recordset, ByVal Fieldidx As Short, ByVal Value As String, Optional ByVal Delimiter As String = "") As ADODB.Recordset
		
		Dim Criteria As String
		
        Dim MyDelimiter As String = ""
		
		If Len(Delimiter) > 0 Then
			MyDelimiter = Delimiter
		End If
		
		If Len(MyDelimiter) = 0 Then
			MyDelimiter = ADO_FieldDelimiter(Rs.Fields(Fieldidx).Type)
		End If
		
		Criteria = Rs.Fields(Fieldidx).Name & "=" & MyDelimiter & Value & MyDelimiter
		
		With Rs
			If .RecordCount > 0 Then
				.MoveFirst()
				.Find(Criteria)
			End If
		End With
		
		RsFind = Rs
		
	End Function
	
	Public Sub SetlblField(ByVal Rs As ADODB.Recordset, ByRef Frm As System.Windows.Forms.Form, ByVal RsStart As Short, ByVal LblStart As Short, Optional ByVal OffSet As Short = 0)
		
        'Dim RsIdx As Short
        'Dim LblFieldIdx As Short

        'Dim Off_Set As Short

        'Off_Set = OffSet

        'If Off_Set = -1 Then
        '	'UPGRADE_ISSUE: Control lblfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '	Off_Set = Frm.lblfield.UBound - LblStart
        'End If

        'If Off_Set > (Rs.Fields.Count - 1) - RsStart Then
        '	Off_Set = (Rs.Fields.Count - 1) - RsStart
        'End If

        'RsIdx = RsStart
        'LblFieldIdx = LblStart

        'For RsIdx = RsStart To (RsStart + Off_Set)

        '	'UPGRADE_ISSUE: Control lblfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '	With Frm.lblfield(LblFieldIdx)
        '		'UPGRADE_ISSUE: Control lblfield could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        '		.Caption = "*" & Rs.Fields(RsIdx).Name & ": "
        '	End With

        '	LblFieldIdx = LblFieldIdx + 1

        'Next 
		
	End Sub
	
	Public Sub FillListBox(ByRef LstBox As System.Windows.Forms.ListBox, ByVal Rs As ADODB.Recordset, ByVal Fieldidx As Short)
		
		LstBox.Items.Clear()
		
		If Rs.RecordCount > 0 Then
			
			Rs.MoveFirst()
			
			Do Until Rs.EOF
				LstBox.Items.Add(Rs.Fields(Fieldidx).Value)
				Rs.MoveNext()
			Loop 
			
		End If
		
		LstBox.Refresh()
		
	End Sub
	
	Public Sub FillComboBox(ByRef CboBox As System.Windows.Forms.ComboBox, ByVal Rs As ADODB.Recordset, ByVal Fieldidx As Short)
		
		CboBox.Items.Clear()
		
		If Rs.RecordCount > 0 Then
			
			Rs.MoveFirst()
			
			Do Until Rs.EOF
				CboBox.Items.Add(Rs.Fields(Fieldidx).Value)
				Rs.MoveNext()
			Loop 
			
		End If
		
	End Sub
	
	Public Sub CopyADORecord(ByRef Source As ADODB.Recordset, ByRef Target As ADODB.Recordset, ByVal SourceStart As Short, ByVal TargetStart As Short, Optional ByVal OffSet As Short = 0, Optional ByVal ByName As Boolean = False, Optional ByVal FillDefault As Boolean = False)
		
		Dim SourceIdx As Short
		Dim TargetIdx As Short
		Dim idx As Short
		
		Dim Off_Set As Short
		
		On Error GoTo ErrorHandler
		
		Off_Set = OffSet
		
		If Off_Set = 0 Then
			Off_Set = Target.Fields.Count - TargetStart
		End If
		
		If Off_Set > Source.Fields.Count - SourceStart Then
			Off_Set = Source.Fields.Count - SourceStart
		End If
		
		SourceIdx = SourceStart
		TargetIdx = TargetStart
		
		If FillDefault = True Then
			
			For idx = TargetIdx To TargetIdx + Off_Set - 1
				Target.Fields(TargetIdx).Value = ADO_DefaultValue(Target.Fields(idx).Type)
			Next 
			
		End If
		
		For idx = SourceIdx To (SourceIdx + Off_Set - 1)
			
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(Source.Fields(idx).Value) = False Then
				
				Select Case ByName
					
					Case True
						If Source.Fields(idx).Name = Target.Fields(TargetIdx).Name Then
							Target.Fields(Source.Fields(idx).Name).Value = Source.Fields(idx).Value
						End If
						
					Case Else
						Target.Fields(Source.Fields(idx).Name).Value = Source.Fields(idx).Value
						
				End Select
				
			End If
			
			TargetIdx = TargetIdx + 1
			
		Next 
		
		Exit Sub
		
ErrorHandler: 
		Call LogFormError(ModuleName, "CopyADORecord", Err.Description)
		Resume Next
		
	End Sub
	
	Public Function OpenRs(ByVal Cn As ADODB.Connection, ByVal SQL As String, Optional ByVal CursorType As Short = 0, Optional ByVal LockType As Short = 0) As ADODB.Recordset
		
		Dim Rs As New ADODB.Recordset
		
		Dim t_CursorType As Short
		Dim t_LockType As Short
		
		t_CursorType = CursorType
		t_LockType = LockType
		
		If t_CursorType = 0 Then
			t_CursorType = 1
		End If
		
		If t_LockType = 0 Then
			t_LockType = 3
		End If
		
		Rs.Open(SQL, Cn, t_CursorType, t_LockType)
		
		OpenRs = Rs
		
	End Function
	
	Public Function RsToText(ByVal Rs As ADODB.Recordset, ByVal FilePath As String) As Boolean
		
		Dim idx As Short
		Dim Ts As Scripting.TextStream
		'UPGRADE_NOTE: WriteLine was upgraded to WriteLine_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim WriteLine_Renamed As String
		
		On Error GoTo ErrorHandler
		
		Ts = FSO.CreateTextFile(FilePath, True, False)
		
		With Rs
			
			WriteLine_Renamed = ""
			
			For idx = 0 To .Fields.Count - 1
				WriteLine_Renamed = WriteLine_Renamed & vbTab & .Fields(idx).Name
			Next 
			
			WriteLine_Renamed = Mid(WriteLine_Renamed, 2)
			
			Ts.WriteLine(WriteLine_Renamed)
			
			If .RecordCount > 0 Then
				
				.MoveFirst()
				
			End If
			
			Do Until .EOF
				
				WriteLine_Renamed = ""
				
				For idx = 0 To .Fields.Count - 1
					WriteLine_Renamed = WriteLine_Renamed & vbTab & .Fields(idx).Value
				Next 
				
				WriteLine_Renamed = Replace(WriteLine_Renamed & " ", Chr(13), "")
				WriteLine_Renamed = Mid(WriteLine_Renamed, 1, Len(WriteLine_Renamed) - 1)
				
				WriteLine_Renamed = Replace(WriteLine_Renamed & " ", Chr(10), "")
				WriteLine_Renamed = Mid(WriteLine_Renamed, 1, Len(WriteLine_Renamed) - 1)
				
				WriteLine_Renamed = Mid(WriteLine_Renamed, 2)
				
				Ts.WriteLine(WriteLine_Renamed)
				
				.MoveNext()
			Loop 
			
		End With
		
		Ts.Close()
		
		RsToText = True
		
		Exit Function
		
ErrorHandler: 
		
		Call LogFormError(ModuleName, "RsToText", Err.Description & vbCrLf & WriteLine_Renamed)
		MsgBox(Err.Description, MsgBoxStyle.Critical)
		RsToText = False
		
	End Function
End Module