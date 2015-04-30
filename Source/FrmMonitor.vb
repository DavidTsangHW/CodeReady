Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class FrmMonitor
	Inherits System.Windows.Forms.Form
	''Code Ready 2006/12/06 22:01:38
	
	'Please add the following component to run this form
	'1. Microsoft Common Dialog Control 6.0, comdlg32.OCX
	'2. Microsoft Windows Common Control 6.0 (SP6), MSCOMCTL.OCX
	
	Public CN As New ADODB.Connection
	Dim SchemaRs As New ADODB.Recordset
	Dim TableIdx As Short
	
	'UPGRADE_WARNING: Lower bound of array LastTableId was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim LastTableId(32767) As Short
	'UPGRADE_WARNING: Lower bound of array LastTableName was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim LastTableName(32767) As String
	'UPGRADE_WARNING: Lower bound of array LastTableRecordCount was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim LastTableRecordCount(32767) As Double
	Dim LastTableNameString As String
	Dim LastTableCount As Short
	
	'UPGRADE_WARNING: Lower bound of array TableId was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim TableId(32767) As Short
	'UPGRADE_WARNING: Lower bound of array TableName was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim TableName(32767) As String
	'UPGRADE_WARNING: Lower bound of array TableRecordCount was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim TableRecordCount(32767) As Double
	Dim TableNameString As String
	Dim TableCount As Short
	
	Public FormGrid As New FrmGrid
	
	Private Sub FrmMonitor_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Timer1.Enabled = True
		
	End Sub
	
	Private Sub Compare()
		
		Call GetDatabaseInformation()
		
		Dim idx As Short
		
		If Len(LastTableNameString) = 0 Then
			
			Exit Sub
			
		End If
		
		If LastTableNameString <> TableNameString Then
			
			Call UpdateTableIdx()
			
		End If
		
		For idx = 1 To TableCount
			
			If TableId(idx) > 0 And LastTableId(idx) > 0 Then
				
				
				If TableRecordCount(idx) <> LastTableRecordCount(TableId(idx)) Then
					
					System.Windows.Forms.Application.DoEvents()
					
					Call ShowRecordChanges(idx)
					
				End If
				
			End If
			
		Next 
		
	End Sub
	
	Private Sub UpdateTableIdx()
		
		Dim IdxA As Short
		Dim idxB As Short
		
		'Find table removed
		
		For IdxA = 1 To LastTableCount
			
			LastTableId(IdxA) = -1
			
		Next 
		
		For IdxA = 1 To LastTableCount
			
			For idxB = 1 To TableCount
				
				System.Windows.Forms.Application.DoEvents()
				
				If LastTableName(IdxA) = TableName(idxB) Then
					
					LastTableId(IdxA) = TableId(idxB)
					
					'Exit interior for
					'Use of "Exit For" exits all loops
					idxB = TableCount + 1
					
				End If
				
			Next 
			
		Next 
		
		Call ShowTableRemoved()
		
		'Find new table
		
		For IdxA = 1 To TableCount
			
			TableId(IdxA) = -1
			
		Next 
		
		For IdxA = 1 To TableCount
			
			For idxB = 1 To LastTableCount
				
				System.Windows.Forms.Application.DoEvents()
				
				If TableName(IdxA) = LastTableName(idxB) Then
					
					TableId(IdxA) = idxB
					
					'Exit interior for
					'Use of "Exit For" exits all loops
					idxB = LastTableCount + 1
					
				End If
				
			Next 
			
		Next 
		
		Call ShowTableAdded()
		
	End Sub
	
	Private Sub ShowTableRemoved()
		
		Dim idx As Short
		
		For idx = 1 To LastTableCount
			
			If LastTableId(idx) = -1 Then
				
				Call ShowTableChange(LastTableName(idx), "Table removed", CStr(TableRecordCount(idx)))
				
			End If
			
		Next 
		
	End Sub
	
	Private Sub ShowTableAdded()
		
		Dim idx As Short
		
		For idx = 1 To TableCount
			
			If TableId(idx) = -1 Then
				
				Call ShowTableChange(TableName(idx), "Table added", CStr(TableRecordCount(idx)))
				
			End If
			
		Next 
		
	End Sub
	
	Private Sub ShowRecordChanges(ByVal ArrayId As Short)
		
		Dim Activity As String
		Dim Different As Double
		
		If LastTableRecordCount(LastTableId(ArrayId)) > TableRecordCount(ArrayId) Then
			
			Different = LastTableRecordCount(LastTableId(ArrayId)) - TableRecordCount(ArrayId)
			Activity = Different & " records removed."
			
		ElseIf LastTableRecordCount(LastTableId(ArrayId)) < TableRecordCount(ArrayId) Then 
			
			Different = TableRecordCount(ArrayId) - LastTableRecordCount(LastTableId(ArrayId))
			Activity = Different & " records added."
			
		End If
		
		If Different < 2 And Different > -2 Then
			
			Activity = Replace(Activity, "records", "record")
			
		End If
		
		Call ShowChange(CStr(ArrayId), Activity)
		
		Exit Sub
		
ErrorHandler: 
		
		MsgBox(Err.Number & ": " & Err.Description, MsgBoxStyle.Critical)
		
	End Sub
	
	Private Sub ShowTableChange(ByVal Table_Name As String, ByVal Activity As String, ByVal RecordCount As String)
		
		Dim NewItem As System.Windows.Forms.ListViewItem
		
		On Error GoTo ErrorHandler
		
		System.Windows.Forms.Application.DoEvents()
		
		'Date Time
		'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		NewItem = ListView1.Items.Insert(ListView1.Items.Count + 1, CStr(Now))
		
		'Table name
		'UPGRADE_WARNING: Lower bound of collection NewItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		If NewItem.SubItems.Count > 1 Then
			NewItem.SubItems(1).Text = Table_Name
		Else
			NewItem.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, Table_Name))
		End If
		
		'Activity
		'UPGRADE_WARNING: Lower bound of collection NewItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		If NewItem.SubItems.Count > 2 Then
			NewItem.SubItems(2).Text = Activity
		Else
			NewItem.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, Activity))
		End If
		
		'Record Count
		'UPGRADE_WARNING: Lower bound of collection NewItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		If NewItem.SubItems.Count > 3 Then
			NewItem.SubItems(3).Text = RecordCount
		Else
			NewItem.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RecordCount))
		End If
		
		
		Exit Sub
		
ErrorHandler: 
		
		MsgBox(Err.Number & ": " & Err.Description, MsgBoxStyle.Critical)
		
	End Sub
	
	Private Sub ShowChange(ByVal ArrayId As String, ByVal Activity As String)
		
		Dim NewItem As System.Windows.Forms.ListViewItem
		
		On Error GoTo ErrorHandler
		
		System.Windows.Forms.Application.DoEvents()
		
		'Date Time
		'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		NewItem = ListView1.Items.Insert(ListView1.Items.Count + 1, CStr(Now))
		
		'Table name
		'UPGRADE_WARNING: Lower bound of collection NewItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		If NewItem.SubItems.Count > 1 Then
			NewItem.SubItems(1).Text = CStr(TableName(CInt(ArrayId)))
		Else
			NewItem.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(TableName(CInt(ArrayId)))))
		End If
		
		'Activity
		'UPGRADE_WARNING: Lower bound of collection NewItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		If NewItem.SubItems.Count > 2 Then
			NewItem.SubItems(2).Text = Activity
		Else
			NewItem.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, Activity))
		End If
		
		'Record Count
		'UPGRADE_WARNING: Lower bound of collection NewItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		If NewItem.SubItems.Count > 3 Then
			NewItem.SubItems(3).Text = CStr(TableRecordCount(CInt(ArrayId)))
		Else
			NewItem.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(TableRecordCount(CInt(ArrayId)))))
		End If
		
		
		Exit Sub
		
ErrorHandler: 
		
		MsgBox(Err.Number & ": " & Err.Description, MsgBoxStyle.Critical)
		
	End Sub
	
	'UPGRADE_NOTE: Update was upgraded to Update_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Update_Renamed()
		
		Dim idx As Short
		
		LastTableCount = TableCount
		LastTableNameString = TableNameString
		
		For idx = 1 To TableCount
			
			LastTableId(idx) = TableId(idx)
			LastTableRecordCount(idx) = TableRecordCount(idx)
			LastTableName(idx) = TableName(idx)
			
		Next 
		
		
	End Sub
	
	Private Sub FrmMonitor_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		On Error Resume Next
		
		CN.Close()
		
	End Sub
	
	Private Sub GetDatabaseInformation()
		
		On Error GoTo ErrorHandler
		
		SchemaRs = CN.OpenSchema(ADODB.SchemaEnum.adSchemaTables)
		
		'List all the tables in database
		
		'Q300948 BUG: Incorrect TABLE_TYPE Is Returned for Excel Worksheets
		'Check if excel driver
		
		System.Windows.Forms.Application.DoEvents()
		
		SchemaRs.MoveFirst()
		TableIdx = 0
		
		TableNameString = ""
		
		If InStr(1, LCase(CN.ConnectionString), "excel") > 0 Or InStr(1, LCase(CN.ConnectionString), ".xls") > 0 Then
			
			Do Until SchemaRs.EOF
				
				Call GetTableRecordCount(SchemaRs.Fields("Table_Name").Value)
				
				SchemaRs.MoveNext()
				
			Loop 
			
		Else
			
			Do Until SchemaRs.EOF
				
				'Q300948
				If SchemaRs.Fields("table_type").Value = "TABLE" Then
					
					Call GetTableRecordCount(SchemaRs.Fields("Table_Name").Value)
					
				End If
				
				SchemaRs.MoveNext()
				
			Loop 
			
		End If
		
		TableCount = TableIdx
		
		SchemaRs.Close()
		
		Cursor = System.Windows.Forms.Cursors.Default
		
		Exit Sub
		
ErrorHandler: 
		
		Exit Sub
		
		
	End Sub
	
	Private Sub GetTableRecordCount(ByVal Table_Name As String)
		
		Dim NewItem As System.Windows.Forms.ListViewItem
		
		Dim Rs As New ADODB.Recordset
		Dim SQL As String
		
		On Error GoTo ErrorHandler
		
		System.Windows.Forms.Application.DoEvents()
		
		SQL = "select * from [" & Table_Name & "]"
		
		Rs.Open(SQL, CN, 1, 1)
		
		TableIdx = TableIdx + 1
		
		TableRecordCount(TableIdx) = Rs.RecordCount
		
		Rs.Close()
		
		'UPGRADE_NOTE: Object Rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Rs = Nothing
		
		TableName(TableIdx) = Table_Name
		
		TableId(TableIdx) = TableIdx
		
		TableNameString = TableNameString & Table_Name
		
		Cursor = System.Windows.Forms.Cursors.Default
		
		Exit Sub
		
ErrorHandler: 
		
		Cursor = System.Windows.Forms.Cursors.Default
		Exit Sub
		
	End Sub
	
	
	Private Sub ListView1_ColumnClick(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
		Dim ColumnHeader As System.Windows.Forms.ColumnHeader = ListView1.Columns(eventArgs.Column)
		
		With ListView1
			
			'UPGRADE_ISSUE: MSComctlLib.ListView property ListView1.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            If .Sorting = ColumnHeader.Index - 1 Then

                If .Sorting = System.Windows.Forms.SortOrder.Ascending Then
                    .Sorting = System.Windows.Forms.SortOrder.Descending
                Else
                    .Sorting = System.Windows.Forms.SortOrder.Ascending
                End If

            Else

                'UPGRADE_ISSUE: MSComctlLib.ListView property ListView1.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                .Sorting = ColumnHeader.Index - 1
                .Sorting = System.Windows.Forms.SortOrder.Ascending

            End If

            .Sort()

        End With
		
	End Sub
	
	'UPGRADE_WARNING: Event FrmMonitor.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FrmMonitor_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		
		On Error Resume Next
		
		With ListView1
			
			.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 500)
			.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 1000)
			
		End With
		
	End Sub
	
	Private Sub ListView1_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		Select Case Button
			
			Case 2
				'UPGRADE_ISSUE: Form method FrmMonitor.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                'PopupMenu(MnuPopUpMenu)
				
		End Select
		
	End Sub
	
	Public Sub MnuClear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuClear.Click
		
		On Error GoTo ErrorHandler
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		ListView1.Items.Clear()
		
		Cursor = System.Windows.Forms.Cursors.Default
		
		Exit Sub
		
ErrorHandler: 
		
		MsgBox(Err.Number & ": " & Err.Description)
		
	End Sub
	
	Public Sub MnuFont_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuFont.Click

        Dim FontDialog1 As New FontDialog

		'UPGRADE_ISSUE: MSComDlg.CommonDialog control CommonDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        Call OpenFontDialog(FontDialog1)
		
		With ListView1
			.Font = VB6.FontChangeName(.Font, CommonDialog1Font.Font.Name)
			.Font = VB6.FontChangeSize(.Font, CommonDialog1Font.Font.Size)
			.Font = VB6.FontChangeUnderline(.Font, CommonDialog1Font.Font.Underline)
			.Font = VB6.FontChangeItalic(.Font, CommonDialog1Font.Font.Italic)
			.ForeColor = CommonDialog1Color.Color
		End With
		
	End Sub
	
	Public Sub MnuGrid_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuGrid.Click
		
		ListView1.GridLines = Not ListView1.GridLines
		MnuGrid.Checked = ListView1.GridLines
		
	End Sub
	
	
	Public Sub MnuOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOpen.Click
		
		'UPGRADE_WARNING: Lower bound of collection ListView1.SelectedItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		Call OpenTable((Me.FormGrid), CN, ListView1.FocusedItem.SubItems(1).Text)
		
	End Sub
	
	Public Sub MnuOpenNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOpenNew.Click
		
		Dim MyFormGrid As New FrmGrid
		
		'UPGRADE_WARNING: Lower bound of collection ListView1.SelectedItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		Call OpenTable(MyFormGrid, CN, ListView1.FocusedItem.SubItems(1).Text)
		
	End Sub
	
	Public Sub MnuPause_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuPause.Click
		
		LastTableNameString = ""
		
		System.Windows.Forms.Application.DoEvents()
		Timer1.Enabled = False
		
		System.Windows.Forms.Application.DoEvents()
		MnuPause.Visible = False
		
		System.Windows.Forms.Application.DoEvents()
		MnuStart.Visible = True
		
	End Sub
	
	Public Sub MnuSaveAs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuSaveAs.Click
		
		'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Filter_Renamed As String
		Dim InitDir As String
		
		Dim Filepath As String
		
		Dim Ts As Scripting.TextStream
		
		Dim idx As Short
		Dim idxB As Short
		
        Dim TempString As String

        Dim SaveFileDialog1 As New SaveFileDialog

		
		On Error GoTo ErrorHandler
		
		Filter_Renamed = "Microsoft Excel (*.xls)|*.xls"
		Filter_Renamed = Filter_Renamed & "|Text (*.txt)|*.txt"
		Filter_Renamed = Filter_Renamed & "|All types|*.*"
		
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        With SaveFileDialog1
            'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            .Filter = Filter_Renamed
            .InitialDirectory = My.Application.Info.DirectoryPath
        End With
		
		'UPGRADE_ISSUE: MSComDlg.CommonDialog control CommonDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        Filepath = GetSaveFileDialog(SaveFileDialog1)
		
		If Len(Filepath) = 0 Then
			
			Exit Sub
			
		End If
		
		System.Windows.Forms.Application.DoEvents()
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		Call Fso.CreateTextFile(Filepath, True)
		
		If Fso.FileExists(Filepath) = True Then
			
			Ts = Fso.OpenTextFile(Filepath, Scripting.IOMode.ForWriting, True)
			
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			TempString = ListView1.Columns.Item(1).Text
			
			For idxB = 2 To ListView1.Columns.Count
				
				'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				TempString = TempString & vbTab & ListView1.Columns.Item(idxB).Text
				
			Next 
			
			Ts.WriteLine(TempString)
			
			For idx = 1 To ListView1.Items.Count
				
				'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				TempString = ListView1.Items.Item(idx).Text
				
				For idxB = 1 To ListView1.Columns.Count - 1
					
					'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					TempString = TempString & vbTab & ListView1.Items.Item(idx).SubItems(idxB).Text
					
				Next 
				
				Ts.WriteLine(TempString)
				
			Next 
			
			Ts.Close()
			
			MsgBox("Saved as " & Filepath)
			
			Call Shell("Notepad.exe " & Filepath, AppWinStyle.NormalNoFocus)
			
		End If
		
		Cursor = System.Windows.Forms.Cursors.Default
		
		Exit Sub
		
ErrorHandler: 
		
		Cursor = System.Windows.Forms.Cursors.Default
		Resume Next
		
	End Sub
	
	Public Sub MnuStart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuStart.Click
		
		System.Windows.Forms.Application.DoEvents()
		Timer1.Enabled = True
		
		System.Windows.Forms.Application.DoEvents()
		MnuPause.Visible = True
		
		System.Windows.Forms.Application.DoEvents()
		MnuStart.Visible = True
		
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		
		Static IsProcessing As Boolean
		Dim StartTime As Double
		Dim EndTime As Double
		
		System.Windows.Forms.Application.DoEvents()
		
		If IsProcessing = True Then
			
			Timer1.Enabled = False
			Exit Sub
			
		End If
		
		StartTime = VB.Timer()
		
		Timer1.Enabled = False
		IsProcessing = True
		
		Call Compare()
		Call Update_Renamed()
		
		EndTime = VB.Timer()
		
		'UPGRADE_WARNING: Timer property Timer1.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
		Timer1.Interval = EndTime - StartTime
		
		If Timer1.Interval < 1000 Then
			
			Timer1.Interval = 1000
			
		End If
		
		Timer1.Enabled = True
		IsProcessing = False
		
		
	End Sub
End Class