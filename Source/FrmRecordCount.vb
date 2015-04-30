Option Strict Off
Option Explicit On
Friend Class FrmRecordCount
	Inherits System.Windows.Forms.Form
	''Code Ready 2006/12/06 22:01:38
	
	'Please add the following component to run this form
	'1. Microsoft Common Dialog Control 6.0, comdlg32.OCX
	'2. Microsoft Windows Common Control 6.0 (SP6), MSCOMCTL.OCX
	
	Public CN As New ADODB.Connection
	Dim SchemaRs As New ADODB.Recordset
	Dim TableIdx As Short
	
	Public FormGrid As New FrmGrid
	
	Private Sub FrmRecordCount_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Call ShowRecordCount()
		
	End Sub
	
	Private Sub FrmRecordCount_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		On Error Resume Next
		
		CN.Close()
		
	End Sub
	
	Private Sub ShowRecordCount()
		
		SchemaRs = CN.OpenSchema(ADODB.SchemaEnum.adSchemaTables)
		
		'List all the tables in database
		
		'Q300948 BUG: Incorrect TABLE_TYPE Is Returned for Excel Worksheets
		'Check if excel driver
		
		System.Windows.Forms.Application.DoEvents()
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		ListView1.Items.Clear()
		
		SchemaRs.MoveFirst()
		TableIdx = 0
		
		If InStr(1, LCase(CN.ConnectionString), "excel") > 0 Or InStr(1, LCase(CN.ConnectionString), ".xls") > 0 Then
			
			Do Until SchemaRs.EOF
				
				TableIdx = TableIdx + 1
				
				Call ShowTableRecordCount(SchemaRs.Fields("Table_Name").Value)
				
				SchemaRs.MoveNext()
				
			Loop 
			
		Else
			
			Do Until SchemaRs.EOF
				
				'Q300948
				If SchemaRs.Fields("table_type").Value = "TABLE" Then
					
					TableIdx = TableIdx + 1
					
					Call ShowTableRecordCount(SchemaRs.Fields("Table_Name").Value)
					
				End If
				
				SchemaRs.MoveNext()
				
			Loop 
			
		End If
		
		SchemaRs.Close()
		
		Cursor = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Private Sub ShowTableRecordCount(ByVal TableName As String)
		
		Dim NewItem As System.Windows.Forms.ListViewItem
		
		Dim Rs As New ADODB.Recordset
		Dim SQL As String
		
		'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        NewItem = ListView1.Items.Insert(TableIdx, TableName)
		
		On Error GoTo ErrorHandler
		
		System.Windows.Forms.Application.DoEvents()
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		SQL = "select * from [" & TableName & "]"
		
		Rs.Open(SQL, CN, 1, 1)
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(Rs.RecordCount) = False Then
			
			'UPGRADE_WARNING: Lower bound of collection NewItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			If NewItem.SubItems.Count > 1 Then
				NewItem.SubItems(1).Text = CStr(Rs.RecordCount)
			Else
				NewItem.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(Rs.RecordCount)))
			End If
			
		End If
		
		Rs.Close()
		
		'UPGRADE_NOTE: Object Rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Rs = Nothing
		
		Cursor = System.Windows.Forms.Cursors.Default
		
		Exit Sub
		
ErrorHandler: 
		Cursor = System.Windows.Forms.Cursors.Default
		MsgBox(Err.Description, MsgBoxStyle.Critical)
		
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
	
	'UPGRADE_WARNING: Event FrmRecordCount.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FrmRecordCount_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		
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
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		Select Case Button
			
			Case 2
				'UPGRADE_ISSUE: Form method FrmRecordCount.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                'PopupMenu(MnuPopUpMenu)
				
		End Select
		
	End Sub
	
	Public Sub MnuFont_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuFont.Click
		
		'UPGRADE_ISSUE: MSComDlg.CommonDialog control CommonDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        Dim FontDialog1 As New FontDialog

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
	
	
	
	Public Sub MnuRefresh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuRefresh.Click
		
		Call ShowRecordCount()
		
	End Sub
	
	Public Sub MnuSaveAs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuSaveAs.Click
		
		'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Filter_Renamed As String
		Dim InitDir As String
		
		Dim Filepath As String
		
		Dim Ts As Scripting.TextStream
		
        Dim Idx As Short

        Dim OpenFileDialog1 As New OpenFileDialog
        Dim SaveFileDialog1 As New SaveFileDialog

		
		On Error GoTo ErrorHandler
		
		Filter_Renamed = "Microsoft Excel (*.xls)|*.xls"
		Filter_Renamed = Filter_Renamed & "|Text (*.txt)|*.txt"
		Filter_Renamed = Filter_Renamed & "|All types|*.*"
		
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        With OpenFileDialog1
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
			
			Ts.WriteLine("Table Name" & vbTab & "Record Count")
			
			For Idx = 1 To ListView1.Items.Count
				
				'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				Ts.WriteLine(ListView1.Items.Item(Idx).Text & vbTab & ListView1.Items.Item(Idx).SubItems(1).Text)
				
			Next 
			
			Ts.Close()
			
			MsgBox("Saved as " & Filepath)
			
			Call Shell("Notepad.exe " & Filepath, AppWinStyle.NormalNoFocus)
			
		End If
		
		Cursor = System.Windows.Forms.Cursors.Default
		
		Exit Sub
		
ErrorHandler: 
		
		Cursor = System.Windows.Forms.Cursors.Default
		MsgBox(Err.Description, MsgBoxStyle.Critical)
		
	End Sub
	
	Public Sub MnuOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOpen.Click
		
		Call OpenTable((Me.FormGrid), CN, ListView1.FocusedItem.Text)
		
	End Sub
	
	
	
	Public Sub MnuOpenNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOpenNew.Click
		
		Dim MyFormGrid As New FrmGrid
		
		Call OpenTable(MyFormGrid, CN, ListView1.FocusedItem.Text)
		
	End Sub
End Class