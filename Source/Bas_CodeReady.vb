Option Strict Off
Option Explicit On
Module Bas_CodeReady
	Private Const ModuleName As String = "Bas_CodeReady"
	
	Public Sub OpenTable(ByRef FormGrid As FrmGrid, ByRef CN As ADODB.Connection, ByRef TableName As String)
		Dim MousePointer As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MousePointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MousePointer = System.Windows.Forms.Cursors.WaitCursor
		
		On Error GoTo ErrorHandler
		
		With FormGrid
			
			With .Adodc1
				
				.ConnectionString = CN.ConnectionString
				.CommandType = ADODB.CommandTypeEnum.adCmdText
				.RecordSource = "select * from [" & TableName & "]"
				.Refresh()
				.Text = .Recordset.RecordCount & " records"
				
			End With
			
			With .DataGrid1
				.DataSource = FormGrid.Adodc1
				'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
                .Refresh()
			End With
			
			.Text = .Adodc1.RecordSource
			
		End With
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MousePointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MousePointer = System.Windows.Forms.Cursors.Default
		
		Exit Sub
		
ErrorHandler: 
		'UPGRADE_WARNING: Couldn't resolve default property of object MousePointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MousePointer = System.Windows.Forms.Cursors.Default
		MsgBox(Err.Description, MsgBoxStyle.Exclamation)
		Resume Next
		
	End Sub
End Module