Option Strict Off
Option Explicit On
Friend Class FrmTable
	Inherits System.Windows.Forms.Form
	Public ConnectionString As String
	Public FormGrid As New FrmGrid
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOpen.Click
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		On Error GoTo ErrorHandler
		
		With FormGrid
			
			With .Adodc1
				
				.ConnectionString = ConnectionString
				.CommandType = ADODB.CommandTypeEnum.adCmdText
				.RecordSource = "select * from [" & ListTable.Text & "]"
				.Refresh()
				.Text = .Recordset.RecordCount & " records"
				
			End With
			
			Call .ShowStatus("Data updated")
			
			With .DataGrid1
				.DataSource = FormGrid.Adodc1
				'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
                .Refresh()
			End With
			
			.Text = .Adodc1.RecordSource

            .Show()


		End With
		
		Me.Close()
		
		Cursor = System.Windows.Forms.Cursors.Default
		
		Exit Sub
		
ErrorHandler: 
		Cursor = System.Windows.Forms.Cursors.Default
		MsgBox(Err.Description, MsgBoxStyle.Exclamation)
		Resume Next
		
	End Sub
	
	Private Sub FrmTable_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		On Error GoTo ErrorHandler
		
		Dim Cn As New ADODB.Connection
		
		'an ADO recordset for storing database schema
		Dim SchemaRs As New ADODB.Recordset
		
		Cn.Open(ConnectionString)
		
		SchemaRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)
		
		'List all the tables in database
		
		'Q300948 BUG: Incorrect TABLE_TYPE Is Returned for Excel Worksheets
		'Check if excel driver
		
		ListTable.Items.Clear()
		
		If InStr(1, LCase(Cn.ConnectionString), "excel") > 0 Or InStr(1, LCase(Cn.ConnectionString), ".xls") > 0 Then
			
			Do Until SchemaRs.EOF
				
				ListTable.Items.Add(SchemaRs.Fields("table_name").Value)
				SchemaRs.MoveNext()
			Loop 
			
		Else
			
			Do Until SchemaRs.EOF
				'Q300948
				If SchemaRs.Fields("table_type").Value = "TABLE" Then
					ListTable.Items.Add(SchemaRs.Fields("table_name").Value)
				End If
				
				SchemaRs.MoveNext()
			Loop 
			
		End If
		SchemaRs.Close()
		
		Cn.Close()
		
		'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Cn = Nothing
		'UPGRADE_NOTE: Object SchemaRs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		SchemaRs = Nothing
		
		Exit Sub
		
ErrorHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation)
		
	End Sub
	
	Private Sub ListTable_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles ListTable.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 13 Then
			cmdOpen_Click(cmdOpen, New System.EventArgs())
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class