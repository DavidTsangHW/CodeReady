Option Strict Off
Option Explicit On
Friend Class FrmSchema
	Inherits System.Windows.Forms.Form
	Public Cn As New ADODB.Connection
	Public Filename As String
	
	Dim Ts As Scripting.TextStream
	
	'UPGRADE_WARNING: Form event FrmSchema.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub FrmSchema_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		Me.Close()
		
	End Sub
	
	Private Sub FrmSchema_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		
		Ts = Fso.OpenTextFile(Filename, Scripting.IOMode.ForWriting, True)
		
		Call ReadTableName()
		
		Ts.Close()
		
		
	End Sub
	
	Private Sub ReadTableName()
		
		Dim SchemaRs As New ADODB.Recordset
		
		SchemaRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)
		
		
		
		
		'List all the tables in database
		
		'Q300948 BUG: Incorrect TABLE_TYPE Is Returned for Excel Worksheets
		'Check if excel driver
		
		System.Windows.Forms.Application.DoEvents()
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		Ts.WriteLine("Table name" & vbTab & "Field name")
		
		SchemaRs.MoveFirst()
		
		If InStr(1, LCase(Cn.ConnectionString), "excel") > 0 Or InStr(1, LCase(Cn.ConnectionString), ".xls") > 0 Then
			
			Do Until SchemaRs.EOF
				
				Call ReadTableFields(SchemaRs.Fields("Table_Name").Value)
				SchemaRs.MoveNext()
				
			Loop 
			
		Else
			
			Do Until SchemaRs.EOF
				
				'Q300948
				If SchemaRs.Fields("table_type").Value = "TABLE" Then
					
					Call ReadTableFields(SchemaRs.Fields("Table_Name").Value)
					
				End If
				
				SchemaRs.MoveNext()
				
			Loop 
			
		End If
		
		SchemaRs.Close()
		
		Cursor = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Private Sub ReadTableFields(ByVal TableName As String)
		
		Dim Rs As New ADODB.Recordset
		Dim SQL As String
		Dim TempString As String
		Dim Idx As Short
		
		On Error GoTo ErrorHandler
		
		System.Windows.Forms.Application.DoEvents()
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		SQL = "select top 1 * from [" & TableName & "]"
		
		Rs.Open(SQL, Cn, 1, 1)
		
		'TempString = TableName & vbTab
		
		For Idx = 0 To Rs.Fields.Count - 1
			
			TempString = TempString & TableName & vbTab & Rs.Fields(Idx).Name & vbCrLf
			
		Next 
		
		TempString = Mid(TempString, 1, Len(TempString) - 2)
		
		Ts.WriteLine(TempString)
		
		Rs.Close()
		
		'UPGRADE_NOTE: Object Rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Rs = Nothing
		
		Exit Sub
		
ErrorHandler: 
		
		MsgBox(Err.Number & ": " & Err.Description, MsgBoxStyle.Critical)
		
	End Sub
End Class