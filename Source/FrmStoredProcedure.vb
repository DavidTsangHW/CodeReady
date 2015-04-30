Option Strict Off
Option Explicit On
Friend Class FrmStoredProcedure
	Inherits System.Windows.Forms.Form
	Public FormGrid As New FrmGrid
	
	Private Sub cmdBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowse.Click
		
		'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Filter_Renamed As String
		Dim InitDir As String
		
        Dim Filepath As String

        Dim OpenFileDialog1 As New OpenFileDialog

		
		Filter_Renamed = "text (*.txt)|*.txt"
		Filter_Renamed = Filter_Renamed & "|sql (*.sql)|*.sql"
		Filter_Renamed = Filter_Renamed & "|All types|*.*"
		
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        With OpenFileDialog1
            'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            .Filter = Filter_Renamed
            .InitialDirectory = My.Application.Info.DirectoryPath
        End With
		
		'UPGRADE_ISSUE: MSComDlg.CommonDialog control CommonDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        TxtStoredProcedureFile.Text = OpenFileDialog(OpenFileDialog1)
		
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdExecute_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExecute.Click
		Dim Cmd As Object
		Dim Cn As New ADODB.Connection
		
		Dim SQL As String
		'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Command_Renamed As String
		Dim ResultLine As String
		
		Dim Counter As String
		Dim Message As String
		
		Dim Response As Short
		
		Dim IgnoreError As Boolean
		
		
		Dim Ts As Scripting.TextStream
		
		On Error GoTo ErrorHandler
		
		Ts = Fso.OpenTextFile(TxtStoredProcedureFile.Text, Scripting.IOMode.ForReading, False)
		
		Cn.Open(FormGrid.Adodc1.ConnectionString)
		
		Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		TxtResult.Text = "'"
		
		Counter = CStr(0)
		
		Do Until Ts.AtEndOfStream
			
			SQL = Ts.ReadLine
			SQL = Trim(SQL)
			
			If Not Len(Trim(SQL)) = 0 Then
				
				Counter = CStr(CDbl(Counter) + 1)
				
				Command_Renamed = Mid(SQL, 1, InStr(1, SQL, " ") - 1)
				
				Select Case LCase(Command_Renamed)
					
					Case "insert", "delete", "update", "drop", "create"
						
						Cn.Execute(SQL)
						
						'Updated 10 Feb 2007
						
					Case "select"
						
						'David Tsang, 10 Feb 2007
						If InStr(1, LCase(SQL), "into") > 0 Then
							
							Cn.Execute(SQL)
							
						Else
							
							With FormGrid.Adodc1
								
								.CommandType = ADODB.CommandTypeEnum.adCmdText
								.RecordSource = Trim(SQL)
								
								Call FormGrid.ShowStatus("Data updated")
								
								System.Windows.Forms.Application.DoEvents()
								.Refresh()
								
							End With
							
						End If
						
				End Select
				
				ResultLine = "Success - " & SQL
				TxtResult.Text = TxtResult.Text & ResultLine & vbNewLine & vbNewLine
				
			End If
			
		Loop 
		
		Ts.Close()
		
		
		ResultLine = " ------------------ Summary ------------------ " & vbNewLine
		ResultLine = ResultLine & "Executed: " & Counter & " Lines"
		TxtResult.Text = TxtResult.Text & ResultLine & vbNewLine
		
		Message = "Accomplished"
		MsgBox(Message, MsgBoxStyle.Information)
		
		
		'UPGRADE_NOTE: Object Ts may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Ts = Nothing
		
		Cn.Close()
		
		'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Cn = Nothing
		'UPGRADE_NOTE: Object Cmd may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Cmd = Nothing
		
		Cursor = System.Windows.Forms.Cursors.Default
		
		Exit Sub
		
ErrorHandler: 
		ResultLine = "Failed " & Err.Description & "- " & SQL
		TxtResult.Text = TxtResult.Text & ResultLine & vbNewLine
		Cursor = System.Windows.Forms.Cursors.Default
		
		If IgnoreError = False Then
			Message = Err.Description & vbCrLf
			Message = "Do you want to continue?"
			
			Response = MsgBox(Message, MsgBoxStyle.Question + MsgBoxStyle.YesNo)
			
			If Response = MsgBoxResult.No Then
				Exit Sub
			End If
			
			IgnoreError = True
		End If
		
		Resume Next
		
	End Sub
End Class