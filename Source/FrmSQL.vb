Option Strict Off
Option Explicit On
Friend Class FrmSQL
	Inherits System.Windows.Forms.Form
	Public FormGrid As New FrmGrid
	Dim Cn As New ADODB.Connection
	
	Private Sub CmdLog_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLog.Click
		
		Static Pressed As Boolean
		
		Select Case Pressed

            Case True
                CmdLog.Text = "Log >>"
                Me.Height = VB6.TwipsToPixelsY(2895)

            Case Else
                CmdLog.Text = "<< Log"
                Me.Height = VB6.TwipsToPixelsY(6240)
        End Select
		
		Pressed = Not Pressed
		
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdExecute_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExecute.Click
		
		
		Dim SQL As String
		'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Command_Renamed As String
		
		On Error GoTo ErrorHandler
		
		SQL = Trim(txtSQL.Text)
		
		SQL = Replace(SQL, Chr(13), "")
		
		If Len(TxtLog.SelectedText) > 0 Then
			
			SQL = TxtLog.SelectedText
			
		End If
		
		If Len(SQL) = 0 Then
			Exit Sub
		End If
		
		Call Execute(SQL)
		
		TxtLog.Text = TxtLog.Text & SQL & vbNewLine & vbNewLine
		
		Exit Sub
		
ErrorHandler: 
		
		MsgBox(Err.Number & ": " & Err.Description, MsgBoxStyle.Exclamation)
		Resume Next
		
		
	End Sub
	Private Sub Execute(ByRef SQL As String)
		
		'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Command_Renamed As String
		
		With FormGrid
			
			Command_Renamed = Mid(SQL, 1, InStr(1, SQL, " ") - 1)
			
			Select Case LCase(Command_Renamed)
				
				'David Tsang, 10 Feb 2007
				'Case drop create added
				Case "insert", "delete", "update", "drop", "create"
					
					Cn.Execute(SQL)
					
					
				Case "select"
					
					'David Tsang, 10 Feb 2007
					If InStr(1, LCase(SQL), "into") > 0 Then
						
						Cn.Execute(SQL)
						
					Else
						
						With .Adodc1
							
							.CommandType = ADODB.CommandTypeEnum.adCmdText
                            .RecordSource = SQL
                            .Refresh()
							.Text = .Recordset.RecordCount & " records"
							Call FormGrid.ShowStatus("Data updated")
							
						End With
						
						With .DataGrid1
							'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
                            .Refresh()
						End With
						
						.Text = SQL
						
					End If
					
			End Select
			
		End With
		
	End Sub
	Private Sub FrmSQL_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Cn.Open(FormGrid.Adodc1.ConnectionString)
		
	End Sub
	
	Private Sub FrmSQL_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		Cn.Close()
		
		'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Cn = Nothing
		
	End Sub
	
	Public Sub MnuExecute_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuExecute.Click
		
		Dim SQL As String
		
		On Error GoTo ErrorHandler
		
		SQL = TxtLog.SelectedText
		
		Call Execute(SQL)
		
		Exit Sub
		
ErrorHandler: 
		
		MsgBox(Err.Number & ": " & Err.Description)
		Resume Next
		
	End Sub
	
	Public Sub MnuSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuSave.Click
		Call SaveLog()
	End Sub
	
	Private Sub TxtLog_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles TxtLog.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If Button = VB6.MouseButtonConstants.RightButton Then
			Select Case Len(TxtLog.SelectedText)
				
				Case 0
					
					MnuExecute.Enabled = False
					
				Case Else
					
					MnuExecute.Enabled = True
					
			End Select
			
			
			'UPGRADE_ISSUE: Form method FrmSQL.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            'Call PopupMenu(MnuFile)
		End If
	End Sub
	
	Private Sub TxtSQL_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtSQL.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 13 Then
			cmdExecute_Click(cmdExecute, New System.EventArgs())
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub SaveLog()
		
		'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Filter_Renamed As String
		Dim InitDir As String
		
        Dim Filepath As String

        Dim SaveFileDialog1 As New SaveFileDialog
		
		Filter_Renamed = "text (*.txt)|*.txt"
		Filter_Renamed = Filter_Renamed & "|sql (*.sql)|*.sql"
		Filter_Renamed = Filter_Renamed & "|All types|*.*"
		
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        With SaveFileDialog1
            'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            .Filter = Filter_Renamed
            .InitialDirectory = My.Application.Info.DirectoryPath
        End With
		
		'UPGRADE_ISSUE: MSComDlg.CommonDialog control CommonDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        Filepath = GetSaveFileDialog(SaveFileDialog1)
		
		If Len(Filepath) > 0 Then
			If WriteNewFile(Filepath, TxtLog.Text) = True Then
				MsgBox("Saved successfully", MsgBoxStyle.Information)
			End If
		End If
		
	End Sub
End Class