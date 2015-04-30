Option Strict Off
Option Explicit On
Module Bas_DataGrid
	Private Const ModuleName As String = "Bas_DataGrid"
	Public Sub SetDataGridColWidth(ByRef DG As AxMSDataGridLib.AxDataGrid)
		
		Dim idx As Double
		Dim Cnt As Double
		
		With DG
			Cnt = 0
			For idx = 0 To .Columns.Count - 1
				If .Columns(idx).Visible = True Then
					Cnt = Cnt + 1
				End If
			Next 
			
			For idx = 0 To .Columns.Count - 1
				
				If .Columns(idx).Visible = True Then
					.Columns(idx).Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(.Width) * 0.96) / Cnt)
				Else
					.Columns(idx).Width = 0
				End If
			Next 
			
		End With
		
	End Sub
	
	Public Sub SetDataGridStyle(ByRef DG As AxMSDataGridLib.AxDataGrid)
		
		With DG
			
            .RowDividerStyle = MSDataGridLib.DividerStyleConstants.dbgLightGrayLine
			.BorderStyle = MSDataGridLib.BorderStyleConstants.dbgFixedSingle
			
		End With
		
	End Sub
End Module