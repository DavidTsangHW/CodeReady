Option Strict Off
Option Explicit On
Module Bas_CommDialog
	Private Const ModuleName As String = "Bas_CommDialog"
	
	'UPGRADE_NOTE: Argument type has been changed to Object. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D0BD8832-D1AC-487C-8AA5-B36F9284E51E"'
	Public Function OpenFileDialog(ByRef CDL as OpenFileDialog) As String
        On Error Resume Next
        With CDL

            If .ShowDialog() = DialogResult.OK Then

                OpenFileDialog = .FileName

            End If

        End With
    End Function
	
	
	'UPGRADE_NOTE: Return type was changed from MSComDlg.CommonDialog to Object Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CCD7D8F7-40F6-4AEF-8A92-C4C69D35ED6D"'
	'UPGRADE_NOTE: Argument type has been changed to Object. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D0BD8832-D1AC-487C-8AA5-B36F9284E51E"'
    Public Function OpenPrinterDialog(ByRef CDL As PrintDialog) As PrintDialog
        On Error Resume Next

        With CDL

            If .ShowDialog() = DialogResult.OK Then
                OpenPrinterDialog = CDL
            End If

        End With

    End Function
	
	'UPGRADE_NOTE: Return type was changed from MSComDlg.CommonDialog to Object Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CCD7D8F7-40F6-4AEF-8A92-C4C69D35ED6D"'
	'UPGRADE_NOTE: Argument type has been changed to Object. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D0BD8832-D1AC-487C-8AA5-B36F9284E51E"'
    Public Function OpenFontDialog(ByRef CDL As FontDialog) As FontDialog

        On Error Resume Next
        With CDL

            If .ShowDialog() = DialogResult.OK Then

                OpenFontDialog = CDL

            End If

        End With

    End Function
	
	'UPGRADE_NOTE: Return type was changed from MSComDlg.CommonDialog to Object Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CCD7D8F7-40F6-4AEF-8A92-C4C69D35ED6D"'
	'UPGRADE_NOTE: Argument type has been changed to Object. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D0BD8832-D1AC-487C-8AA5-B36F9284E51E"'
    Public Function OpenColorDialog(ByRef CDL As ColorDialog) As ColorDialog

        On Error Resume Next
        With CDL
            If .ShowDialog() = DialogResult.OK Then

                OpenColorDialog = CDL

            End If
        End With

    End Function
	
	
	'UPGRADE_NOTE: Argument type has been changed to Object. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D0BD8832-D1AC-487C-8AA5-B36F9284E51E"'
    Public Function GetSaveFileDialog(ByRef CDL As SaveFileDialog) As String
        On Error Resume Next

        With CDL
            If .ShowDialog() = DialogResult.OK Then

                GetSaveFileDialog = .FileName

            End If

        End With

    End Function

    Public Function FolderBrowserDialog(ByRef CDL As FolderBrowserDialog) As String


        Dim path As String
        Dim message As String

        path = ""

        message = "Enter path or press cancel to select"

        path = InputBox(message)

        If Len(path) > 0 Then

            FolderBrowserDialog = path

            Exit Function

        End If

        With CDL

            FolderBrowserDialog = ""

            If .ShowDialog() = DialogResult.OK Then

                FolderBrowserDialog = .SelectedPath

            End If

        End With

    End Function
End Module