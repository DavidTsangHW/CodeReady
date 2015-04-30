Option Strict Off
Option Explicit On
Friend Class FrmGrid
	Inherits System.Windows.Forms.Form
	Dim FindStr As String
	Dim Criteria As String
	
	Private Sub DataGrid1_HeadClick(ByVal eventSender As System.Object, ByVal eventArgs As AxMSDataGridLib.DDataGridEvents_HeadClickEvent) Handles DataGrid1.HeadClick
		
		Static Field As String
		Static IsAsc As Boolean
		
		Dim Criteria As String
		
		On Error GoTo ErrorHandler
		
		With DataGrid1.Columns(eventArgs.ColIndex)
			
			If Field = .DataField Then
				
				IsAsc = Not IsAsc
				
			Else
				
				Field = .DataField
				IsAsc = True
				
			End If
			
		End With
		
		Select Case IsAsc
			
			Case True
				Criteria = Field & " ASC"
				
			Case False
				Criteria = Field & " DESC"
				
		End Select
		
		With Adodc1.Recordset
			
			.Sort = Criteria
			
		End With
		
		Exit Sub
		
ErrorHandler: 
		
		MsgBox(Err.Number & ": " & Err.Description, MsgBoxStyle.Critical)
		
	End Sub
	
	
	
	Private Sub FrmGrid_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		If IsLicensed = True Then
			
			MnuRegister.Visible = False
			
		End If
		
	End Sub
	

	Public Sub MnuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuAbout.Click
		
		Dim FormAbout As New frmAbout
		
		With FormAbout
			.ShowDialog()
		End With
		
	End Sub
	
	Public Sub MnuApperance_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuApperance.Click
		
		'UPGRADE_ISSUE: MSComDlg.CommonDialog control CommonDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        Call OpenColorDialog(CommonDialog1Color)
		
		With DataGrid1
			.BackColor = CommonDialog1Color.Color
		End With
		
	End Sub
	
	Public Sub MnuASPBlankForm_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuASPBlankForm.Click
		
		Dim Rs As ADODB.Recordset
		Dim Filename As String
		'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Filter_Renamed As String
		Dim Message As String
		
		On Error GoTo ErrorHandler
		
		Rs = Adodc1.Recordset.Clone
		
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        With SaveFileDialog1

            .DefaultExt = "Default.htm"
            .Title = "HTML"
            .FileName = "default.htm"
            Filter_Renamed = "HTML document| *.htm|"
            Filter_Renamed = Filter_Renamed & "All types|*.*|"
            'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            .Filter = Filter_Renamed

        End With
		
		'UPGRADE_ISSUE: MSComDlg.CommonDialog control CommonDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        Filename = GetSaveFileDialog(SaveFileDialog1)
		
		If Len(Filename) > 0 Then
			Cursor = System.Windows.Forms.Cursors.WaitCursor
			Call BuildASPBlankForm(Adodc1.ConnectionString, Rs, Filename)
			Cursor = System.Windows.Forms.Cursors.Default
		Else
			Exit Sub
		End If
		
		Message = "Exported as " & Filename
		MsgBox(Message, MsgBoxStyle.Information)
		
		If MsgBox("Open " & Filename & "?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.Yes Then
			
			Shell("notepad.exe " & Filename, AppWinStyle.NormalFocus)
			
		End If
		
		Exit Sub
		
ErrorHandler: 
		Cursor = System.Windows.Forms.Cursors.Default
		LogFormError(Me.Name, "MnuExportVisualBasic6Form", Err.Description)
		MsgBox(Err.Description, MsgBoxStyle.Critical)
		
	End Sub
	
	Public Sub MnuASPEditForm_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuASPEditForm.Click
		
		Static Path As String
		Dim tempPath As String
		Dim Rs As New ADODB.Recordset


        On Error GoTo ErrorHandler



        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Len(Trim(tempPath)) = 0 Then
            Exit Sub
        End If

        If Fso.FolderExists(tempPath) = False Then
            MsgBox("Path not found", MsgBoxStyle.Critical)
        End If

        Path = tempPath

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        Rs = Adodc1.Recordset.Clone

        Call BuildASPEditForm(Adodc1.ConnectionString, Rs, Path)

        Cursor = System.Windows.Forms.Cursors.Default

        MsgBox("Completed", MsgBoxStyle.Information)

        Exit Sub

ErrorHandler:
        Cursor = System.Windows.Forms.Cursors.Default
        LogFormError(Me.Name, "MnuASPEditForm", Err.Description)
        MsgBox(Err.Description, MsgBoxStyle.Critical)

    End Sub

    Public Sub MnuASPGrid_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuASPGrid.Click

        Static Path As String
        Dim tempPath As String
        Dim Rs As New ADODB.Recordset


        On Error GoTo ErrorHandler



        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Len(Trim(tempPath)) = 0 Then
            Exit Sub
        End If

        If Fso.FolderExists(tempPath) = False Then
            MsgBox("Path not found", MsgBoxStyle.Critical)
        End If

        Path = tempPath

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        Rs = Adodc1.Recordset.Clone

        Call BuildASPGrid(Adodc1.ConnectionString, Rs, Path)

        Cursor = System.Windows.Forms.Cursors.Default

        MsgBox("Completed", MsgBoxStyle.Information)

        Exit Sub

ErrorHandler:
        Cursor = System.Windows.Forms.Cursors.Default
        LogFormError(Me.Name, "MnuASPGrid", Err.Description)
        MsgBox(Err.Description, MsgBoxStyle.Critical)


    End Sub

    Public Sub MnuASPProjectBlankForm_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuASPProjectBlankForm.Click

        Dim tempPath As String
        Dim Cn As New ADODB.Connection
        Dim SchemaRs As New ADODB.Recordset

        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Fso.FolderExists(tempPath) = False Then

            Exit Sub

        End If

        Cn.Open(Adodc1.ConnectionString)

        SchemaRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        Call CreateASPForms(tempPath, SchemaRs, Adodc1.ConnectionString, "blankform")

        SchemaRs.Close()

        Cn.Close()

        'UPGRADE_NOTE: Object SchemaRs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        SchemaRs = Nothing
        'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Cn = Nothing

        MsgBox("Blank forms have been created successfully", MsgBoxStyle.Information)

    End Sub

    Public Sub MnuASPProjectEditForm_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuASPProjectEditForm.Click

        Dim tempPath As String
        Dim SchemaRs As New ADODB.Recordset

        Dim Cn As New ADODB.Connection


        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Fso.FolderExists(tempPath) = False Then

            Exit Sub

        End If

        Cn.Open(Adodc1.ConnectionString)

        SchemaRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        Call CreateASPForms(tempPath, SchemaRs, Adodc1.ConnectionString, "editform")


        SchemaRs.Close()

        Cn.Close()

        'UPGRADE_NOTE: Object SchemaRs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        SchemaRs = Nothing
        'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Cn = Nothing

        MsgBox("Edit forms have been created successfully", MsgBoxStyle.Information)

    End Sub

    Public Sub MnuASPProjectGrid_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuASPProjectGrid.Click

        Dim tempPath As String
        Dim SchemaRs As New ADODB.Recordset

        Dim Cn As New ADODB.Connection

        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Fso.FolderExists(tempPath) = False Then

            Exit Sub

        End If

        Cn.Open(Adodc1.ConnectionString)

        SchemaRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        Call CreateASPForms(tempPath, SchemaRs, Adodc1.ConnectionString, "grid")

        SchemaRs.Close()

        Cn.Close()

        'UPGRADE_NOTE: Object SchemaRs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        SchemaRs = Nothing
        'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Cn = Nothing

        MsgBox("Grids have been created successfully", MsgBoxStyle.Information)

    End Sub

    Public Sub MnuASPProjectReportList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuASPProjectReportList.Click

        Dim tempPath As String
        Dim Cn As New ADODB.Connection
        Dim SchemaRs As New ADODB.Recordset

        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Fso.FolderExists(tempPath) = False Then

            Exit Sub

        End If

        Cn.Open(Adodc1.ConnectionString)

        SchemaRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        Call CreateASPForms(tempPath, SchemaRs, Adodc1.ConnectionString, "print")

        SchemaRs.Close()

        Cn.Close()

        'UPGRADE_NOTE: Object SchemaRs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        SchemaRs = Nothing
        'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Cn = Nothing

        MsgBox("Report lists have been created successfully", MsgBoxStyle.Information)

    End Sub

    Public Sub MnuCodeReadyHomepage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCodeReadyHomepage.Click

        Call OpenBrowser("http://www24.brinkster.com/david6648668/projects/codeready")

    End Sub

    Public Sub MnuEnterConnectionString_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuEnterConnectionString.Click

        Dim FormGrid As New FrmGrid
        Dim FormTable As New FrmTable
        Dim Connstr As String
        Dim Message As String

        Message = "Enter Connection String"

        Connstr = InputBox(Message, "Connection String", Adodc1.ConnectionString)

        'Open connection
        'If cancel was pressed, the connectionstring will become null
        If Len(Connstr) = 0 Then
            Exit Sub
        End If

        With FormTable
            .ConnectionString = Connstr
            .FormGrid = FormGrid
            .ShowDialog()
        End With

    End Sub

    Public Sub MnuExportTextFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuExportTextFile.Click

        'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Filter_Renamed As String
        Dim InitDir As String

        Dim Filepath As String

        On Error GoTo ErrorHandler

        Filter_Renamed = "Text (*.txt)|*.txt"
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

            Cursor = System.Windows.Forms.Cursors.WaitCursor
            System.Windows.Forms.Application.DoEvents()
            Call ShowStatus("Exporting as " & Filepath & ". Please wait ...")
            Call RsToText(Adodc1.Recordset.Clone, Filepath)
            Call ShowStatus("Exported as " & Filepath)
            MsgBox("Exported as " & Filepath)
            Call Shell("Notepad.exe " & Filepath, AppWinStyle.NormalNoFocus)

        End If


        Cursor = System.Windows.Forms.Cursors.Default

        Exit Sub

ErrorHandler:
        Cursor = System.Windows.Forms.Cursors.Default
        MsgBox(Err.Description, MsgBoxStyle.Critical)

    End Sub

    Public Sub MnuFind_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuFind.Click

        'UPGRADE_NOTE: Str was upgraded to Str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Str_Renamed As String
        Dim Message As String
        Dim bmark As String
        Dim Delimiter As String

        On Error GoTo ErrorHandler

        With DataGrid1
            Message = "Find in " & .Columns(.Col).Caption
            Str_Renamed = InputBox(Message, Message, FindStr)

            If Len(Str_Renamed) = 0 Then
                Exit Sub
            End If

            FindStr = Str_Renamed

            Delimiter = ADO_FieldDelimiter(Adodc1.Recordset.Fields(.Col).Type)

            Criteria = .Columns(.Col).Caption & " like "
            Criteria = Criteria & Delimiter

            Select Case Delimiter

                Case "'"
                    Criteria = Criteria & FindStr & "*"

                Case Else

                    Criteria = Criteria & FindStr

            End Select

            Criteria = Criteria & Delimiter

        End With

        Call Find()

        Exit Sub

ErrorHandler:
        Message = "Record not found"
        MsgBox(Message, MsgBoxStyle.Information)

    End Sub

    Private Sub Find()
        Dim bmark As Object

        Dim Message As String

        On Error GoTo ErrorHandler

        With Adodc1.Recordset

            'UPGRADE_WARNING: Couldn't resolve default property of object Adodc1.Recordset.Bookmark. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object bmark. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            bmark = .Bookmark

            .Find(Criteria, .AbsolutePosition, ADODB.SearchDirectionEnum.adSearchForward)

            If .EOF = True Then
                Message = "Record not found"
                MsgBox(Message, MsgBoxStyle.Information)

                On Error Resume Next
                'UPGRADE_WARNING: Couldn't resolve default property of object bmark. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Bookmark = bmark
                'UPGRADE_WARNING: Couldn't resolve default property of object bmark. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                bmark = ""
            End If

        End With

        Exit Sub

ErrorHandler:
        Message = "Record not found"
        MsgBox(Message, MsgBoxStyle.Information)



    End Sub

    Public Sub MnuFindNext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuFindNext.Click

        Call Find()

    End Sub

    Public Sub MnuFonts_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuFonts.Click

        'UPGRADE_ISSUE: MSComDlg.CommonDialog control CommonDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        Call OpenFontDialog(FontDialog1)

        'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        With FontDialog1

            DataGrid1.Font = VB6.FontChangeName(DataGrid1.Font, .Font.Name)
            DataGrid1.Font = VB6.FontChangeItalic(DataGrid1.Font, .Font.Italic)
            DataGrid1.Font = VB6.FontChangeSize(DataGrid1.Font, .Font.Size)
            DataGrid1.Font = VB6.FontChangeBold(DataGrid1.Font, .Font.Bold)
            DataGrid1.Font = VB6.FontChangeUnderline(DataGrid1.Font, .Font.Underline)
            DataGrid1.Font = VB6.FontChangeStrikeout(DataGrid1.Font, .Font.Strikeout)
            DataGrid1.ForeColor = .Color

        End With

    End Sub

    Public Sub MnuMonitorAddRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuMonitorAddRemove.Click
        Dim FormGrid As Object

        Dim FormMonitor As New FrmMonitor

        With FormMonitor

            Cursor = System.Windows.Forms.Cursors.WaitCursor

            .FormGrid = Me

            .CN.Open(Adodc1.ConnectionString)

            .Show()

            Cursor = System.Windows.Forms.Cursors.Default

        End With

    End Sub

    Public Sub MnuMultipleStatement_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuMultipleStatement.Click
        Dim FormGrid As Object
        Dim FormStoredProcedure As New FrmStoredProcedure

        With FormStoredProcedure

            .FormGrid = Me

            .Show()
        End With

    End Sub

    Public Sub MnuRecordCount_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuRecordCount.Click
        Dim FormGrid As Object

        Dim FormRecordCount As New FrmRecordCount

        With FormRecordCount

            .FormGrid = Me

            .CN.Open(Adodc1.ConnectionString)
            .Show()

        End With

    End Sub

    Public Sub MnuSaveConnectionString_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuSaveConnectionString.Click

        'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Filter_Renamed As String
        Dim InitDir As String

        Dim Filepath As String

        Dim Ts As Scripting.TextStream

        Dim Idx As Short
        Dim idxB As Short

        Dim TempString As String

        On Error GoTo ErrorHandler

        Filter_Renamed = "Text (*.txt)|*.txt"
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

        Ts = Fso.OpenTextFile(Filepath, Scripting.IOMode.ForWriting, True)

        Ts.WriteLine(Adodc1.ConnectionString)

        Ts.Close()

        MsgBox("Connection string has been saved successfully", MsgBoxStyle.Information)

        Exit Sub

ErrorHandler:

        MsgBox(Err.Number & ": " & Err.Description, MsgBoxStyle.Critical)

    End Sub

    Public Sub MnuSchema_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuSchema.Click

        Dim FormSchema As New FrmSchema

        'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Filter_Renamed As String
        Dim InitDir As String

        Dim Filepath As String

        Dim Ts As Scripting.TextStream

        Dim Idx As Short
        Dim idxB As Short

        Dim TempString As String

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

        With FormSchema

            .Cn.Open(Me.Adodc1.ConnectionString)
            .Filename = Filepath
            .ShowDialog()

        End With

        MsgBox("Schema has been created successfully", MsgBoxStyle.Information)

        Call Shell("Notepad.exe " & Filepath, AppWinStyle.NormalNoFocus)

        Exit Sub

ErrorHandler:

        MsgBox(Err.Number & ": " & Err.Description, MsgBoxStyle.Critical)

    End Sub

    Public Sub MnuSinglStatement_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuSinglStatement.Click
        'Dim FormGrid As Object 
        Dim FormSQL As New FrmSQL
        'Dim SQL As String

        With FormSQL

            .FormGrid = Me

            'UPGRADE_WARNING: TextRTF was upgraded to Text and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            .txtSQL.Text = Me.Adodc1.RecordSource

            .Show()

        End With

    End Sub


    Public Sub MnuTable_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuTable.Click
        'Dim FormGrid As Object

        Dim FormTable As New FrmTable

        With FormTable

            .FormGrid = Me

            .ConnectionString = Adodc1.ConnectionString

            .Show()

        End With

    End Sub

    Public Sub MnuCascade_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCascade.Click
        MDIForm1.LayoutMdi(System.Windows.Forms.MdiLayout.Cascade)
    End Sub

    Public Sub MnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuExit.Click
        MDIForm1.Close()
    End Sub

    Public Sub MnuOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOpen.Click

        Dim FormGrid As New FrmGrid
        Dim FormTable As New FrmTable
        Dim Connstr As String

        Dim Message As String
        Dim Response As Short

        'UPGRADE_WARNING: Couldn't resolve default property of object GetDataLinks. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Connstr = GetDataLinks()

        'Open connection
        'If cancel was pressed, the connectionstring will become null
        If Len(Connstr) = 0 Then
            Exit Sub
        End If

        Message = "Build this connection?"
        Message = Message & vbCr
        Message = Message & Connstr
        Response = MsgBox(Message, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel)

        If Not Response = MsgBoxResult.Yes Then
            Message = "Enter Connection String"
            Connstr = InputBox(Message, "Connection String", Connstr)
            If Len(Connstr) = 0 Then
                Exit Sub
            End If
        End If

        With FormTable
            .ConnectionString = Connstr
            .FormGrid = FormGrid
            .ShowDialog()
        End With


    End Sub

    Public Sub MnuTileHorizontally_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuTileHorizontally.Click
        MDIForm1.LayoutMdi(System.Windows.Forms.MdiLayout.TileHorizontal)
    End Sub

    Public Sub MnuTileVertically_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuTileVertically.Click
        MDIForm1.LayoutMdi(System.Windows.Forms.MdiLayout.ArrangeIcons)
    End Sub

    Public Sub MnuVBEditForm_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuVBEditForm.Click

        Dim Rs As ADODB.Recordset
        Dim Filename As String
        'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Filter_Renamed As String
        Dim Message As String

        On Error GoTo ErrorHandler

        Rs = Adodc1.Recordset.Clone

        'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        With SaveFileDialog1
            .DefaultExt = "Form1.frm"
            .Title = "Visual Basic 6 Edit Form"
            .FileName = "Form1.frm"
            Filter_Renamed = "Visual Basic 6 Form | *.frm|"
            Filter_Renamed = Filter_Renamed & "All types|*.*|"
            'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            .Filter = Filter_Renamed

        End With

        'UPGRADE_ISSUE: MSComDlg.CommonDialog control CommonDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        Filename = GetSaveFileDialog(SaveFileDialog1)

        If Len(Filename) > 0 Then
            Cursor = System.Windows.Forms.Cursors.WaitCursor
            Call ExportVBEditForm(Adodc1.ConnectionString, Rs, Filename)
            Cursor = System.Windows.Forms.Cursors.Default
        Else
            Exit Sub
        End If

        Message = "Exported as " & Filename
        MsgBox(Message, MsgBoxStyle.Information)

        Shell("notepad.exe " & Filename, AppWinStyle.NormalFocus)

        Exit Sub

ErrorHandler:
        Cursor = System.Windows.Forms.Cursors.Default
        LogFormError(Me.Name, "MnuExportVisualBasic6Form", Err.Description)
        MsgBox(Err.Description, MsgBoxStyle.Critical)

    End Sub

    Public Sub MnuVBLookupForm_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuVBLookupForm.Click

        Dim Rs As ADODB.Recordset
        Dim Filename As String
        'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Filter_Renamed As String
        Dim Message As String

        On Error GoTo ErrorHandler

        Rs = Adodc1.Recordset.Clone

        'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        With SaveFileDialog1

            .DefaultExt = "Form1.frm"
            .Title = "Visual Basic 6 Form"
            .FileName = "Form1.frm"
            Filter_Renamed = "Visual Basic 6 Form | *.frm|"
            Filter_Renamed = Filter_Renamed & "All types|*.*|"
            'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            .Filter = Filter_Renamed

        End With

        'UPGRADE_ISSUE: MSComDlg.CommonDialog control CommonDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        Filename = GetSaveFileDialog(SaveFileDialog1)

        If Len(Filename) > 0 Then
            Cursor = System.Windows.Forms.Cursors.WaitCursor
            Call ExportVBLookupForm(Adodc1.ConnectionString, Rs, Filename)
            Cursor = System.Windows.Forms.Cursors.Default
        Else
            Exit Sub
        End If

        Message = "Exported as " & Filename
        MsgBox(Message, MsgBoxStyle.Information)

        Shell("notepad.exe " & Filename, AppWinStyle.NormalFocus)

        Exit Sub

ErrorHandler:

        Cursor = System.Windows.Forms.Cursors.Default
        LogFormError(Me.Name, "MnuExportVisualBasic6Form", Err.Description)
        MsgBox(Err.Description, MsgBoxStyle.Critical)


    End Sub

    Public Sub MnuVBEditFormProject_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuVBEditFormProject.Click

        Dim tempPath As String
        Dim projectPath As String
        Dim ProjectName As String
        Dim Cn As New ADODB.Connection
        Dim SchemaRs As New ADODB.Recordset



        On Error GoTo ErrorHandler



        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Fso.FolderExists(tempPath) = False Then

            Exit Sub

        End If

        ProjectName = "Project 1"

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating project folders")
        Call CreateVBProjectFolders(tempPath, ProjectName)

        projectPath = tempPath & "\" & ProjectName

        projectPath = Replace(projectPath, "\\", "\")

        Cn.Open(Adodc1.ConnectionString)

        SchemaRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating project files")

        Call CreateVBProjectFile(projectPath & "\Source", ProjectName, SchemaRs, Cn)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating workspace")

        Call CreateVBWorkSpace(projectPath & "\Source", ProjectName, SchemaRs, Cn)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating modules")

        Call CreateVBModuleFile(projectPath & "\Source\Modules", Adodc1.ConnectionString)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating MDI form")

        Call CreateVBProjectMDIForm(projectPath & "\Source\Forms", ProjectName, SchemaRs, Cn)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating forms")

        Call CreateVBProjectAboutForm(projectPath & "\Source\Forms", ProjectName)

        Call CreateVBProjectEditForms(projectPath & "\Source\Forms", ProjectName, SchemaRs, Cn)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Project created")

        SchemaRs.Close()

        Cn.Close()

        'UPGRADE_NOTE: Object SchemaRs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        SchemaRs = Nothing
        'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Cn = Nothing

        MsgBox(ProjectName & " has been created successfully", MsgBoxStyle.Information)

        Call Shell("explorer.exe " & projectPath, AppWinStyle.NormalFocus)

        Exit Sub

ErrorHandler:

        MsgBox(Err.Number & ": " & Err.Description, MsgBoxStyle.Critical)


    End Sub

    Public Sub MnuVBLookupFormProject_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuVBLookupFormProject.Click

        Dim tempPath As String
        Dim projectPath As String
        Dim ProjectName As String
        Dim Cn As New ADODB.Connection
        Dim SchemaRs As New ADODB.Recordset



        On Error GoTo ErrorHandler

        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Fso.FolderExists(tempPath) = False Then

            Exit Sub

        End If

        ProjectName = "Project 1"

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating project folders ... ")

        Call CreateVBProjectFolders(tempPath, ProjectName)

        projectPath = tempPath & "\" & ProjectName

        projectPath = Replace(projectPath, "\\", "\")

        Cn.Open(Adodc1.ConnectionString)

        SchemaRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating project files ... ")

        Call CreateVBProjectFile(projectPath & "\Source", ProjectName, SchemaRs, Cn)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating workspace ... ")

        Call CreateVBWorkSpace(projectPath & "\Source", ProjectName, SchemaRs, Cn)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating modules ... ")

        Call CreateVBModuleFile(projectPath & "\Source\Modules", Adodc1.ConnectionString)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating MDI Form ... ")

        Call CreateVBProjectMDIForm(projectPath & "\Source\Forms", ProjectName, SchemaRs, Cn)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Creating Forms ... ")

        Call CreateVBProjectAboutForm(projectPath & "\Source\Forms", ProjectName)

        Call CreateVBProjectLookupForms(projectPath & "\Source\Forms", ProjectName, SchemaRs, Cn)

        System.Windows.Forms.Application.DoEvents()
        Call ShowStatus("Project created")

        SchemaRs.Close()

        Cn.Close()

        'UPGRADE_NOTE: Object SchemaRs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        SchemaRs = Nothing
        'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Cn = Nothing

        MsgBox(ProjectName & " has been created successfully", MsgBoxStyle.Information)

        Call Shell("explorer.exe " & projectPath, AppWinStyle.NormalFocus)


        Exit Sub

ErrorHandler:

        MsgBox(Err.Number & ": " & Err.Description, MsgBoxStyle.Critical)

    End Sub
	
	Public Sub MnuWebReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuWebReport.Click
		
		Dim Rs As ADODB.Recordset
        Dim Filepath As String
		'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Filter_Renamed As String
		Dim Message As String
		
		On Error GoTo ErrorHandler
		
		Rs = Adodc1.Recordset.Clone
		
        Filepath = FolderBrowserDialog(New FolderBrowserDialog)
		
        If Len(Filepath) > 0 Then
            Cursor = System.Windows.Forms.Cursors.WaitCursor
            Call BuildASPPrint(Adodc1.ConnectionString, Rs, Filepath)
            Cursor = System.Windows.Forms.Cursors.Default
        Else
            Exit Sub
        End If
		
        Message = "Exported as " & Filepath
		MsgBox(Message, MsgBoxStyle.Information)

		Exit Sub
		
ErrorHandler: 
		Cursor = System.Windows.Forms.Cursors.Default
		LogFormError(Me.Name, "MnuExportVisualBasic6Form", Err.Description)
		MsgBox(Err.Description, MsgBoxStyle.Critical)
		
	End Sub
	
	Public Sub ShowStatus(ByVal Message As String)
		
		On Error GoTo ErrorHandler
		
		System.Windows.Forms.Application.DoEvents()
        StatusStrip1.Text = Message
		
		Exit Sub
		
ErrorHandler: 
		
		Call LogFormError(Me.Name, "ShowStatus(" & Message & ")", Err.Number & ": " & Err.Description)
		
	End Sub


    Private Sub MnuASPExportForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuASPExportForm.Click

        Static Path As String
        Dim tempPath As String
        Dim Rs As New ADODB.Recordset


        On Error GoTo ErrorHandler

        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Len(Trim(tempPath)) = 0 Then
            Exit Sub
        End If

        If Fso.FolderExists(tempPath) = False Then
            MsgBox("Path not found", MsgBoxStyle.Critical)
        End If

        Path = tempPath

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        Rs = Adodc1.Recordset.Clone

        Call BuildASPExportForm(Adodc1.ConnectionString, Rs, Path)

        Cursor = System.Windows.Forms.Cursors.Default

        MsgBox("Completed", MsgBoxStyle.Information)

        Exit Sub

ErrorHandler:
        Cursor = System.Windows.Forms.Cursors.Default
        LogFormError(Me.Name, "MnuASPExportForm", Err.Description)
        MsgBox(Err.Description, MsgBoxStyle.Critical)


    End Sub

    Private Sub MnuASPProjectExportForm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MnuASPProjectExportForm.Click

        Dim tempPath As String
        Dim SchemaRs As New ADODB.Recordset

        Dim Cn As New ADODB.Connection


        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Fso.FolderExists(tempPath) = False Then

            Exit Sub

        End If

        Cn.Open(Adodc1.ConnectionString)

        SchemaRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        Call CreateASPForms(tempPath, SchemaRs, Adodc1.ConnectionString, "export")


        SchemaRs.Close()

        Cn.Close()

        'UPGRADE_NOTE: Object SchemaRs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        SchemaRs = Nothing
        'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Cn = Nothing

        MsgBox("Export forms have been created successfully", MsgBoxStyle.Information)
    End Sub

    Private Sub MnuASPProjectImportForm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MnuASPProjectImportForm.Click
        Dim tempPath As String
        Dim SchemaRs As New ADODB.Recordset

        Dim Cn As New ADODB.Connection

        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Fso.FolderExists(tempPath) = False Then

            Exit Sub

        End If

        Cn.Open(Adodc1.ConnectionString)

        SchemaRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        Call CreateASPForms(tempPath, SchemaRs, Adodc1.ConnectionString, "import")

        SchemaRs.Close()

        Cn.Close()

        'UPGRADE_NOTE: Object SchemaRs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        SchemaRs = Nothing
        'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Cn = Nothing

        MsgBox("Import forms have been created successfully", MsgBoxStyle.Information)

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub ImportFormToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportFormToolStripMenuItem.Click

        Static Path As String
        Dim tempPath As String
        Dim Rs As New ADODB.Recordset


        On Error GoTo ErrorHandler

        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Len(Trim(tempPath)) = 0 Then
            Exit Sub
        End If

        If Fso.FolderExists(tempPath) = False Then
            MsgBox("Path not found", MsgBoxStyle.Critical)
        End If

        Path = tempPath

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        Rs = Adodc1.Recordset.Clone

        Call BuildASPImportForm(Adodc1.ConnectionString, Rs, Path)

        Cursor = System.Windows.Forms.Cursors.Default

        MsgBox("Import form has been built successfully", MsgBoxStyle.Information)

        Exit Sub

ErrorHandler:
        Cursor = System.Windows.Forms.Cursors.Default
        LogFormError(Me.Name, "MnuASPImportForm", Err.Description)
        MsgBox(Err.Description, MsgBoxStyle.Critical)

    End Sub

    Private Sub CompactCodeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuCompactCode.Click

        Dim tempPath As String
        Dim SchemaRs As New ADODB.Recordset

        Dim Cn As New ADODB.Connection

        tempPath = FolderBrowserDialog(New FolderBrowserDialog)

        If Fso.FolderExists(tempPath) = False Then

            Exit Sub

        End If

    End Sub
End Class