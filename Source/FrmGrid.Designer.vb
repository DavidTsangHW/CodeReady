<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FrmGrid
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()

	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			VB6_RemoveADODataBinding()
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents MnuOpen As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuEnterConnectionString As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuSaveConnectionString As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuConnectionString As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuExportTextFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuExport As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuASPProjectBlankForm As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuASPProjectEditForm As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuASPProjectGrid As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuASPProjectReportList As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuASPProject As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuASPBlankForm As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuASPEditForm As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuASPGrid As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuWebReport As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuFormForm As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuASP As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuVBEditFormProject As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuVBLookupFormProject As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuVBProject As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuVBEditForm As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuVBLookupForm As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuVBForms As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuExportVisualBasic6Form As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuCodeGen As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuSeperator1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents MnuExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuFind As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuFindNext As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuSinglStatement As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuMultipleStatement As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuSQL As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuTable As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuEdit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuFonts As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuApperance As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuFormat As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuRecordCount As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuMonitorAddRemove As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuMonitor As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuSchema As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuTools As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuTileHorizontally As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuTileVertically As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuCascade As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuWindows As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuRegister As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuTechinicalSupport As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuCodeReadyHomepage As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuSeparator1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents MnuAbout As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public WithEvents Adodc1 As VB6.ADODC
	Public WithEvents DataGrid1 As AxMSDataGridLib.AxDataGrid
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmGrid))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.MnuFile = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuOpen = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuConnectionString = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuEnterConnectionString = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuSaveConnectionString = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuExport = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuExportTextFile = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuCodeGen = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASP = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASPProject = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASPProjectBlankForm = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASPProjectEditForm = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASPProjectGrid = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASPProjectReportList = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASPProjectImportForm = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASPProjectExportForm = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuFormForm = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASPBlankForm = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASPEditForm = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASPGrid = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuWebReport = New System.Windows.Forms.ToolStripMenuItem
        Me.ImportFormToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuASPExportForm = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuExportVisualBasic6Form = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuVBProject = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuVBEditFormProject = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuVBLookupFormProject = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuVBForms = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuVBEditForm = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuVBLookupForm = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuSeperator1 = New System.Windows.Forms.ToolStripSeparator
        Me.MnuExit = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuEdit = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuFind = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuFindNext = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuSQL = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuSinglStatement = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuMultipleStatement = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuTable = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuFormat = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuFonts = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuApperance = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuTools = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuRecordCount = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuMonitor = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuMonitorAddRemove = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuSchema = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.MnuCompactCode = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuWindows = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuTileHorizontally = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuTileVertically = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuCascade = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuRegister = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuTechinicalSupport = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuCodeReadyHomepage = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.MnuAbout = New System.Windows.Forms.ToolStripMenuItem
        Me.CommonDialog1Font = New System.Windows.Forms.FontDialog
        Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog
        Me.Adodc1 = New Microsoft.VisualBasic.Compatibility.VB6.ADODC
        Me.DataGrid1 = New AxMSDataGridLib.AxDataGrid
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.FontDialog1 = New System.Windows.Forms.FontDialog
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.MainMenu1.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuFile, Me.MnuEdit, Me.MnuFormat, Me.MnuTools, Me.MnuWindows, Me.MnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(578, 24)
        Me.MainMenu1.TabIndex = 3
        '
        'MnuFile
        '
        Me.MnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuOpen, Me.MnuConnectionString, Me.MnuExport, Me.MnuCodeGen, Me.MnuSeperator1, Me.MnuExit})
        Me.MnuFile.MergeAction = System.Windows.Forms.MergeAction.Replace
        Me.MnuFile.Name = "MnuFile"
        Me.MnuFile.Size = New System.Drawing.Size(35, 20)
        Me.MnuFile.Text = "&File"
        '
        'MnuOpen
        '
        Me.MnuOpen.Name = "MnuOpen"
        Me.MnuOpen.Size = New System.Drawing.Size(159, 22)
        Me.MnuOpen.Text = "&Open ODBC"
        '
        'MnuConnectionString
        '
        Me.MnuConnectionString.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuEnterConnectionString, Me.MnuSaveConnectionString})
        Me.MnuConnectionString.Name = "MnuConnectionString"
        Me.MnuConnectionString.Size = New System.Drawing.Size(159, 22)
        Me.MnuConnectionString.Text = "&Connection String"
        '
        'MnuEnterConnectionString
        '
        Me.MnuEnterConnectionString.Name = "MnuEnterConnectionString"
        Me.MnuEnterConnectionString.Size = New System.Drawing.Size(100, 22)
        Me.MnuEnterConnectionString.Text = "&Enter"
        '
        'MnuSaveConnectionString
        '
        Me.MnuSaveConnectionString.Name = "MnuSaveConnectionString"
        Me.MnuSaveConnectionString.Size = New System.Drawing.Size(100, 22)
        Me.MnuSaveConnectionString.Text = "&Save"
        '
        'MnuExport
        '
        Me.MnuExport.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuExportTextFile})
        Me.MnuExport.Name = "MnuExport"
        Me.MnuExport.Size = New System.Drawing.Size(159, 22)
        Me.MnuExport.Text = "&Export"
        '
        'MnuExportTextFile
        '
        Me.MnuExportTextFile.Name = "MnuExportTextFile"
        Me.MnuExportTextFile.Size = New System.Drawing.Size(115, 22)
        Me.MnuExportTextFile.Text = "&Text File"
        '
        'MnuCodeGen
        '
        Me.MnuCodeGen.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuASP, Me.MnuExportVisualBasic6Form})
        Me.MnuCodeGen.Name = "MnuCodeGen"
        Me.MnuCodeGen.Size = New System.Drawing.Size(159, 22)
        Me.MnuCodeGen.Text = "&Code Gen"
        '
        'MnuASP
        '
        Me.MnuASP.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuASPProject, Me.MnuFormForm})
        Me.MnuASP.Name = "MnuASP"
        Me.MnuASP.Size = New System.Drawing.Size(152, 22)
        Me.MnuASP.Text = "&ASP"
        '
        'MnuASPProject
        '
        Me.MnuASPProject.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuASPProjectBlankForm, Me.MnuASPProjectEditForm, Me.MnuASPProjectGrid, Me.MnuASPProjectReportList, Me.MnuASPProjectImportForm, Me.MnuASPProjectExportForm})
        Me.MnuASPProject.Name = "MnuASPProject"
        Me.MnuASPProject.Size = New System.Drawing.Size(152, 22)
        Me.MnuASPProject.Text = "&Project"
        '
        'MnuASPProjectBlankForm
        '
        Me.MnuASPProjectBlankForm.Name = "MnuASPProjectBlankForm"
        Me.MnuASPProjectBlankForm.Size = New System.Drawing.Size(152, 22)
        Me.MnuASPProjectBlankForm.Text = "&Blank Form"
        '
        'MnuASPProjectEditForm
        '
        Me.MnuASPProjectEditForm.Name = "MnuASPProjectEditForm"
        Me.MnuASPProjectEditForm.Size = New System.Drawing.Size(152, 22)
        Me.MnuASPProjectEditForm.Text = "&Edit Form"
        '
        'MnuASPProjectGrid
        '
        Me.MnuASPProjectGrid.Name = "MnuASPProjectGrid"
        Me.MnuASPProjectGrid.Size = New System.Drawing.Size(152, 22)
        Me.MnuASPProjectGrid.Text = "&Grid"
        '
        'MnuASPProjectReportList
        '
        Me.MnuASPProjectReportList.Name = "MnuASPProjectReportList"
        Me.MnuASPProjectReportList.Size = New System.Drawing.Size(152, 22)
        Me.MnuASPProjectReportList.Text = "&Report List"
        '
        'MnuASPProjectImportForm
        '
        Me.MnuASPProjectImportForm.Name = "MnuASPProjectImportForm"
        Me.MnuASPProjectImportForm.Size = New System.Drawing.Size(152, 22)
        Me.MnuASPProjectImportForm.Text = "&Import Form"
        '
        'MnuASPProjectExportForm
        '
        Me.MnuASPProjectExportForm.Name = "MnuASPProjectExportForm"
        Me.MnuASPProjectExportForm.Size = New System.Drawing.Size(152, 22)
        Me.MnuASPProjectExportForm.Text = "E&xport Form"
        '
        'MnuFormForm
        '
        Me.MnuFormForm.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuASPBlankForm, Me.MnuASPEditForm, Me.MnuASPGrid, Me.MnuWebReport, Me.ImportFormToolStripMenuItem, Me.MnuASPExportForm})
        Me.MnuFormForm.Name = "MnuFormForm"
        Me.MnuFormForm.Size = New System.Drawing.Size(152, 22)
        Me.MnuFormForm.Text = "&Form"
        '
        'MnuASPBlankForm
        '
        Me.MnuASPBlankForm.Name = "MnuASPBlankForm"
        Me.MnuASPBlankForm.Size = New System.Drawing.Size(133, 22)
        Me.MnuASPBlankForm.Text = "&Blank Form"
        '
        'MnuASPEditForm
        '
        Me.MnuASPEditForm.Name = "MnuASPEditForm"
        Me.MnuASPEditForm.Size = New System.Drawing.Size(133, 22)
        Me.MnuASPEditForm.Text = "&Edit Form"
        '
        'MnuASPGrid
        '
        Me.MnuASPGrid.Name = "MnuASPGrid"
        Me.MnuASPGrid.Size = New System.Drawing.Size(133, 22)
        Me.MnuASPGrid.Text = "&Grid"
        '
        'MnuWebReport
        '
        Me.MnuWebReport.Name = "MnuWebReport"
        Me.MnuWebReport.Size = New System.Drawing.Size(133, 22)
        Me.MnuWebReport.Text = "&Report List"
        '
        'ImportFormToolStripMenuItem
        '
        Me.ImportFormToolStripMenuItem.Name = "ImportFormToolStripMenuItem"
        Me.ImportFormToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.ImportFormToolStripMenuItem.Text = "&Import Form"
        '
        'MnuASPExportForm
        '
        Me.MnuASPExportForm.Name = "MnuASPExportForm"
        Me.MnuASPExportForm.Size = New System.Drawing.Size(133, 22)
        Me.MnuASPExportForm.Text = "E&xport Form"
        '
        'MnuExportVisualBasic6Form
        '
        Me.MnuExportVisualBasic6Form.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuVBProject, Me.MnuVBForms})
        Me.MnuExportVisualBasic6Form.Name = "MnuExportVisualBasic6Form"
        Me.MnuExportVisualBasic6Form.Size = New System.Drawing.Size(152, 22)
        Me.MnuExportVisualBasic6Form.Text = "&Visual Basic 6"
        '
        'MnuVBProject
        '
        Me.MnuVBProject.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuVBEditFormProject, Me.MnuVBLookupFormProject})
        Me.MnuVBProject.Name = "MnuVBProject"
        Me.MnuVBProject.Size = New System.Drawing.Size(108, 22)
        Me.MnuVBProject.Text = "&Project"
        '
        'MnuVBEditFormProject
        '
        Me.MnuVBEditFormProject.Name = "MnuVBEditFormProject"
        Me.MnuVBEditFormProject.Size = New System.Drawing.Size(135, 22)
        Me.MnuVBEditFormProject.Text = "&Edit Form"
        '
        'MnuVBLookupFormProject
        '
        Me.MnuVBLookupFormProject.Name = "MnuVBLookupFormProject"
        Me.MnuVBLookupFormProject.Size = New System.Drawing.Size(135, 22)
        Me.MnuVBLookupFormProject.Text = "&Lookup Form"
        '
        'MnuVBForms
        '
        Me.MnuVBForms.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuVBEditForm, Me.MnuVBLookupForm})
        Me.MnuVBForms.Name = "MnuVBForms"
        Me.MnuVBForms.Size = New System.Drawing.Size(108, 22)
        Me.MnuVBForms.Text = "&Form"
        '
        'MnuVBEditForm
        '
        Me.MnuVBEditForm.Name = "MnuVBEditForm"
        Me.MnuVBEditForm.Size = New System.Drawing.Size(135, 22)
        Me.MnuVBEditForm.Text = "&Edit Form"
        '
        'MnuVBLookupForm
        '
        Me.MnuVBLookupForm.Name = "MnuVBLookupForm"
        Me.MnuVBLookupForm.Size = New System.Drawing.Size(135, 22)
        Me.MnuVBLookupForm.Text = "&Lookup Form"
        '
        'MnuSeperator1
        '
        Me.MnuSeperator1.Name = "MnuSeperator1"
        Me.MnuSeperator1.Size = New System.Drawing.Size(156, 6)
        '
        'MnuExit
        '
        Me.MnuExit.Name = "MnuExit"
        Me.MnuExit.Size = New System.Drawing.Size(159, 22)
        Me.MnuExit.Text = "E&xit"
        '
        'MnuEdit
        '
        Me.MnuEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuFind, Me.MnuFindNext, Me.MnuSQL, Me.MnuTable})
        Me.MnuEdit.Name = "MnuEdit"
        Me.MnuEdit.Size = New System.Drawing.Size(37, 20)
        Me.MnuEdit.Text = "&Edit"
        '
        'MnuFind
        '
        Me.MnuFind.Name = "MnuFind"
        Me.MnuFind.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
        Me.MnuFind.Size = New System.Drawing.Size(159, 22)
        Me.MnuFind.Text = "&Find"
        '
        'MnuFindNext
        '
        Me.MnuFindNext.Name = "MnuFindNext"
        Me.MnuFindNext.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.MnuFindNext.Size = New System.Drawing.Size(159, 22)
        Me.MnuFindNext.Text = "Find &Next"
        '
        'MnuSQL
        '
        Me.MnuSQL.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuSinglStatement, Me.MnuMultipleStatement})
        Me.MnuSQL.Name = "MnuSQL"
        Me.MnuSQL.Size = New System.Drawing.Size(159, 22)
        Me.MnuSQL.Text = "&SQL"
        '
        'MnuSinglStatement
        '
        Me.MnuSinglStatement.Name = "MnuSinglStatement"
        Me.MnuSinglStatement.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.MnuSinglStatement.Size = New System.Drawing.Size(140, 22)
        Me.MnuSinglStatement.Text = "&Single"
        '
        'MnuMultipleStatement
        '
        Me.MnuMultipleStatement.Name = "MnuMultipleStatement"
        Me.MnuMultipleStatement.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
        Me.MnuMultipleStatement.Size = New System.Drawing.Size(140, 22)
        Me.MnuMultipleStatement.Text = "&Batch"
        '
        'MnuTable
        '
        Me.MnuTable.Name = "MnuTable"
        Me.MnuTable.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.T), System.Windows.Forms.Keys)
        Me.MnuTable.Size = New System.Drawing.Size(159, 22)
        Me.MnuTable.Text = "&Table"
        '
        'MnuFormat
        '
        Me.MnuFormat.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuFonts, Me.MnuApperance})
        Me.MnuFormat.Name = "MnuFormat"
        Me.MnuFormat.Size = New System.Drawing.Size(53, 20)
        Me.MnuFormat.Text = "F&ormat"
        '
        'MnuFonts
        '
        Me.MnuFonts.Name = "MnuFonts"
        Me.MnuFonts.Size = New System.Drawing.Size(126, 22)
        Me.MnuFonts.Text = "&Fonts"
        '
        'MnuApperance
        '
        Me.MnuApperance.Name = "MnuApperance"
        Me.MnuApperance.Size = New System.Drawing.Size(126, 22)
        Me.MnuApperance.Text = "&Apperance"
        '
        'MnuTools
        '
        Me.MnuTools.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuRecordCount, Me.MnuMonitor, Me.MnuSchema, Me.ToolStripSeparator1, Me.MnuCompactCode})
        Me.MnuTools.Name = "MnuTools"
        Me.MnuTools.Size = New System.Drawing.Size(44, 20)
        Me.MnuTools.Text = "&Tools"
        '
        'MnuRecordCount
        '
        Me.MnuRecordCount.Name = "MnuRecordCount"
        Me.MnuRecordCount.Size = New System.Drawing.Size(159, 22)
        Me.MnuRecordCount.Text = "&Record Count"
        '
        'MnuMonitor
        '
        Me.MnuMonitor.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuMonitorAddRemove})
        Me.MnuMonitor.Name = "MnuMonitor"
        Me.MnuMonitor.Size = New System.Drawing.Size(159, 22)
        Me.MnuMonitor.Text = "&Database Monitor"
        '
        'MnuMonitorAddRemove
        '
        Me.MnuMonitorAddRemove.Name = "MnuMonitorAddRemove"
        Me.MnuMonitorAddRemove.Size = New System.Drawing.Size(172, 22)
        Me.MnuMonitorAddRemove.Text = "&Add Remove Record"
        '
        'MnuSchema
        '
        Me.MnuSchema.Name = "MnuSchema"
        Me.MnuSchema.Size = New System.Drawing.Size(159, 22)
        Me.MnuSchema.Text = "&Create Schema"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(156, 6)
        '
        'MnuCompactCode
        '
        Me.MnuCompactCode.Name = "MnuCompactCode"
        Me.MnuCompactCode.Size = New System.Drawing.Size(159, 22)
        Me.MnuCompactCode.Text = "Compact code"
        '
        'MnuWindows
        '
        Me.MnuWindows.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuTileHorizontally, Me.MnuTileVertically, Me.MnuCascade})
        Me.MnuWindows.Name = "MnuWindows"
        Me.MnuWindows.Size = New System.Drawing.Size(62, 20)
        Me.MnuWindows.Text = "&Windows"
        '
        'MnuTileHorizontally
        '
        Me.MnuTileHorizontally.Name = "MnuTileHorizontally"
        Me.MnuTileHorizontally.Size = New System.Drawing.Size(149, 22)
        Me.MnuTileHorizontally.Text = "Tile &Horizontally"
        '
        'MnuTileVertically
        '
        Me.MnuTileVertically.Name = "MnuTileVertically"
        Me.MnuTileVertically.Size = New System.Drawing.Size(149, 22)
        Me.MnuTileVertically.Text = "Tile &Vertically"
        '
        'MnuCascade
        '
        Me.MnuCascade.Name = "MnuCascade"
        Me.MnuCascade.Size = New System.Drawing.Size(149, 22)
        Me.MnuCascade.Text = "&Cascade"
        '
        'MnuHelp
        '
        Me.MnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuRegister, Me.MnuTechinicalSupport, Me.MnuCodeReadyHomepage, Me.MnuSeparator1, Me.MnuAbout})
        Me.MnuHelp.Name = "MnuHelp"
        Me.MnuHelp.Size = New System.Drawing.Size(40, 20)
        Me.MnuHelp.Text = "&Help"
        '
        'MnuRegister
        '
        Me.MnuRegister.Name = "MnuRegister"
        Me.MnuRegister.Size = New System.Drawing.Size(184, 22)
        Me.MnuRegister.Text = "&Register"
        '
        'MnuTechinicalSupport
        '
        Me.MnuTechinicalSupport.Name = "MnuTechinicalSupport"
        Me.MnuTechinicalSupport.Size = New System.Drawing.Size(184, 22)
        Me.MnuTechinicalSupport.Text = "&Online Support"
        '
        'MnuCodeReadyHomepage
        '
        Me.MnuCodeReadyHomepage.Name = "MnuCodeReadyHomepage"
        Me.MnuCodeReadyHomepage.Size = New System.Drawing.Size(184, 22)
        Me.MnuCodeReadyHomepage.Text = "&CodeReady Homepage"
        '
        'MnuSeparator1
        '
        Me.MnuSeparator1.Name = "MnuSeparator1"
        Me.MnuSeparator1.Size = New System.Drawing.Size(181, 6)
        '
        'MnuAbout
        '
        Me.MnuAbout.Name = "MnuAbout"
        Me.MnuAbout.Size = New System.Drawing.Size(184, 22)
        Me.MnuAbout.Text = "&About Code Ready"
        '
        'Adodc1
        '
        Me.Adodc1.BackColor = System.Drawing.SystemColors.Window
        Me.Adodc1.BOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.BOFActionEnum.adDoMoveFirst
        Me.Adodc1.CommandTimeout = 0
        Me.Adodc1.CommandType = ADODB.CommandTypeEnum.adCmdUnknown
        Me.Adodc1.ConnectionString = Nothing
        Me.Adodc1.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Me.Adodc1.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
        Me.Adodc1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Adodc1.EOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.EOFActionEnum.adDoMoveLast
        Me.Adodc1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Adodc1.Location = New System.Drawing.Point(0, 399)
        Me.Adodc1.LockType = ADODB.LockTypeEnum.adLockOptimistic
        Me.Adodc1.Mode = ADODB.ConnectModeEnum.adModeUnknown
        Me.Adodc1.Name = "Adodc1"
        Me.Adodc1.Orientation = Microsoft.VisualBasic.Compatibility.VB6.ADODC.OrientationEnum.adHorizontal
        Me.Adodc1.Size = New System.Drawing.Size(578, 25)
        Me.Adodc1.TabIndex = 2
        Me.Adodc1.Text = "Adodc1"
        '
        'DataGrid1
        '
        Me.DataGrid1.DataSource = Nothing
        Me.DataGrid1.Dock = System.Windows.Forms.DockStyle.Top
        Me.DataGrid1.Location = New System.Drawing.Point(0, 24)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.OcxState = CType(resources.GetObject("DataGrid1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.DataGrid1.Size = New System.Drawing.Size(578, 400)
        Me.DataGrid1.TabIndex = 0
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 377)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(578, 22)
        Me.StatusStrip1.TabIndex = 4
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'FrmGrid
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(578, 424)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Adodc1)
        Me.Controls.Add(Me.DataGrid1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(4, 42)
        Me.Name = "FrmGrid"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "CR/2"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support"
    Public Sub VB6_AddADODataBinding()
        DataGrid1.DataSource = CType(Adodc1, msdatasrc.DataSource)
    End Sub
    Public Sub VB6_RemoveADODataBinding()
        DataGrid1.DataSource = Nothing
    End Sub
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FontDialog1 As System.Windows.Forms.FontDialog
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents MnuASPProjectImportForm As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuASPProjectExportForm As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportFormToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuASPExportForm As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents MnuCompactCode As System.Windows.Forms.ToolStripMenuItem
#End Region
End Class