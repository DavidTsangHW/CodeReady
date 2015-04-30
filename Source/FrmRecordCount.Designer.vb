<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FrmRecordCount
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		'This form is an MDI child.
		'This code simulates the VB6 
		' functionality of automatically
		' loading and showing an MDI
		' child's parent.
        'Me.MDIParent = CodeReady.MDIForm1
        'CodeReady.MDIForm1.Show
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
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
	Public WithEvents MnuOpenNew As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuSaveAs As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuSeparator1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents MnuRefresh As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuSeparator As System.Windows.Forms.ToolStripSeparator
	Public WithEvents MnuFont As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuGrid As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuPopUpMenu As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public WithEvents _ListView1_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListView1_ColumnHeader_2 As System.Windows.Forms.ColumnHeader
	Public WithEvents ListView1 As System.Windows.Forms.ListView
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmRecordCount))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.MnuPopUpMenu = New System.Windows.Forms.ToolStripMenuItem
		Me.MnuOpen = New System.Windows.Forms.ToolStripMenuItem
		Me.MnuOpenNew = New System.Windows.Forms.ToolStripMenuItem
		Me.MnuSaveAs = New System.Windows.Forms.ToolStripMenuItem
		Me.MnuSeparator1 = New System.Windows.Forms.ToolStripSeparator
		Me.MnuRefresh = New System.Windows.Forms.ToolStripMenuItem
		Me.MnuSeparator = New System.Windows.Forms.ToolStripSeparator
		Me.MnuFont = New System.Windows.Forms.ToolStripMenuItem
		Me.MnuGrid = New System.Windows.Forms.ToolStripMenuItem
		Me.CommonDialog1Font = New System.Windows.Forms.FontDialog
		Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog
		Me.ListView1 = New System.Windows.Forms.ListView
		Me._ListView1_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me._ListView1_ColumnHeader_2 = New System.Windows.Forms.ColumnHeader
		Me.MainMenu1.SuspendLayout()
		Me.ListView1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Record Count"
		Me.ClientSize = New System.Drawing.Size(234, 552)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
		Me.MinimizeBox = False
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmRecordCount"
		Me.MnuPopUpMenu.Name = "MnuPopUpMenu"
		Me.MnuPopUpMenu.Text = "PopUp Menu"
		Me.MnuPopUpMenu.Visible = False
		Me.MnuPopUpMenu.Checked = False
		Me.MnuPopUpMenu.Enabled = True
		Me.MnuOpen.Name = "MnuOpen"
		Me.MnuOpen.Text = "Open"
		Me.MnuOpen.Checked = False
		Me.MnuOpen.Enabled = True
		Me.MnuOpen.Visible = True
		Me.MnuOpenNew.Name = "MnuOpenNew"
		Me.MnuOpenNew.Text = "Open as new"
		Me.MnuOpenNew.Checked = False
		Me.MnuOpenNew.Enabled = True
		Me.MnuOpenNew.Visible = True
		Me.MnuSaveAs.Name = "MnuSaveAs"
		Me.MnuSaveAs.Text = "&Save as"
		Me.MnuSaveAs.Checked = False
		Me.MnuSaveAs.Enabled = True
		Me.MnuSaveAs.Visible = True
		Me.MnuSeparator1.Enabled = True
		Me.MnuSeparator1.Visible = True
		Me.MnuSeparator1.Name = "MnuSeparator1"
		Me.MnuRefresh.Name = "MnuRefresh"
		Me.MnuRefresh.Text = "&Refresh"
		Me.MnuRefresh.Checked = False
		Me.MnuRefresh.Enabled = True
		Me.MnuRefresh.Visible = True
		Me.MnuSeparator.Enabled = True
		Me.MnuSeparator.Visible = True
		Me.MnuSeparator.Name = "MnuSeparator"
		Me.MnuFont.Name = "MnuFont"
		Me.MnuFont.Text = "Font"
		Me.MnuFont.Checked = False
		Me.MnuFont.Enabled = True
		Me.MnuFont.Visible = True
		Me.MnuGrid.Name = "MnuGrid"
		Me.MnuGrid.Text = "Grid"
		Me.MnuGrid.Checked = False
		Me.MnuGrid.Enabled = True
		Me.MnuGrid.Visible = True
		Me.ListView1.Size = New System.Drawing.Size(217, 337)
		Me.ListView1.Location = New System.Drawing.Point(8, 32)
		Me.ListView1.TabIndex = 0
		Me.ListView1.View = System.Windows.Forms.View.Details
		Me.ListView1.LabelEdit = False
		Me.ListView1.MultiSelect = True
		Me.ListView1.LabelWrap = False
		Me.ListView1.HideSelection = False
		Me.ListView1.AllowColumnReorder = -1
		Me.ListView1.GridLines = True
		Me.ListView1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ListView1.BackColor = System.Drawing.SystemColors.Window
		Me.ListView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ListView1.Name = "ListView1"
		Me._ListView1_ColumnHeader_1.Text = "Table Name"
		Me._ListView1_ColumnHeader_1.Width = 170
		Me._ListView1_ColumnHeader_2.Text = "Record Count"
		Me._ListView1_ColumnHeader_2.Width = 170
		Me.Controls.Add(ListView1)
		Me.ListView1.Columns.Add(_ListView1_ColumnHeader_1)
		Me.ListView1.Columns.Add(_ListView1_ColumnHeader_2)
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.MnuPopUpMenu})
		MnuPopUpMenu.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.MnuOpen, Me.MnuOpenNew, Me.MnuSaveAs, Me.MnuSeparator1, Me.MnuRefresh, Me.MnuSeparator, Me.MnuFont, Me.MnuGrid})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ListView1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class