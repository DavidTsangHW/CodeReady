<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FrmSQL
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
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
	Public WithEvents MnuExecute As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuSave As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents txtSQL As System.Windows.Forms.RichTextBox
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
	Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
	Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public CommonDialog1Print As System.Windows.Forms.PrintDialog
	Public WithEvents TxtLog As System.Windows.Forms.RichTextBox
	Public WithEvents CmdLog As System.Windows.Forms.Button
	Public WithEvents cmdExecute As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmSQL))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.MnuFile = New System.Windows.Forms.ToolStripMenuItem
		Me.MnuExecute = New System.Windows.Forms.ToolStripMenuItem
		Me.MnuSave = New System.Windows.Forms.ToolStripMenuItem
		Me.txtSQL = New System.Windows.Forms.RichTextBox
		Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog
		Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog
		Me.CommonDialog1Font = New System.Windows.Forms.FontDialog
		Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog
		Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog
		Me.TxtLog = New System.Windows.Forms.RichTextBox
		Me.CmdLog = New System.Windows.Forms.Button
		Me.cmdExecute = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Single Statement"
		Me.ClientSize = New System.Drawing.Size(341, 192)
		Me.Location = New System.Drawing.Point(3, 21)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmSQL"
		Me.MnuFile.Name = "MnuFile"
		Me.MnuFile.Text = "File"
		Me.MnuFile.Visible = False
		Me.MnuFile.Checked = False
		Me.MnuFile.Enabled = True
		Me.MnuExecute.Name = "MnuExecute"
		Me.MnuExecute.Text = "Execute"
		Me.MnuExecute.Checked = False
		Me.MnuExecute.Enabled = True
		Me.MnuExecute.Visible = True
		Me.MnuSave.Name = "MnuSave"
		Me.MnuSave.Text = "Save to file"
		Me.MnuSave.Checked = False
		Me.MnuSave.Enabled = True
		Me.MnuSave.Visible = True
		Me.txtSQL.Size = New System.Drawing.Size(337, 105)
		Me.txtSQL.Location = New System.Drawing.Point(0, 48)
		Me.txtSQL.TabIndex = 1
		Me.txtSQL.Enabled = True
		Me.txtSQL.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
		Me.txtSQL.RTF = resources.GetString("txtSQL.TextRTF")
		Me.txtSQL.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSQL.Name = "txtSQL"
		Me.TxtLog.Size = New System.Drawing.Size(337, 217)
		Me.TxtLog.Location = New System.Drawing.Point(0, 192)
		Me.TxtLog.TabIndex = 5
		Me.TxtLog.Enabled = True
		Me.TxtLog.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Both
		Me.TxtLog.RTF = resources.GetString("TxtLog.TextRTF")
		Me.TxtLog.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TxtLog.Name = "TxtLog"
		Me.CmdLog.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CmdLog.Text = "Log >>"
		Me.CmdLog.Size = New System.Drawing.Size(73, 25)
		Me.CmdLog.Location = New System.Drawing.Point(264, 160)
		Me.CmdLog.TabIndex = 4
		Me.CmdLog.BackColor = System.Drawing.SystemColors.Control
		Me.CmdLog.CausesValidation = True
		Me.CmdLog.Enabled = True
		Me.CmdLog.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CmdLog.Cursor = System.Windows.Forms.Cursors.Default
		Me.CmdLog.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CmdLog.TabStop = True
		Me.CmdLog.Name = "CmdLog"
		Me.cmdExecute.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdExecute.Text = "&Execute"
		Me.cmdExecute.Size = New System.Drawing.Size(73, 25)
		Me.cmdExecute.Location = New System.Drawing.Point(104, 160)
		Me.cmdExecute.TabIndex = 2
		Me.cmdExecute.BackColor = System.Drawing.SystemColors.Control
		Me.cmdExecute.CausesValidation = True
		Me.cmdExecute.Enabled = True
		Me.cmdExecute.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdExecute.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdExecute.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdExecute.TabStop = True
		Me.cmdExecute.Name = "cmdExecute"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "&Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(73, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(184, 160)
		Me.cmdCancel.TabIndex = 3
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.Label1.Text = "SQL statement:"
		Me.Label1.Size = New System.Drawing.Size(81, 17)
		Me.Label1.Location = New System.Drawing.Point(0, 32)
		Me.Label1.TabIndex = 0
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(txtSQL)
		Me.Controls.Add(TxtLog)
		Me.Controls.Add(CmdLog)
		Me.Controls.Add(cmdExecute)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(Label1)
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.MnuFile})
		MnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.MnuExecute, Me.MnuSave})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class