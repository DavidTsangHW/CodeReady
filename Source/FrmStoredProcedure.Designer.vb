<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FrmStoredProcedure
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
	Public WithEvents cmdExecute As System.Windows.Forms.Button
	Public WithEvents cmdBrowse As System.Windows.Forms.Button
	Public WithEvents TxtStoredProcedureFile As System.Windows.Forms.TextBox
	Public WithEvents TxtResult As System.Windows.Forms.RichTextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
	Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
	Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public CommonDialog1Print As System.Windows.Forms.PrintDialog
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmStoredProcedure))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdExecute = New System.Windows.Forms.Button
		Me.cmdBrowse = New System.Windows.Forms.Button
		Me.TxtStoredProcedureFile = New System.Windows.Forms.TextBox
		Me.TxtResult = New System.Windows.Forms.RichTextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog
		Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog
		Me.CommonDialog1Font = New System.Windows.Forms.FontDialog
		Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog
		Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Batch SQL"
		Me.ClientSize = New System.Drawing.Size(414, 347)
		Me.Location = New System.Drawing.Point(3, 22)
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
		Me.Name = "FrmStoredProcedure"
		Me.cmdExecute.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdExecute.Text = "&Execute"
		Me.cmdExecute.Size = New System.Drawing.Size(73, 25)
		Me.cmdExecute.Location = New System.Drawing.Point(248, 312)
		Me.cmdExecute.TabIndex = 3
		Me.cmdExecute.BackColor = System.Drawing.SystemColors.Control
		Me.cmdExecute.CausesValidation = True
		Me.cmdExecute.Enabled = True
		Me.cmdExecute.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdExecute.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdExecute.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdExecute.TabStop = True
		Me.cmdExecute.Name = "cmdExecute"
		Me.cmdBrowse.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdBrowse.Text = "..."
		Me.cmdBrowse.Size = New System.Drawing.Size(17, 17)
		Me.cmdBrowse.Location = New System.Drawing.Point(384, 16)
		Me.cmdBrowse.TabIndex = 1
		Me.cmdBrowse.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowse.CausesValidation = True
		Me.cmdBrowse.Enabled = True
		Me.cmdBrowse.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowse.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowse.TabStop = True
		Me.cmdBrowse.Name = "cmdBrowse"
		Me.TxtStoredProcedureFile.AutoSize = False
		Me.TxtStoredProcedureFile.Size = New System.Drawing.Size(337, 19)
		Me.TxtStoredProcedureFile.Location = New System.Drawing.Point(40, 16)
		Me.TxtStoredProcedureFile.TabIndex = 0
		Me.TxtStoredProcedureFile.AcceptsReturn = True
		Me.TxtStoredProcedureFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TxtStoredProcedureFile.BackColor = System.Drawing.SystemColors.Window
		Me.TxtStoredProcedureFile.CausesValidation = True
		Me.TxtStoredProcedureFile.Enabled = True
		Me.TxtStoredProcedureFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TxtStoredProcedureFile.HideSelection = True
		Me.TxtStoredProcedureFile.ReadOnly = False
		Me.TxtStoredProcedureFile.Maxlength = 0
		Me.TxtStoredProcedureFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TxtStoredProcedureFile.MultiLine = False
		Me.TxtStoredProcedureFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TxtStoredProcedureFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TxtStoredProcedureFile.TabStop = True
		Me.TxtStoredProcedureFile.Visible = True
		Me.TxtStoredProcedureFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TxtStoredProcedureFile.Name = "TxtStoredProcedureFile"
		Me.TxtResult.Size = New System.Drawing.Size(393, 249)
		Me.TxtResult.Location = New System.Drawing.Point(8, 56)
		Me.TxtResult.TabIndex = 2
		Me.TxtResult.Enabled = True
		Me.TxtResult.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Both
		Me.TxtResult.RTF = resources.GetString("TxtResult.TextRTF")
		Me.TxtResult.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TxtResult.Name = "TxtResult"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "&Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(73, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(328, 312)
		Me.cmdCancel.TabIndex = 4
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.Label2.Text = "Results"
		Me.Label2.Size = New System.Drawing.Size(81, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 40)
		Me.Label2.TabIndex = 6
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.Text = "File:"
		Me.Label1.Size = New System.Drawing.Size(25, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 16)
		Me.Label1.TabIndex = 5
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
		Me.Controls.Add(cmdExecute)
		Me.Controls.Add(cmdBrowse)
		Me.Controls.Add(TxtStoredProcedureFile)
		Me.Controls.Add(TxtResult)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class