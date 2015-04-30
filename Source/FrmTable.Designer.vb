<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FrmTable
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
	Public WithEvents ListTable As System.Windows.Forms.ListBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOpen As System.Windows.Forms.Button
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmTable))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ListTable = New System.Windows.Forms.ListBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOpen = New System.Windows.Forms.Button
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Tables"
		Me.ClientSize = New System.Drawing.Size(205, 389)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.MaximizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmTable"
		Me.ListTable.Size = New System.Drawing.Size(201, 332)
		Me.ListTable.Location = New System.Drawing.Point(0, 8)
		Me.ListTable.TabIndex = 2
		Me.ListTable.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ListTable.BackColor = System.Drawing.SystemColors.Window
		Me.ListTable.CausesValidation = True
		Me.ListTable.Enabled = True
		Me.ListTable.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ListTable.IntegralHeight = True
		Me.ListTable.Cursor = System.Windows.Forms.Cursors.Default
		Me.ListTable.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.ListTable.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ListTable.Sorted = False
		Me.ListTable.TabStop = True
		Me.ListTable.Visible = True
		Me.ListTable.MultiColumn = False
		Me.ListTable.Name = "ListTable"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "&Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(136, 352)
		Me.cmdCancel.TabIndex = 1
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdOpen.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOpen.Text = "&Open"
		Me.cmdOpen.Size = New System.Drawing.Size(65, 25)
		Me.cmdOpen.Location = New System.Drawing.Point(64, 352)
		Me.cmdOpen.TabIndex = 0
		Me.cmdOpen.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOpen.CausesValidation = True
		Me.cmdOpen.Enabled = True
		Me.cmdOpen.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOpen.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOpen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOpen.TabStop = True
		Me.cmdOpen.Name = "cmdOpen"
		Me.Controls.Add(ListTable)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOpen)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class