<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLicense
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
	Public WithEvents cmdRegister As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents TxtKey As System.Windows.Forms.TextBox
	Public WithEvents TxtRegisterName As System.Windows.Forms.TextBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLicense))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdRegister = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.TxtKey = New System.Windows.Forms.TextBox
		Me.TxtRegisterName = New System.Windows.Forms.TextBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "License"
		Me.ClientSize = New System.Drawing.Size(538, 236)
		Me.Location = New System.Drawing.Point(3, 29)
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
		Me.Name = "frmLicense"
		Me.cmdRegister.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdRegister
		Me.cmdRegister.Text = "&Register"
		Me.AcceptButton = Me.cmdRegister
		Me.cmdRegister.Size = New System.Drawing.Size(84, 23)
		Me.cmdRegister.Location = New System.Drawing.Point(336, 192)
		Me.cmdRegister.TabIndex = 3
		Me.cmdRegister.BackColor = System.Drawing.SystemColors.Control
		Me.cmdRegister.CausesValidation = True
		Me.cmdRegister.Enabled = True
		Me.cmdRegister.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdRegister.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdRegister.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdRegister.TabStop = True
		Me.cmdRegister.Name = "cmdRegister"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "&Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(84, 23)
		Me.cmdCancel.Location = New System.Drawing.Point(432, 192)
		Me.cmdCancel.TabIndex = 4
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.Frame1.Text = "License"
		Me.Frame1.Size = New System.Drawing.Size(505, 161)
		Me.Frame1.Location = New System.Drawing.Point(16, 16)
		Me.Frame1.TabIndex = 0
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.TxtKey.AutoSize = False
		Me.TxtKey.Size = New System.Drawing.Size(345, 19)
		Me.TxtKey.Location = New System.Drawing.Point(120, 80)
		Me.TxtKey.TabIndex = 2
		Me.TxtKey.AcceptsReturn = True
		Me.TxtKey.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TxtKey.BackColor = System.Drawing.SystemColors.Window
		Me.TxtKey.CausesValidation = True
		Me.TxtKey.Enabled = True
		Me.TxtKey.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TxtKey.HideSelection = True
		Me.TxtKey.ReadOnly = False
		Me.TxtKey.Maxlength = 0
		Me.TxtKey.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TxtKey.MultiLine = False
		Me.TxtKey.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TxtKey.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TxtKey.TabStop = True
		Me.TxtKey.Visible = True
		Me.TxtKey.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TxtKey.Name = "TxtKey"
		Me.TxtRegisterName.AutoSize = False
		Me.TxtRegisterName.Size = New System.Drawing.Size(345, 19)
		Me.TxtRegisterName.Location = New System.Drawing.Point(120, 32)
		Me.TxtRegisterName.TabIndex = 1
		Me.TxtRegisterName.AcceptsReturn = True
		Me.TxtRegisterName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TxtRegisterName.BackColor = System.Drawing.SystemColors.Window
		Me.TxtRegisterName.CausesValidation = True
		Me.TxtRegisterName.Enabled = True
		Me.TxtRegisterName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TxtRegisterName.HideSelection = True
		Me.TxtRegisterName.ReadOnly = False
		Me.TxtRegisterName.Maxlength = 0
		Me.TxtRegisterName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TxtRegisterName.MultiLine = False
		Me.TxtRegisterName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TxtRegisterName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TxtRegisterName.TabStop = True
		Me.TxtRegisterName.Visible = True
		Me.TxtRegisterName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TxtRegisterName.Name = "TxtRegisterName"
		Me.Label2.Text = "License Key:"
		Me.Label2.Size = New System.Drawing.Size(81, 17)
		Me.Label2.Location = New System.Drawing.Point(16, 80)
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
		Me.Label1.Text = "Register Name:"
		Me.Label1.Size = New System.Drawing.Size(97, 17)
		Me.Label1.Location = New System.Drawing.Point(16, 32)
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
		Me.Controls.Add(cmdRegister)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(Frame1)
		Me.Frame1.Controls.Add(TxtKey)
		Me.Frame1.Controls.Add(TxtRegisterName)
		Me.Frame1.Controls.Add(Label2)
		Me.Frame1.Controls.Add(Label1)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class