<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class MDIForm1
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
    Public WithEvents MnuOpen As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuConnectionString As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuSeparator3 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents MnuExit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuFile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuTileHorizontally As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuTileVertically As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuCascade As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuWindows As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuCompactDatabase As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuTools As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuRegister As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuTechnicalSupport As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuCodeReadyHomepage As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuSeperator1 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents MnuAbout As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
    Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
    Public CommonDialog1Font As System.Windows.Forms.FontDialog
    Public CommonDialog1Color As System.Windows.Forms.ColorDialog
    Public CommonDialog1Print As System.Windows.Forms.PrintDialog
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MDIForm1))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.MnuFile = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuOpen = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuConnectionString = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuSeparator3 = New System.Windows.Forms.ToolStripSeparator
        Me.MnuExit = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuWindows = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuTileHorizontally = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuTileVertically = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuCascade = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuTools = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuCompactDatabase = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuRegister = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuTechnicalSupport = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuCodeReadyHomepage = New System.Windows.Forms.ToolStripMenuItem
        Me.MnuSeperator1 = New System.Windows.Forms.ToolStripSeparator
        Me.MnuAbout = New System.Windows.Forms.ToolStripMenuItem
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog
        Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog
        Me.CommonDialog1Font = New System.Windows.Forms.FontDialog
        Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog
        Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuFile, Me.MnuWindows, Me.MnuTools, Me.MnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(312, 24)
        Me.MainMenu1.TabIndex = 1
        '
        'MnuFile
        '
        Me.MnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuOpen, Me.MnuConnectionString, Me.MnuSeparator3, Me.MnuExit})
        Me.MnuFile.Name = "MnuFile"
        Me.MnuFile.Size = New System.Drawing.Size(35, 20)
        Me.MnuFile.Text = "&File"
        '
        'MnuOpen
        '
        Me.MnuOpen.Name = "MnuOpen"
        Me.MnuOpen.Size = New System.Drawing.Size(188, 22)
        Me.MnuOpen.Text = "&Open ODBC"
        '
        'MnuConnectionString
        '
        Me.MnuConnectionString.Name = "MnuConnectionString"
        Me.MnuConnectionString.Size = New System.Drawing.Size(188, 22)
        Me.MnuConnectionString.Text = "&Enter Connection String"
        '
        'MnuSeparator3
        '
        Me.MnuSeparator3.Name = "MnuSeparator3"
        Me.MnuSeparator3.Size = New System.Drawing.Size(185, 6)
        '
        'MnuExit
        '
        Me.MnuExit.Name = "MnuExit"
        Me.MnuExit.Size = New System.Drawing.Size(188, 22)
        Me.MnuExit.Text = "E&xit"
        '
        'MnuWindows
        '
        Me.MnuWindows.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuTileHorizontally, Me.MnuTileVertically, Me.MnuCascade})
        Me.MnuWindows.MergeAction = System.Windows.Forms.MergeAction.Remove
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
        'MnuTools
        '
        Me.MnuTools.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuCompactDatabase})
        Me.MnuTools.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.MnuTools.Name = "MnuTools"
        Me.MnuTools.Size = New System.Drawing.Size(44, 20)
        Me.MnuTools.Text = "&Tools"
        '
        'MnuCompactDatabase
        '
        Me.MnuCompactDatabase.Name = "MnuCompactDatabase"
        Me.MnuCompactDatabase.Size = New System.Drawing.Size(232, 22)
        Me.MnuCompactDatabase.Text = "&Compact Database (Access only)"
        '
        'MnuHelp
        '
        Me.MnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuRegister, Me.MnuTechnicalSupport, Me.MnuCodeReadyHomepage, Me.MnuSeperator1, Me.MnuAbout})
        Me.MnuHelp.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.MnuHelp.Name = "MnuHelp"
        Me.MnuHelp.Size = New System.Drawing.Size(40, 20)
        Me.MnuHelp.Text = "&Help"
        '
        'MnuRegister
        '
        Me.MnuRegister.Name = "MnuRegister"
        Me.MnuRegister.Size = New System.Drawing.Size(187, 22)
        Me.MnuRegister.Text = "&Register"
        '
        'MnuTechnicalSupport
        '
        Me.MnuTechnicalSupport.Name = "MnuTechnicalSupport"
        Me.MnuTechnicalSupport.Size = New System.Drawing.Size(187, 22)
        Me.MnuTechnicalSupport.Text = "&Online Support"
        '
        'MnuCodeReadyHomepage
        '
        Me.MnuCodeReadyHomepage.Name = "MnuCodeReadyHomepage"
        Me.MnuCodeReadyHomepage.Size = New System.Drawing.Size(187, 22)
        Me.MnuCodeReadyHomepage.Text = "&Code Ready Homepage"
        '
        'MnuSeperator1
        '
        Me.MnuSeperator1.Name = "MnuSeperator1"
        Me.MnuSeperator1.Size = New System.Drawing.Size(184, 6)
        '
        'MnuAbout
        '
        Me.MnuAbout.Name = "MnuAbout"
        Me.MnuAbout.Size = New System.Drawing.Size(187, 22)
        Me.MnuAbout.Text = "&About Code Ready"
        '
        'MDIForm1
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ClientSize = New System.Drawing.Size(312, 237)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Location = New System.Drawing.Point(11, 57)
        Me.Name = "MDIForm1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Code Ready"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class