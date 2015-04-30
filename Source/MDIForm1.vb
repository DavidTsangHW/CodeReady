Option Strict Off
Option Explicit On
Friend Class MDIForm1
	Inherits System.Windows.Forms.Form
	'Code Ready
	'9 August 2002
	
	Private Sub MnuArrangeIcon_Click()
		Me.LayoutMDI(4)
	End Sub
	
	'UPGRADE_WARNING: Form event MDIForm1.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub MDIForm1_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Static IsLoaded As Boolean

        'If IsLoaded = False Then

        '	MnuOpen_Click(MnuOpen, New System.EventArgs())

        'End If

        'IsLoaded = True
		
	End Sub
	
	Private Sub MDIForm1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		If IsLicensed = True Then
			
			MnuRegister.Visible = False
			
		End If
		
	End Sub
	
	Private Sub MDIForm1_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		Dim Response As String
		Dim Message As String
		
		Message = "Are you sure to exit program?"
		
		Response = CStr(MsgBox(Message, MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question))
		
		If Not Response = CStr(MsgBoxResult.Yes) Then
			Cancel = 1
		Else
			End
		End If
		
		eventArgs.Cancel = Cancel
	End Sub
	
	Public Sub MnuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuAbout.Click
		
		Dim FormAbout As New frmAbout
		
		With FormAbout
			.ShowDialog()
		End With
		
	End Sub
	
	Public Sub MnuCascade_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCascade.Click
		Me.LayoutMDI(System.Windows.Forms.MDILayout.Cascade)
	End Sub
	
	Public Sub MnuCodeReadyHomepage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCodeReadyHomepage.Click
		
		Call OpenBrowser("http://www24.brinkster.com/david6648668/projects/codeready")
		
	End Sub
	
	Public Sub MnuCompactDatabase_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuCompactDatabase.Click
		
		Dim Filepath As String
		'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Filter_Renamed As String
		
		Filter_Renamed = "Access Database|*.mdb"
		
		'UPGRADE_ISSUE: MSComDlg.CommonDialog control CommonDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        Filepath = OpenFileDialog(New OpenFileDialog)
		
		If Fso.FileExists(Filepath) = True Then
			
			If CompactDatabase(Filepath) = True Then
				
				MsgBox(Filepath & " compacted", MsgBoxStyle.Information)
				
			End If
			
		End If
		
	End Sub
	
	Public Sub MnuConnectionString_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuConnectionString.Click
		
		Dim FormGrid As New FrmGrid
		Dim FormTable As New FrmTable
		Dim Connstr As String
		Dim Message As String
		
		Message = "Enter Connection String"
		
		Connstr = InputBox(Message, "Connection String")
		
		'Open connection
		'If cancel was pressed, the connectionstring will become null
		If Len(Connstr) = 0 Then
			Exit Sub
		End If
		
		With FormTable
            .ConnectionString = Connstr
            FormGrid.MdiParent = Me
            .FormGrid = FormGrid
            .ShowDialog()
		End With
		
	End Sub
	
	'This sample program opens databases by using the ADO.
	'It can open number of database tables in a same time.
	
	Public Sub MnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuExit.Click
		Me.Close()
	End Sub
	
	Public Sub MnuOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuOpen.Click
		
		Dim FormGrid As New FrmGrid
		Dim FormTable As New FrmTable
		Dim Connstr As String
		Dim Message As String
		Dim Response As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object GetDataLinks. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Connstr = GetDataLinks
		
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
            FormGrid.MdiParent = Me
            .FormGrid = FormGrid
			.ShowDialog()
		End With
		
		
	End Sub
	
	
	Public Sub MnuRegister_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuRegister.Click
		
		Dim FormLicense As New frmLicense
		
		FormLicense.ShowDialog()
		
	End Sub
	
	Public Sub MnuTechnicalSupport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuTechnicalSupport.Click
		
		Call OpenBrowser("http://groups.msn.com/codeready")
		
	End Sub
	
	Public Sub MnuTileHorizontally_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuTileHorizontally.Click
		Me.LayoutMDI(System.Windows.Forms.MDILayout.TileHorizontal)
	End Sub
	
	Public Sub MnuTileVertically_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuTileVertically.Click
		Me.LayoutMDI(System.Windows.Forms.MDILayout.ArrangeIcons)
	End Sub
End Class