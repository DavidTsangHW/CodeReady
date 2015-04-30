Option Strict Off
Option Explicit On
Friend Class frmLicense
	Inherits System.Windows.Forms.Form
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		
		Me.Close()
		
	End Sub
	
	Private Sub cmdRegister_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRegister.Click
		
		Static Counter As Short
		
		If ValidateLicense(TxtRegisterName.Text, TxtKey.Text) = True Then
			
			Call SaveLicense(TxtRegisterName.Text, TxtKey.Text)
			
			RegisterName = TxtRegisterName.Text
			LicenseKey = TxtKey.Text
			IsLicensed = True
			
			MsgBox("Registration completed", MsgBoxStyle.Information)
			
			Me.Close()
			
		Else
			
			Counter = Counter + 1
			
			If Counter > 5 Then
				
				MsgBox("Invalid License Key" & vbCrLf & "Please contact Harvesoft", MsgBoxStyle.Critical)
				
				End
				
			Else
				
				MsgBox("Invalid License Key", MsgBoxStyle.Critical)
				
				TxtKey.Focus()
				
			End If
			
		End If
		
		
	End Sub
	
	Private Sub TxtKey_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtKey.Enter
		
		TxtKey.SelectionStart = 0
		TxtKey.SelectionLength = Len(TxtKey.Text)
		
	End Sub
	
	Private Sub TxtRegisterName_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtRegisterName.Enter
		
		TxtRegisterName.SelectionStart = 0
		TxtRegisterName.SelectionLength = Len(TxtRegisterName.Text)
		
	End Sub
End Class