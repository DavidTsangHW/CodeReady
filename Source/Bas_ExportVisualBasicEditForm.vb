Option Strict Off
Option Explicit On
Module Bas_ExportVisualBasicEditForm
	Private Const ModuleName As String = "Bas_ExportVisualBasicEditForm"
	
	Public Sub ExportVBEditForm(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal VBForm As String)
		
		Dim WriteTs As Scripting.TextStream
		
		Dim Fieldidx As Short
		Dim lblfield As String
		Dim txtfield As String
		
		Dim txtrs As String
		
		Dim FormHeader As String
		Dim FormContent As String
		Dim Line As String
		
		WriteTs = FSO.OpenTextFile(VBForm, Scripting.IOMode.ForWriting, True)
		
		WriteTs.WriteLine("VERSION 5.00")
		WriteTs.WriteLine("Begin VB.Form Form1 ")
		WriteTs.WriteLine("   Caption         =   ""Form1""")
		WriteTs.WriteLine("   ClientHeight    =   7920")
		WriteTs.WriteLine("   ClientLeft      =   60")
		WriteTs.WriteLine("   ClientTop       =   345")
		WriteTs.WriteLine("   ClientWidth     =   6930")
		WriteTs.WriteLine("   LinkTopic       =   ""Form1""")
		WriteTs.WriteLine("   ScaleHeight     =   7920")
		WriteTs.WriteLine("   ScaleWidth      =   6930")
		WriteTs.WriteLine("   StartUpPosition =   2  'CenterScreen")
		WriteTs.WriteLine("   Begin VB.CommandButton cmdFindNext ")
		WriteTs.WriteLine("      Caption         =   ""Find &Next""")
		WriteTs.WriteLine("      Height          =   375")
		WriteTs.WriteLine("      Left            =   1440")
		WriteTs.WriteLine("      TabIndex        =   65")
		WriteTs.WriteLine("      Top             =   6480")
		WriteTs.WriteLine("      Width           =   1095")
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("   Begin VB.CommandButton cmdFind ")
		WriteTs.WriteLine("      Caption         =   ""&Find""")
		WriteTs.WriteLine("      Height          =   375")
		WriteTs.WriteLine("      Left            =   120")
		WriteTs.WriteLine("      TabIndex        =   64")
		WriteTs.WriteLine("      Top             =   6480")
		WriteTs.WriteLine("      Width           =   1095")
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("   Begin VB.VScrollBar VScroll1 ")
		WriteTs.WriteLine("      Height          =   5055")
		WriteTs.WriteLine("      Left            =   6600")
		WriteTs.WriteLine("      TabIndex        =   63")
		WriteTs.WriteLine("      Top             =   240")
		WriteTs.WriteLine("      Width           =   255")
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("   Begin VB.CommandButton cmdDelete ")
		WriteTs.WriteLine("      Caption         =   ""&Delete""")
		WriteTs.WriteLine("      Height          =   375")
		WriteTs.WriteLine("      Left            =   4080")
		WriteTs.WriteLine("      TabIndex        =   74")
		WriteTs.WriteLine("      Top             =   6000")
		WriteTs.WriteLine("      Width           =   1095")
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("   Begin VB.CommandButton cmdCancelUpdate ")
		WriteTs.WriteLine("      Caption         =   ""&Cancel""")
		WriteTs.WriteLine("      Height          =   375")
		WriteTs.WriteLine("      Left            =   1440")
		WriteTs.WriteLine("      TabIndex        =   73")
		WriteTs.WriteLine("      Top             =   6000")
		WriteTs.WriteLine("      Width           =   1095")
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("   Begin VB.CommandButton cmdUpdate ")
		WriteTs.WriteLine("      Caption         =   ""Update""")
		WriteTs.WriteLine("      Height          =   375")
		WriteTs.WriteLine("      Left            =   2760")
		WriteTs.WriteLine("      TabIndex        =   72")
		WriteTs.WriteLine("      Top             =   6000")
		WriteTs.WriteLine("      Width           =   1095")
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("   Begin VB.CommandButton cmdAdd ")
		WriteTs.WriteLine("      Caption         =   ""&Add New""")
		WriteTs.WriteLine("      Height          =   375")
		WriteTs.WriteLine("      Left            =   120")
		WriteTs.WriteLine("      TabIndex        =   71")
		WriteTs.WriteLine("      Top             =   6000")
		WriteTs.WriteLine("      Width           =   1095")
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("   Begin VB.CommandButton cmdTop ")
		WriteTs.WriteLine("      Caption         =   ""&Top""")
		WriteTs.WriteLine("      Height          =   375")
		WriteTs.WriteLine("      Left            =   120")
		WriteTs.WriteLine("      TabIndex        =   70")
		WriteTs.WriteLine("      Top             =   5520")
		WriteTs.WriteLine("      Width           =   1095")
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("   Begin VB.CommandButton cmdend ")
		WriteTs.WriteLine("      Caption         =   ""&End""")
		WriteTs.WriteLine("      Height          =   375")
		WriteTs.WriteLine("      Left            =   4080")
		WriteTs.WriteLine("      TabIndex        =   69")
		WriteTs.WriteLine("      Top             =   5520")
		WriteTs.WriteLine("      Width           =   1095")
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("   Begin VB.CommandButton cmdPrevious ")
		WriteTs.WriteLine("      Caption         =   ""&Previous""")
		WriteTs.WriteLine("      Height          =   375")
		WriteTs.WriteLine("      Left            =   1440")
		WriteTs.WriteLine("      TabIndex        =   68")
		WriteTs.WriteLine("      Top             =   5520")
		WriteTs.WriteLine("      Width           =   1095")
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("   Begin VB.CommandButton cmdnext ")
		WriteTs.WriteLine("      Caption         =   ""&Next""")
		WriteTs.WriteLine("      Height          =   375")
		WriteTs.WriteLine("      Left            =   2760")
		WriteTs.WriteLine("      TabIndex        =   67")
		WriteTs.WriteLine("      Top             =   5520")
		WriteTs.WriteLine("      Width           =   1095")
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("Begin VB.Frame Frame1")
		WriteTs.WriteLine("Caption = ""Frame1""")
		WriteTs.WriteLine("Height = 5175")
		WriteTs.WriteLine("Left = 120")
		WriteTs.WriteLine("TabIndex = 0")
		WriteTs.WriteLine("Top = 120")
		WriteTs.WriteLine("Width = 6375")
		
		For Fieldidx = 0 To Rs.Fields.Count - 1
			
			lblfield = lblfield & "      Begin VB.TextBox lblfield " & vbCrLf
			lblfield = lblfield & "         BackColor       =   &H8000000F&" & vbCrLf
			lblfield = lblfield & "         BorderStyle     =   0  'None" & vbCrLf
			lblfield = lblfield & "         Height          =   285" & vbCrLf
			lblfield = lblfield & "         Index          =   " & CStr(Fieldidx) & vbCrLf
			lblfield = lblfield & "         Left            =   120" & vbCrLf
			lblfield = lblfield & "         Locked          =   -1  'True" & vbCrLf
			lblfield = lblfield & "         TabIndex        =   5" & vbCrLf
			lblfield = lblfield & "         TabStop         =   0   'False" & vbCrLf
			lblfield = lblfield & "         Text            =   """ & Rs.Fields(Fieldidx).Name & ":""" & vbCrLf
			lblfield = lblfield & "         Top             =   " & CStr(240 + Fieldidx * 480) & vbCrLf
			lblfield = lblfield & "         Width           =   1935" & vbCrLf
			lblfield = lblfield & "      End"
			
			txtfield = txtfield & "      Begin VB.TextBox txtfield " & vbCrLf
			txtfield = txtfield & "         Height          =   285" & vbCrLf
			txtfield = txtfield & "         Index           =   " & CStr(Fieldidx) & vbCrLf
			txtfield = txtfield & "         Left            =   2280" & vbCrLf
			txtfield = txtfield & "         MaxLength       =   " & Rs.Fields(Fieldidx).DefinedSize & vbCrLf
			txtfield = txtfield & "         TabIndex        =   1" & vbCrLf
			txtfield = txtfield & "         Top             =   " & CStr(240 + Fieldidx * 480) & vbCrLf
			txtfield = txtfield & "         Width           =   3375" & vbCrLf
			txtfield = txtfield & "      End"
			
			txtrs = txtrs & vbTab & "With txtfield(" & Fieldidx & ")" & vbCrLf
			txtrs = txtrs & vbTab & vbTab & "Set .DataSource = Rs" & vbCrLf
			txtrs = txtrs & vbTab & vbTab & ".DataField = Rs(" & Fieldidx & ").Name" & vbCrLf
			txtrs = txtrs & vbTab & "end with" & vbCrLf
			
		Next 
		
		WriteTs.WriteLine(lblfield)
		WriteTs.WriteLine(txtfield)
		
		WriteTs.WriteLine("   End")
		WriteTs.WriteLine("   Begin VB.Label lblrecordcount ")
		WriteTs.WriteLine("      Height          =   255")
		WriteTs.WriteLine("      Left            =   120")
		WriteTs.WriteLine("      TabIndex        =   66")
		WriteTs.WriteLine("      Top             =   7560")
		WriteTs.WriteLine("      Width           =   6495")
		WriteTs.WriteLine("   End")
		
		WriteTs.WriteLine("End")
		WriteTs.WriteLine("Attribute VB_Name = ""Form1""")
		WriteTs.WriteLine("Attribute VB_GlobalNameSpace = False")
		WriteTs.WriteLine("Attribute VB_Creatable = False")
		WriteTs.WriteLine("Attribute VB_PredeclaredId = True")
		WriteTs.WriteLine("Attribute VB_Exposed = False")
		WriteTs.WriteLine("''Code Ready " & VB6.Format(Now, "yyyy/MM/dd HH:mm:ss"))
		WriteTs.WriteLine("''Please add reference Microsoft ActiveX Objects Data Library")
		WriteTs.WriteLine("Dim Cn As New ADODB.Connection")
		WriteTs.WriteLine("Dim Rs As New ADODB.Recordset")
		WriteTs.WriteLine("Dim FindCriteria As String")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub cmdFind_Click()")
		WriteTs.WriteLine("    Dim Criteria As String")
		WriteTs.WriteLine("    Dim Message As String")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    Message = ""Please enter find criteria""")
		WriteTs.WriteLine("    Criteria = InputBox(Message, ""Find"", FindCriteria)")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    If Len(Criteria) = 0 Then")
		WriteTs.WriteLine("        Exit Sub")
		WriteTs.WriteLine("    End If")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    FindCriteria = Criteria")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    Call Find")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub cmdFindNext_Click()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    Call Find")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub Find()")
		WriteTs.WriteLine("     ")
		WriteTs.WriteLine("    Dim bmark As Variant")
		WriteTs.WriteLine("     ")
		WriteTs.WriteLine("    bmark = Rs.Bookmark")
		WriteTs.WriteLine("    Rs.Find FindCriteria, , adSearchForward, 0")
		WriteTs.WriteLine("        ")
		WriteTs.WriteLine("    If Rs.EOF = True Then")
		WriteTs.WriteLine("        Rs.Bookmark = bmark")
		WriteTs.WriteLine("        MsgBox ""No match"", vbInformation")
		WriteTs.WriteLine("    End If")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("End Sub")
		
		WriteTs.WriteLine("Private Sub Form_Load()")
		WriteTs.WriteBlankLines((1))
		WriteTs.WriteLine(vbTab & "Dim connstr as string" & vbCrLf)
		WriteTs.WriteLine(vbTab & "Dim SQL as String" & vbCrLf)
		WriteTs.WriteBlankLines((1))
		WriteTs.WriteLine(vbTab & "connstr =""" & ConnectionString & """")
		WriteTs.WriteLine(vbTab & "Cn.open connstr")
		WriteTs.WriteBlankLines((1))
		WriteTs.WriteLine(vbTab & "SQL = """ & Rs.Source & """")
		WriteTs.WriteLine(vbTab & "Rs.open SQL,cn,1,3")
		WriteTs.WriteBlankLines((1))
		WriteTs.WriteLine(txtrs)
		WriteTs.WriteLine("        Call SetScrollBar")
		WriteTs.WriteLine("        FindCriteria = Rs(0).Name & ""=''""")
		WriteTs.WriteLine("End Sub")
		
		WriteTs.WriteLine("Private Sub cmdAdd_Click()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    Rs.AddNew")
		WriteTs.WriteLine("    txtfield(0).SetFocus")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub cmdCancelUpdate_Click()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    Rs.CancelUpdate")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub cmdClose_Click()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    Unload Me")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub cmdDelete_Click()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    If Rs.AbsolutePage > 0 Then")
		WriteTs.WriteLine("        Rs.Delete")
		WriteTs.WriteLine("        Call SetScrollBar")
		WriteTs.WriteLine("    End If")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub cmdend_Click()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    VScroll1.Value = VScroll1.Max")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub cmdnext_Click()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    If Rs.EOF = False Then")
		WriteTs.WriteLine("        VScroll1.Value = VScroll1.Value + 1")
		WriteTs.WriteLine("    End If")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub cmdPrevious_Click()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    If Rs.BOF = False Then")
		WriteTs.WriteLine("        VScroll1.Value = VScroll1.Value - 1")
		WriteTs.WriteLine("    End If")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub cmdTop_Click()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    VScroll1.Value = VScroll1.Min")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub cmdUpdate_Click()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    Rs.Update")
		WriteTs.WriteLine("    Call SetScrollBar")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("Private Sub Form_Unload(Cancel As Integer)")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    Rs.Close")
		WriteTs.WriteLine("    Cn.Close")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    Set Rs = Nothing")
		WriteTs.WriteLine("    Set Cn = Nothing")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub SetScrollBar()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    With VScroll1")
		WriteTs.WriteLine("        .Max = Rs.RecordCount")
		WriteTs.WriteLine("        .Value = 0")
		WriteTs.WriteLine("    End With")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub VScroll1_Change()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    Rs.Move VScroll1.Value, 1")
		WriteTs.WriteLine("    Call ShowRecordCount")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("End Sub")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("Private Sub ShowRecordCount()")
		WriteTs.WriteLine("    ")
		WriteTs.WriteLine("    lblrecordcount.Caption = Rs.AbsolutePosition & "" / "" & Rs.RecordCount")
		WriteTs.WriteLine("")
		WriteTs.WriteLine("End Sub")
		
		WriteTs.Close()
		
		
	End Sub
End Module