Option Strict Off
Option Explicit On
Module Bas_BuildASPEditForm
	Private Const ModuleName As String = "Bas_BuildASPEditForm"
	
    Public Sub BuildASPEditForm(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Call BuildFramework(ConnectionString, Rs, Path)

        Call BuildControls(ConnectionString, Rs, Path)
        Call BuildControlLabels(ConnectionString, Rs, Path)
        Call BuildSave(ConnectionString, Rs, Path & "\save")
        Call BuildFrmFind(ConnectionString, Rs, Path & "\forms\editform\find")

        Call BuildForm(ConnectionString, Rs, Path)
        Call BuildForm(ConnectionString, Rs, Path & "\forms\editform\forms\TwoColumn", 2)

    End Sub

    Private Sub BuildForm(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String, Optional ByVal numberOfColumns As Short = 1)

        Dim fieldidx As Short
        Dim code As String
        Dim Criteria As String

        Dim columnRs As New ADODB.Recordset
        Dim Cn As New ADODB.Connection


        Dim codeFileName As String = ""

        Cn.Open(ConnectionString)

        columnRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        code = ""

        For fieldidx = 0 To Rs.Fields.Count - 1

            code = code & "<!-- #include File=""controls/" & Rs.Fields(fieldidx).Name & ".asp"" -->" & vbCrLf

        Next

        For fieldidx = 0 To Rs.Fields.Count - 1

            code = code & "<!-- #include File=""controlLabels/" & Rs.Fields(fieldidx).Name & ".asp"" -->" & vbCrLf

        Next

        For fieldidx = 0 To Rs.Fields.Count - 1

            'code = code & "<!-- #include File=""controlRemarks/" & Rs.Fields(fieldidx).Name & ".asp"" -->" & vbCrLf

        Next

        code = code & "" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\buttons\btnReload.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\buttons\btnPreviousRecord.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\buttons\btnNextRecord.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\buttons\btnFirstRecord.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\buttons\btnLastRecord.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\buttons\btnDelete.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\buttons\btnUpdate.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\buttons\btnAddnew.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\buttons\btnUpdateNew.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\buttons\btnCancel.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\recordControl\top.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\recordControl\bottom.asp"" -->" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<!-- #include File=""forms\editform\toolbar.asp"" -->" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "Dim formId" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "Sub showForm(Cn, Rs)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	Dim tempRs" & vbCrLf

        code = code & "	Dim tempSQL" & vbCrLf

        code = code & "	Dim hideForm" & vbCrLf

        code = code & "	Dim nextFormAction" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "nextFormAction = ""Update""" & vbCrLf

        code = code & "	select case formAction" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	case ""addNew""" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		set tempRs = server.createObject(""adodb.recordset"")" & vbCrLf

        code = code & "		tempSQL = ""select top 1 * from ("" & Rs.source & "")""" & vbCrLf

        code = code & "		tempRs.open tempSQL, CN, 1, 3" & vbCrLf

        code = code & "		tempRs.addnew" & vbCrLf

        code = code & "		Call setDefaultValue(tempRs)" & vbCrLf

        code = code & "		tempRs.cancelUpdate" & vbCrLf

        code = code & "		tempRs.close" & vbCrLf

        code = code & "		formId = ""form"" & rs.recordcount + 1" & vbCrLf

        code = code & "nextFormAction = ""updateNew""" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	case ""updateNew""" & vbCrLf

        code = code & "		" & vbCrLf

        code = code & "		if isSaved = true then" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "			Call setDefaultValue(Rs)" & vbCrLf

        code = code & "			formId = ""form"" & rs.absolutePosition" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		else" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "			formId = ""form"" & rs.recordcount + 1" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		end if" & vbCrLf

        code = code & "		" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	case ""Update""" & vbCrLf

        code = code & "		" & vbCrLf

        code = code & "		if isSaved = true then" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "			Call setDefaultValue(Rs)" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "		end if " & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		formId = ""form"" & rs.absolutePosition		" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	case else" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "		Call setDefaultValue(Rs)" & vbCrLf

        code = code & "		formId = ""form"" & abs(rs.absolutePosition)			" & vbCrLf

        code = code & "		if rs.recordcount = 0 then" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "			hideForm = true" & vbCrLf

        code = code & "		end if" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	end select" & vbCrLf

        code = code & "" & vbCrLf

        code = code & " Dim lookupRs" & vbCrLf

        code = code & " set lookupRs = server.createObject(""adodb.recordset"")" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf

        code = code & "  <tr>" & vbCrLf

        code = code & "    <td>" & vbCrLf

        code = code & "      <b><font face=""Arial, Helvetica, sans-serif"" size=""2"">Total record(s):</font></b> "

        code = code & "      <font face=""Arial, Helvetica, sans-serif"" size=""2""><%=rs.recordcount%> </font><br>" & vbCrLf

        code = code & "    </td>" & vbCrLf

        code = code & "  </tr>" & vbCrLf

        code = code & "</table>	" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "call showToolbar" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "<form name=""<%=formId%>"" method=""post"" action=""edit.asp"" style=""margin-top:0;"" >" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "		call recordControl_top" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "if len(message) > 0 then" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "<p><font face=""Arial, Helvetica, sans-serif"" size=""2"" color=""red""><%=message%></font></p>" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "end if" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "if hideForm = false then" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""editform"">" & vbCrLf

        For fieldidx = 0 To Rs.Fields.Count - 1

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(fieldidx).Properties(1).Value & "'"

            columnRs.Filter = Criteria

            If fieldidx Mod numberOfColumns = 0 Then

                code = code & "    <tr height = ""40""> " & vbCrLf

            End If

            code = code & "      <td width=""30%"">" & vbCrLf

            code = code & " <% call showControlLabel_" & Rs.Fields(fieldidx).Name & "%>" & vbCrLf

            code = code & "      </td>" & vbCrLf

            code = code & "      <td> " & vbCrLf

            code = code & " <% call showControls_" & Rs.Fields(fieldidx).Name & "(Value(" & fieldidx & "), """", """")%>"

            If columnRs("IS_NULLABLE").Value = False Then

                code = code & "<font face=""Arial, Helvetica, sans-serif"" size=""2"">* REQUIRED</font>" & vbCrLf

            End If

            code = code & "<br><font face=""Arial, Helvetica, sans-serif"" size=""2"" color=""red""><%=fieldMessage(" & fieldidx & ")%></font>" & vbCrLf

            code = code & "      </td>" & vbCrLf

            If (fieldidx + 1) Mod numberOfColumns = 0 Then

                code = code & "    </tr>" & vbCrLf

            End If

        Next

        code = code & "  </table>" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "	Call recordControl_bottom" & vbCrLf

        code = code & "end if" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<input type=""hidden"" name=""formAction"" value=""<%=nextFormAction%>"">" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<input type=""hidden"" name=""Pos"" value=""<%=Rs.absolutePosition%>"">" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "</form>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\record_onkeydown.js""></script>" & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\record_onfocus.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\record_onfocusout.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\column_onfocus.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\column_onfocusout.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\deleteRecord.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\moveNextRecord.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\movePreviousRecord.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\moveFirstRecord.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\moveLastRecord.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\cancelUpdate.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\addNew.js""></SCRIPT>" & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\updateNew.js""></SCRIPT>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "End sub" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "sub recordControl()" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	select case formAction" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	case ""addNew""" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "                		 call showBtnUpdateNew(formId)" & vbCrLf

        code = code & "			 call showBtnCancel(formId)" & vbCrLf

        code = code & "			" & vbCrLf

        code = code & "			 exit sub" & vbCrLf

        code = code & "		" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	case ""updateNew"" " & vbCrLf

        code = code & "			" & vbCrLf

        code = code & "			if isSaved = false then" & vbCrLf

        code = code & "			" & vbCrLf

        code = code & "               		 	call showBtnUpdateNew(formId)" & vbCrLf

        code = code & "				call showBtnCancel(formId)" & vbCrLf

        code = code & "				" & vbCrLf

        code = code & "				exit sub" & vbCrLf

        code = code & "				" & vbCrLf

        code = code & "			end if" & vbCrLf

        code = code & "	case else " & vbCrLf

        code = code & "" & vbCrLf

        code = code & "        if  Rs.EOF = True and Rs.BOF = True  then" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "            		call showBtnAddnew(Rs)" & vbCrLf

        code = code & "		exit sub" & vbCrLf

        code = code & "		" & vbCrLf

        code = code & "        end if" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	end select" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "        " & vbCrLf

        code = code & "            	call showBtnReload()" & vbCrLf

        code = code & "            	call showBtnFirstRecord" & vbCrLf

        code = code & "            	call showBtnPreviousRecord" & vbCrLf

        code = code & "             	call showBtnNextRecord" & vbCrLf

        code = code & "             	call showBtnLastRecord" & vbCrLf

        code = code & "  	call showBtnDelete(Rs)" & vbCrLf

        code = code & "  	call showBtnAddnew(Rs)" & vbCrLf

        code = code & "	call showBtnUpdate(Rs)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "end sub" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "%>" & vbCrLf

        codeFileName = Path & "\form.asp"

        Call WriteCodeFile(codeFileName, code)

        columnRs.Close()

        Cn.Close()


    End Sub



End Module