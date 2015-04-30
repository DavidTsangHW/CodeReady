Option Strict Off
Option Explicit On
Module Bas_BuildASPBlankForm
    Private Const ModuleName As String = "Bas_ASPBlankForm"
    

    Public Sub BuildASPBlankForm(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal HTMLForm As String)

        Dim Path As String

        Path = Fso.GetParentFolderName(HTMLForm)

        Call BuildFramework(ConnectionString, Rs, Path)
        Call BuildControls(ConnectionString, Rs, Path)
        Call BuildControlLabels(ConnectionString, Rs, Path)
        Call BuildSave(ConnectionString, Rs, Path & "\save")
        Call BuildForm(ConnectionString, Rs, Path)
        Call builtConfirmPage(ConnectionString, Rs, Path & "\forms\blankform")

    End Sub
    Private Sub BuildForm(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim fieldidx As Short
        Dim code As String
        Dim Criteria As String

        Dim columnRs As New ADODB.Recordset
        Dim Cn As New ADODB.Connection


        Cn.Open(ConnectionString)

        columnRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        code = ""

        code = code & "<!-- #include File=""forms/blankform/confirm.asp"" -->" & vbCrLf

        For fieldidx = 0 To Rs.Fields.Count - 1

            code = code & "<!-- #include File=""controls/" & Rs.Fields(fieldidx).Name & ".asp"" -->" & vbCrLf

        Next

        For fieldidx = 0 To Rs.Fields.Count - 1

            code = code & "<!-- #include File=""controlLabels/" & Rs.Fields(fieldidx).Name & ".asp"" -->" & vbCrLf

        Next

        code = code & "<%" & vbCrLf

        code = code & "Sub showForm(cn, rs)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & " Dim lookupRs" & vbCrLf

        code = code & " set lookupRs = server.createObject(""adodb.recordset"")" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	Dim tempRs" & vbCrLf

        code = code & "	Dim tempSQL" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	select case formAction" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	case ""updateNew""" & vbCrLf

        code = code & "		" & vbCrLf

        code = code & "		if isSaved = true then" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "			Call setDefaultValue(Rs)" & vbCrLf

        code = code & "			call showConfirm" & vbCrLf

        code = code & "			exit sub" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		end if" & vbCrLf

        code = code & "		" & vbCrLf

        code = code & "	case else" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "		set tempRs = server.createObject(""adodb.recordset"")" & vbCrLf

        code = code & "		tempSQL = ""select top 1 * from ("" & Rs.source & "")""" & vbCrLf

        code = code & "		tempRs.open tempSQL, CN, 1, 3" & vbCrLf

        code = code & "		tempRs.addnew" & vbCrLf

        code = code & "		Call setDefaultValue(tempRs)" & vbCrLf

        code = code & "		tempRs.cancelUpdate" & vbCrLf

        code = code & "		tempRs.close" & vbCrLf

        code = code & "		formId = ""form"" & rs.recordcount + 1		" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	end select" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "if len(message) > 0 then" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "<p><font face=""Arial, Helvetica, sans-serif"" size=""2"" color=""red""><%=message%></font></p>" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "end if" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & " <center>" & vbCrLf

        code = code & "<form name=""frm_" & Rs.Fields(0).Properties(1).Value & """ method=""post"" action=""edit.asp"">" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""editform"">" & vbCrLf

        For fieldidx = 0 To Rs.Fields.Count - 1

            'If it is an auto increment field
            '16 Oct 2009

            If Rs.Fields(fieldidx).Properties.Item("ISAUTOINCREMENT").Value = False Then

                'Search description of field
                '25 August 2010
                Criteria = "COLUMN_NAME = '" & Rs.Fields(fieldidx).Name & "'"
                Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(fieldidx).Properties(1).Value & "'"

                columnRs.Filter = Criteria

                code = code & "    <tr height = ""40""> " & vbCrLf

                code = code & "      <td width=""30%"">" & vbCrLf

                code = code & " <% call showControlLabel_" & Rs.Fields(fieldidx).Name & "%>" & vbCrLf

                code = code & "      </td>" & vbCrLf

                code = code & "      <td>" & vbCrLf

                code = code & " <% call showControls_" & Rs.Fields(fieldidx).Name & "(Value(" & fieldidx & "), """", """")%>"

                If columnRs("IS_NULLABLE").Value = False Then

                    code = code & "<font face=""Arial, Helvetica, sans-serif"" size=""2"">* REQUIRED</font>" & vbCrLf

                End If

                code = code & "<br><font face=""Arial, Helvetica, sans-serif"" size=""2"" color=""red""><%=fieldMessage(" & fieldidx & ")%></font>" & vbCrLf

                code = code & "      </td>" & vbCrLf

                code = code & "    </tr>" & vbCrLf

            End If

        Next


        code = code & "  </table>" & vbCrLf

        code = code & "<br>" & vbCrLf

        code = code & "	<input type=""hidden"" name=""formAction"" value=""updateNew"">" & vbCrLf

        code = code & "          <input type=""submit"" name=""Submit"" value=""Submit"">" & vbCrLf

        code = code & "          <input type=""button"" name=""Reset"" value=""Reset"" onClick=""javascript:window.location = '.'"">" & vbCrLf

        code = code & " </center>" & vbCrLf

        code = code & "</form>" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & " End sub" & vbCrLf

        code = code & "%>" & vbCrLf

        Call WriteCodeFile(Path & "\form.asp", code)

        columnRs.Close()

        Cn.Close()


    End Sub
 

    Private Sub builtConfirmPage(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Cn As New ADODB.Connection

        Dim ColumnRS As New ADODB.Recordset

        Dim Fieldidx As Short

        Dim Criteria As String

        Dim code As String

        Cn.Open(ConnectionString)

        ColumnRS = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        code = "<%" & vbCrLf

        code = code & "Sub showConfirm" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "<br>" & vbCrLf

        code = code & "<font face=""Arial, Helvetica, sans-serif"" size=""2""><b>Record saved</b></font>" & vbCrLf

        code = code & "  <center>" & vbCrLf

        code = code & "  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf

        For Fieldidx = 0 To Rs.Fields.Count - 1

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(Fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            ColumnRS.Filter = Criteria

            code = code & "    <tr height = ""45""> " & vbCrLf

            code = code & "      <td width=""30%""><font face=""Arial, Helvetica, sans-serif"" size=""2"">" & PrintColumnName(ColumnRS, Rs.Fields(Fieldidx).Name) & ":</font></td>" & vbCrLf

            code = code & "      <td> " & vbCrLf

            code = code & "      <td><font face=""Arial, Helvetica, sans-serif"" size=""2""><%=Value(" & Fieldidx & ")%></font></td>" & vbCrLf

            code = code & "      </td>" & vbCrLf

            code = code & "    </tr>" & vbCrLf

        Next

        code = code & "  </table>" & vbCrLf

        code = code & "<p>" & vbCrLf

        code = code & "<input type=""button"" value = ""  Create next  "" onClick=""javascript:window.navigate('.')""> " & vbCrLf

        code = code & "</p>" & vbCrLf

        code = code & "  </center>" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "end sub" & vbCrLf

        code = code & "%>" & vbCrLf

        Call WriteCodeFile(Path & "\confirm.asp", code)

        ColumnRS.Close()

        Cn.Close()

    End Sub
End Module