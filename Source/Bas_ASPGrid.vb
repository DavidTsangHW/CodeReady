Option Strict Off
Option Explicit On
Module Bas_ASPGrid
    Private Const ModuleName As String = "Bas_ASPGrid"
	Public Sub BuildASPGrid(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Call BuildFramework(ConnectionString, Rs, Path)

        Call BuildForm(ConnectionString, Rs, Path)

        Call BuildControls(ConnectionString, Rs, Path)
        Call BuildControlLabels(ConnectionString, Rs, Path)
        Call BuildSave(ConnectionString, Rs, Path & "\save")
        Call BuildFrmFind(ConnectionString, Rs, Path & "\forms\grid\find")

        Call BuildAddnew(ConnectionString, Rs, Path & "\forms\grid\rows")
        Call BuildColumnHeaders_Action(ConnectionString, Rs, Path & "\forms\grid\rows\columnHeaders")

        Call BuildColumnHeaders_Fields(ConnectionString, Rs, Path & "\forms\grid\rows\columnHeaders\sort")
        Call BuildColumnHeaders_Fields_Text(ConnectionString, Rs, Path & "\forms\grid\rows\columnHeaders\text")
        Call BuildColumnHeaders_Fields_Abstract(ConnectionString, Rs, Path & "\forms\grid\rows\columnHeaders\abstract")

        Call BuildColumnHeaders_Fields(ConnectionString, Rs, Path & "\forms\grid\rows\columnHeaders")

        Call BuildRowColumnLabeling(ConnectionString, Rs, Path & "\forms\grid\rows\columnHeaders")
        Call BuildRowColumnHeader(ConnectionString, Rs, Path & "\forms\grid\rows")
        Call BuildRowMessage(ConnectionString, Rs, Path & "\forms\grid\rows")
        Call BuildRowRecord(ConnectionString, Rs, Path & "\forms\grid\rows")

    End Sub

    Private Sub BuildRowColumnLabeling(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim WriteTs As Scripting.TextStream

        WriteTs = Fso.OpenTextFile(Path & "\columnLabeling.asp", Scripting.IOMode.ForWriting, True)

        WriteTs.WriteLine("<%")

        WriteTs.WriteLine("sub showColumnHeader_Label")

        WriteTs.WriteLine("%>")

        WriteTs.WriteLine(" <tr bgcolor=""#EEEEEE"">")

        WriteTs.WriteLine("<td>")

        WriteTs.WriteLine("	</td>")

        WriteTs.WriteLine("	<td>")

        WriteTs.WriteLine("	</td>")

        WriteTs.WriteLine("")

        WriteTs.WriteLine("<%	for idx = 0 to " & Rs.Fields.Count - 1 & " %>")

        WriteTs.WriteLine("	")

        WriteTs.WriteLine("	<td >  	")

        WriteTs.WriteLine("		<div language=""javascript"" onmouseover=""column_onfocus(<%=idx%>)"">")

        WriteTs.WriteLine("		<font face=""Arial, Helvetica, sans-serif"" size=""2"">")

        WriteTs.WriteLine("		<center><%=chr(idx+65)%></center>")

        WriteTs.WriteLine("		</font>")

        WriteTs.WriteLine("		</div>")

        WriteTs.WriteLine("	</td>")

        WriteTs.WriteLine("<%	next %>")

        WriteTs.WriteLine("	<td>")

        WriteTs.WriteLine("	</td>")

        WriteTs.WriteLine("	</tr>")

        WriteTs.WriteLine("  <tr>")

        WriteTs.WriteLine("")

        WriteTs.WriteLine("<%")

        WriteTs.WriteLine("")

        WriteTs.WriteLine("end sub")

        WriteTs.WriteLine("%>")

        WriteTs.Close()


    End Sub

    Private Sub BuildColumnHeaders_Fields_Text(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim code As String = ""
        Dim codeFilename As String = ""
        Dim Fieldidx As Short

        Dim CN As New ADODB.Connection

        Dim ColumnRS As New ADODB.Recordset

        Dim Criteria As String

        CN.Open(ConnectionString)

        ColumnRS = CN.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFilename = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

            code = "<%" & vbCrLf

            code = code & "sub showColumnHeader_" & Rs.Fields(Fieldidx).Name & vbCrLf

            code = code & "%>" & vbCrLf

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(Fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            ColumnRS.Filter = Criteria

            code = code & "    <td  bgcolor=""#999999""><font face=""Arial, Helvetica, sans-serif"" size=""2"">" & vbCrLf
            code = code & PrintColumnName(ColumnRS, Rs.Fields(Fieldidx).Name) & vbCrLf
            code = code & "     </font></td>" & vbCrLf

            code = code & "<%" & vbCrLf

            code = code & "end sub" & vbCrLf

            code = code & "%>" & vbCrLf

            Call WriteCodeFile(codeFileName, code)

        Next

    End Sub
    Public Sub BuildControls_RadioButton(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim code As String = ""
        Dim codeFilename As String = ""

        Dim Fieldidx As Short

        For Fieldidx = 0 To Rs.Fields.Count - 1

            Select Case Rs.Fields(Fieldidx).Type

                Case ADODB.DataTypeEnum.adBoolean

                    codeFilename = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

                    code = "<%" & vbCrLf

                    code = code & "sub showControls_" & Rs.Fields(Fieldidx).Name & "(defaultValue, properties, javascript)" & vbCrLf

                    code = code & "%>" & vbCrLf

                    code = code & "<font face=""Arial, Helvetica, sans-serif"" size=""2"">" & vbCrLf

                    code = code & "        <input type=""radio"" name=""txtfield_" & Rs.Fields(Fieldidx).Name & """ value=""True""" & vbCrLf

                    code = code & " <%if Value(" & Fieldidx & ") = ""True"" then" & vbCrLf
                    code = code & "     Response.Write "" checked"" " & vbCrLf
                    code = code & " end if%>" & vbCrLf

                    code = code & "> Yes" & vbCrLf

                    code = code & "        <input type=""radio"" name=""txtfield_" & Rs.Fields(Fieldidx).Name & """ value=""False""" & vbCrLf

                    code = code & " <%if Value(" & Fieldidx & ") = ""False"" then" & vbCrLf
                    code = code & "     Response.Write "" checked"" " & vbCrLf
                    code = code & " end if%>" & vbCrLf

                    code = code & "> No" & vbCrLf

                    code = code & "" & vbCrLf

                    code = code & "</font>" & vbCrLf

                    code = code & "<%" & vbCrLf

                    code = code & " End sub" & vbCrLf

                    code = code & "%>" & vbCrLf

                    Call WriteCodeFile(codeFilename, code)


            End Select

        Next

        'UPGRADE_NOTE: Object WriteTs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        codeFilename = Nothing

    End Sub
    Private Sub BuildColumnHeaders_Fields_Abstract(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim code As String = ""
        Dim codeFilename As String = ""
        Dim Fieldidx As Short

        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFilename = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

            code = "<%" & vbCrLf

            code = code & "sub showColumnHeader_" & Rs.Fields(Fieldidx).Name & vbCrLf

            code = code & "%>" & vbCrLf

            code = code & "    <td  bgcolor=""#999999"">" & vbCrLf
            code = code & "     </td>" & vbCrLf

            code = code & "<%" & vbCrLf

            code = code & "end sub" & vbCrLf

            code = code & "%>" & vbCrLf

            Call WriteCodeFile(codeFileName, code)

        Next

    End Sub

    Private Sub BuildColumnHeaders_Fields(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim code As String = ""
        Dim codeFilename As String = ""
        Dim Fieldidx As Short

        Dim CN As New ADODB.Connection

        Dim ColumnRS As New ADODB.Recordset

        Dim Criteria As String

        CN.Open(ConnectionString)

        ColumnRS = CN.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFilename = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

            code = "<%" & vbCrLf

            code = code & "sub showColumnHeader_" & Rs.Fields(Fieldidx).Name & vbCrLf

            code = code & "%>" & vbCrLf

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(Fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            ColumnRS.Filter = Criteria

            code = code & "    <td  bgcolor=""#999999""><font face=""Arial, Helvetica, sans-serif"" size=""2"">" & vbCrLf
            code = code & "      <a href= ""edit.asp?sort=" & Rs.Fields(Fieldidx).Name & """>" & PrintColumnName(ColumnRS, Rs.Fields(Fieldidx).Name) & vbCrLf
            code = code & "      </a> </font></td>"

            code = code & "<%" & vbCrLf

            code = code & "end sub" & vbCrLf

            code = code & "%>" & vbCrLf

            Call WriteCodeFile(codeFileName, code)

        Next

    End Sub
    Private Sub BuildColumnHeaders_Action(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFilename As String = ""

        codeFilename = Path & "\Action.asp"

        Code = Code & "<%" & vbCrLf
        Code = Code & "sub showColumnHeader_Action" & vbCrLf
        Code = Code & "%>" & vbCrLf

        Code = Code & "<td  bgcolor=""#999999""><font face=""Arial, Helvetica, sans-serif"" size=""2"">Action" & vbCrLf
        Code = Code & "       </font></td>" & vbCrLf

        Code = Code & "<%" & vbCrLf
        Code = Code & "end sub" & vbCrLf
        Code = Code & "%>" & vbCrLf

        WriteCodeFile(codeFilename, Code)

    End Sub
    Private Sub BuildRowColumnHeader(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFilename As String = ""
        Dim Fieldidx As Short

        codeFilename = Path & "\columnHeader.asp"

        For Fieldidx = 0 To Rs.Fields.Count - 1

            Code = Code & "<!-- #include File=""columnHeaders\" & Rs.Fields(Fieldidx).Name & ".asp"" -->"

        Next

        Code = Code & "<!-- #include File=""columnHeaders\Action.asp"" -->" & vbCrLf
        Code = Code & "<!-- #include File=""columnHeaders\rowMarker.asp"" -->" & vbCrLf
        Code = Code & "<!-- #include File=""columnHeaders\rowLabeling.asp"" -->" & vbCrLf
        Code = Code & "<!-- #include File=""columnHeaders\columnLabeling.asp"" -->" & vbCrLf

        Code = Code & "<%" & vbCrLf

        Code = Code & "sub showRowColumnHeader" & vbCrLf

        Code = Code & "%>" & vbCrLf

        Code = Code & "  <tr> " & vbCrLf

        Code = Code & "    <td bgcolor=""#CCCCCC"" colspan=""" & Rs.Fields.Count + 3 & """></td>" & vbCrLf

        Code = Code & "  </tr>" & vbCrLf

        Code = Code & "<%" & vbCrLf

        Code = Code & "	  call showColumnHeader_Label" & vbCrLf

        Code = Code & "	  call showColumnHeader_rowMarker" & vbCrLf

        Code = Code & "	  call showColumnHeader_rowLabeling" & vbCrLf

        For Fieldidx = 0 To Rs.Fields.Count - 1

            Code = Code & "	  call showColumnHeader_" & Rs.Fields(Fieldidx).Name & vbCrLf

        Next

        Code = Code & "  	  call showColumnHeader_Action" & vbCrLf

        Code = Code & "%>" & vbCrLf

        Code = Code & "  </tr>" & vbCrLf

        Code = Code & "  <tr> " & vbCrLf

        Code = Code & "    <td bgcolor=""#666666"" colspan=""" & Rs.Fields.Count + 3 & """></td>" & vbCrLf

        Code = Code & "  </tr>" & vbCrLf

        Code = Code & "<%" & vbCrLf

        Code = Code & "end sub" & vbCrLf

        Code = Code & "%>" & vbCrLf

        Call WriteCodeFile(codeFilename, Code)

    End Sub
    Private Sub BuildRowMessage(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFilename As String = ""

        codeFilename = Path & "\message.asp"

        Code = "<%" & vbCrLf

        Code = Code & "	sub showRowMessage(rowId, Message)" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "		" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "%>" & vbCrLf

        Code = Code & "		<tr id=msg_row<%=rowId%>>" & vbCrLf

        Code = Code & "			<td bgcolor=""#CCCCCC"">" & vbCrLf

        Code = Code & "				<!-- Marker -->" & vbCrLf

        Code = Code & "			</td>" & vbCrLf

        Code = Code & "			<td bgcolor=""#CCCCCC"">" & vbCrLf

        Code = Code & "				<!-- Numbering -->" & vbCrLf

        Code = Code & "			</td>" & vbCrLf

        Code = Code & "			<td colspan=""" & Rs.Fields.Count & """ align=""center"">	" & vbCrLf

        Code = Code & "<div LANGUAGE=javascript onmouseover=""return record_onfocus(row<%=rowid%>, form<%=rowid%>)"">" & vbCrLf

        Code = Code & "				<font face=""Arial, Helvetica, sans-serif"" size=""2"" color=""red"">" & vbCrLf

        Code = Code & "					<%=Message%>" & vbCrLf

        Code = Code & "				</font>" & vbCrLf

        Code = Code & "</div>" & vbCrLf

        Code = Code & "			</td>" & vbCrLf

        Code = Code & "				<!-- Action -->" & vbCrLf

        Code = Code & "			<td>" & vbCrLf

        Code = Code & "			</td>" & vbCrLf

        Code = Code & "		</tr>	" & vbCrLf

        Code = Code & "<%" & vbCrLf

        Code = Code & "	end sub" & vbCrLf

        Code = Code & "%>" & vbCrLf

        Call WriteCodeFile(codeFilename, Code)

    End Sub
    Private Sub BuildRowRecord(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim code As String = ""
        Dim codeFilename As String = ""
        Dim Fieldidx As Short

        Dim setReadOnly As String
        Dim fieldProperties As String

        codeFilename = Path & "\Record.asp"

        For Fieldidx = 0 To Rs.Fields.Count - 1

            code = code & "<!-- #include File=""../../../controls/" & Rs.Fields(Fieldidx).Name & ".asp"" -->" & vbCrLf

        Next

        code = code & "<!-- #include File=""controlEvents.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""rowMarker.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""rowLabeling.asp"" -->" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "Sub showRowRecord(Rs)" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & " 	Call setDefaultValue(Rs)" & vbCrLf

        code = code & "	    dim fieldProperties" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "	if request(""formAction"") = ""Update"" then	" & vbCrLf

        code = code & "		if defaultPosition = rs.absoluteposition then" & vbCrLf

        code = code & " 			for idx = 0 to " & Rs.Fields.Count - 1 & vbCrLf

        code = code & " 				rowFieldMessage(idx) = fieldMessage(idx)" & vbCrLf

        code = code & " 				value(idx) = savedValue(idx)" & vbCrLf

        code = code & " 			next" & vbCrLf

        code = code & " 		end if" & vbCrLf

        code = code & "	end if" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "  <tr id = row<%=Rs.AbsolutePosition%> align=""left"" valign=""top""> " & vbCrLf

        code = code & "    <form name=""form<%=Rs.AbsolutePosition%>"" method=""post"" action=""edit.asp"">" & vbCrLf

        code = code & "	<td bgcolor=""#CCCCCC"">" & vbCrLf

        code = code & "			<!-- Marker -->" & vbCrLf

        code = code & "			<% " & vbCrLf

        code = code & "				rowMessage = """"" & vbCrLf

        code = code & "				if defaultPosition = rs.absoluteposition then" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "					rowMessage = Message" & vbCrLf

        code = code & "					if len(rowMessage) > 0 then" & vbCrLf

        code = code & "						call showControls_rowMarker" & vbCrLf

        code = code & "					end if	" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "				end if" & vbCrLf

        code = code & "			%>" & vbCrLf

        code = code & "	</td>" & vbCrLf

        code = code & "	<td bgcolor=""#CCCCCC"">" & vbCrLf

        code = code & "			<!-- Row numbering -->" & vbCrLf

        code = code & "			<%" & vbCrLf

        code = code & "				call showControls_rowLabeling" & vbCrLf

        code = code & "			%>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	</td>" & vbCrLf


        For Fieldidx = 0 To Rs.Fields.Count - 1

            code = code & "      <td>" & vbCrLf

            code = code & "<div LANGUAGE=javascript onmouseover=""return record_onfocus(row<%=Rs.AbsolutePosition%>, form<%=Rs.AbsolutePosition%>)"">" & vbCrLf

            setReadOnly = ""

            fieldProperties = ""

            If Rs.Fields(Fieldidx).Properties.Item("ISAUTOINCREMENT").Value = True Then

                setReadOnly = " readonly = ""readonly"""

            End If

            Select Case Rs.Fields(Fieldidx).Type

                Case 203

                    fieldProperties = " maxlength=""" & Rs.Fields(Fieldidx).DefinedSize & """ " & setReadOnly

                Case Else

                    fieldProperties = " maxlength=""" & Rs.Fields(Fieldidx).DefinedSize & """ size =""" & Replace(CStr(Rs.Fields(Fieldidx).DefinedSize / 10), "0.", "") & """"

            End Select

            fieldProperties = Replace(fieldProperties, """", """""")

            code = code & " <%" & vbCrLf

            code = code & " fieldProperties = """ & fieldProperties & """" & vbCrLf

            code = code & " call showControls_" & Rs.Fields(Fieldidx).Name & "(Value(" & Fieldidx & "),  fieldProperties , controlEvents(Rs.AbsolutePosition," & Fieldidx & "))" & vbCrLf

            code = code & "%>" & vbCrLf

            code = code & "<font face=""Arial, Helvetica, sans-serif"" size=""2"" color=""red""><%=rowFieldMessage(" & Fieldidx & ")%></font>" & vbCrLf

            code = code & "</div>" & vbCrLf

            code = code & "      </td>" & vbCrLf

        Next

        code = code & "     <td>" & vbCrLf

        code = code & "<div LANGUAGE=javascript onmouseover=""return record_onfocus(row<%=Rs.AbsolutePosition%>,  form<%=Rs.AbsolutePosition%>)"">" & vbCrLf

        code = code & "                <input type=""hidden"" name = ""Pos""  value=""<%=Rs.AbsolutePosition%>"">" & vbCrLf

        code = code & "<input type=""hidden"" name=""formAction"" value=""Update"">" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "if not isSaved = false then" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	call showBtnDetails(Rs)" & vbCrLf

        code = code & "	call showBtnUpdate(Rs)" & vbCrLf

        code = code & "	call showBtnDelete(Rs)" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "else" & vbCrLf

        code = code & "	if defaultPosition = Rs.absolutePosition then" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		select case formAction" & vbCrLf

        code = code & "		" & vbCrLf

        code = code & "		case ""Update""" & vbCrLf

        code = code & "			call showBtnUpdate(Rs)" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "		case ""Delete""" & vbCrLf

        code = code & "			call showBtnDelete(Rs)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		end select" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		call showBtnCancel(Rs.AbsolutePosition)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	end if" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "end if" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "&nbsp" & vbCrLf

        code = code & "</div>" & vbCrLf

        code = code & "            </td>" & vbCrLf

        code = code & "    </form>" & vbCrLf

        code = code & "   " & vbCrLf

        code = code & "  </tr>" & vbCrLf

        code = code & "<%	" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "	call showRowMessage(rs.absoluteposition, rowMessage)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	for idx = 0 to " & Rs.Fields.Count - 1

        code = code & " 	" & vbCrLf

        code = code & "		rowFieldMessage(idx) = """"" & vbCrLf

        code = code & " 	" & vbCrLf

        code = code & "	next         " & vbCrLf

        code = code & "" & vbCrLf

        code = code & "End sub" & vbCrLf

        code = code & "%>" & vbCrLf


        Call WriteCodeFile(codeFileName, code)

    End Sub
    Private Sub BuildForm(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim code As String

        Dim codeFilename As String = ""

        codeFilename = Path & "\form.asp"

        code = "<!-- #include File=""forms\grid\rows\addNew.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\rows\Message.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\rows\Record.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\rows\columnHeader.asp"" -->" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\buttons\btnPreviousPage.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\buttons\btnNextPage.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\buttons\btnDelete.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\buttons\btnDetails.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\buttons\btnUpdate.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\buttons\btnUpdateNew.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\buttons\btnCancel.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\pageNavigation\pageNavigation.asp"" -->" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<!-- #include File=""forms\grid\toolbar.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""recordset\pageSize.asp"" -->" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "	Dim rowfieldMessage(" & Rs.Fields.Count - 1 & ")" & vbCrLf

        code = code & "	Dim savedValue(" & Rs.Fields.Count - 1 & ")" & vbCrLf

        code = code & "	Dim rowMessage	" & vbCrLf

        code = code & "	Dim rowLabel" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	Dim mode" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	sub showForm(Cn, Rs)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		for idx = 0 to " & Rs.Fields.Count - 1 & vbCrLf

        code = code & "			savedValue(idx) = Value(idx)" & vbCrLf

        code = code & "		next" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		mode = request(""mode"")" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		call RsPageSize" & vbCrLf

        code = code & "" & vbCrLf

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

        code = code & "		call showToolbar" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<table width=""100%"" border=""0"" cellspacing=""1""  bgcolor=""#666666"" cellpadding=""0"" height=""70%"">" & vbCrLf

        code = code & "<tr  valign=""top""  bgcolor=""#FFFFFF"">" & vbCrLf

        code = code & "<td>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" id=""dataTable"" class=""grid"">" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "   call showRowColumnHeader" & vbCrLf

        code = code & "  " & vbCrLf

        code = code & "  EndPosition = 0" & vbCrLf

        code = code & "  rowLabel = 1  " & vbCrLf

        code = code & "" & vbCrLf

        code = code & "  if Rs.RecordCount > 0 then" & vbCrLf

        code = code & "    " & vbCrLf

        code = code & "    for pageidx=1 to Rs.PageSize " & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "	Call showRowRecord(Rs)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	EndPosition = Rs.absoluteposition" & vbCrLf

        code = code & " 	  " & vbCrLf

        code = code & "	rs.movenext" & vbCrLf

        code = code & "	rowLabel = rowLabel + 1" & vbCrLf

        code = code & "                " & vbCrLf

        code = code & "                if Rs.EOF = true then" & vbCrLf

        code = code & "                    exit for" & vbCrLf

        code = code & "                end if" & vbCrLf

        code = code & "        " & vbCrLf

        code = code & "        next" & vbCrLf

        code = code & "        " & vbCrLf

        code = code & "        If mode = 2 then" & vbCrLf

        code = code & "            DefaultPosition = EndPosition" & vbCrLf

        code = code & "        end if" & vbCrLf

        code = code & "    " & vbCrLf

        code = code & "        Rs.Absolutepage = Absolutepage   " & vbCrLf

        code = code & "" & vbCrLf

        code = code & "    end if  " & vbCrLf

        code = code & "    " & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "<% call showRowAddNew(Cn, Rs)%>" & vbCrLf

        code = code & "</table>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "</td>" & vbCrLf

        code = code & "</tr>" & vbCrLf

        code = code & "</table>" & vbCrLf

        code = code & "<br>" & vbCrLf

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

        code = code & "<script language=""JavaScript"" src=""javascript\goDetails.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\cancelUpdate.js""></SCRIPT> " & vbCrLf

        code = code & "<script language=""JavaScript"" src=""javascript\addNewRecord.js""></SCRIPT>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<script language=""javascript"">" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "    var recordcount;" & vbCrLf

        code = code & "    var absolutepage;" & vbCrLf

        code = code & "    var pagecount;" & vbCrLf

        code = code & "    var pagesize; " & vbCrLf

        code = code & "    var startposition;" & vbCrLf

        code = code & "    var endposition;" & vbCrLf

        code = code & "    var fieldcount;" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "    var isloading;" & vbCrLf

        code = code & "    var isscrolling;" & vbCrLf

        code = code & "    " & vbCrLf

        code = code & "    var currentform;" & vbCrLf

        code = code & "    var currentrow;" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "    var currentColumnId;" & vbCrLf

        code = code & "    " & vbCrLf

        code = code & "    recordcount = <%=Rs.Recordcount%>;" & vbCrLf

        code = code & "    absolutepage = <%=AbsolutePage%>;" & vbCrLf

        code = code & "    pagecount = <%=Rs.pagecount%>;" & vbCrLf

        code = code & "    pagesize = <%=Rs.pagesize%>;" & vbCrLf

        code = code & "    startposition = <%=StartPosition%>;" & vbCrLf

        code = code & "    endposition = <%=EndPosition%>;" & vbCrLf

        code = code & "    fieldcount = <%=Rs.Fields.count%>;" & vbCrLf

        code = code & "    " & vbCrLf

        code = code & "    isloading = -1;" & vbCrLf

        code = code & "    isscrolling = -1;" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<% if defaultposition > 0 then%>" & vbCrLf

        code = code & "            " & vbCrLf

        code = code & "             currentform= form<%=DefaultPosition%>;" & vbCrLf

        code = code & "             currentrow= row<%=DefaultPosition%>;" & vbCrLf

        code = code & "             record_onfocus(row<%=DefaultPosition%>,form<%=DefaultPosition%>);" & vbCrLf

        code = code & "" & vbCrLf

        code = code & " if (document.forms(""form<%=DefaultPosition%>"").elements[0] != null)" & vbCrLf

        code = code & "{" & vbCrLf

        code = code & "     try" & vbCrLf

        code = code & " {" & vbCrLf

        code = code & "	 	document.forms(""form<%=DefaultPosition%>"").elements[0].focus();" & vbCrLf

        code = code & " }" & vbCrLf

        code = code & "	    catch(err)" & vbCrLf

        code = code & " {" & vbCrLf

        code = code & " }" & vbCrLf

        code = code & "}" & vbCrLf

        code = code & "            " & vbCrLf

        code = code & "        <%end if%>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%select case mode%>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%case 2%>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & " if (form<%=rs.recordcount+1%>.elements[0]  != null)" & vbCrLf

        code = code & "{" & vbCrLf

        code = code & "        form<%=rs.recordcount+1%>.elements[0].focus();" & vbCrLf

        code = code & "}" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%case else%>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "          document.frmfind.findstring.focus();" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%end select%>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "                                " & vbCrLf

        code = code & "        function record_onkeyup(recordid,txtfieldid) {" & vbCrLf

        code = code & "        " & vbCrLf

        code = code & "            isscrolling = -1;" & vbCrLf

        code = code & "        };" & vbCrLf

        code = code & "        " & vbCrLf

        code = code & "</script>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		call pageNavigation	" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	end sub" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "" & vbCrLf

        Call WriteCodeFile(codeFileName, code)

    End Sub
    Private Sub BuildAddnew(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim code As String
        Dim codeFilename As String = ""

        Dim fieldidx As Short

        Dim fieldProperties As String

        Dim setReadOnly As String

        codeFilename = Path & "\addnew.asp"

        code = "<%" & vbCrLf

        code = code & "sub showRowAddNew(Cn, Rs)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	Dim tempRs" & vbCrLf

        code = code & "	Dim tempSQL" & vbCrLf

        code = code & "	Dim rowId" & vbCrLf

        code = code & "	Dim fieldProperties" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	if not request(""formAction"") = ""updateNew"" and isSaved = false then" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		exit sub" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	end if" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	set tempRs = server.createObject(""adodb.recordset"")" & vbCrLf

        code = code & "	tempSQL = ""select top 1 * from ("" & Rs.source & "") tempTable""" & vbCrLf

        code = code & "	tempRs.open tempSQL, CN, 1, 3" & vbCrLf

        code = code & "	tempRs.addnew" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	Call setDefaultValue(tempRs)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	rowId = rs.recordcount + 1" & vbCrLf

        code = code & "" & vbCrLf

        code = code & " 	if defaultPosition > rs.recordcount and isSaved = false then" & vbCrLf

        code = code & " 		for idx = 0 to " & Rs.Fields.Count - 1 & vbCrLf

        code = code & " 			rowFieldMessage(idx) = fieldMessage(idx)" & vbCrLf

        code = code & " 			value(idx) = savedValue(idx)" & vbCrLf

        code = code & " 		next" & vbCrLf

        code = code & " 	end if" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "<tr id = row<%=rowid%> align=""left"" valign=""top""> " & vbCrLf

        code = code & "    " & vbCrLf

        code = code & "	<form name=""form<%=rowid%>""  method=""post"" action=""edit.asp"">" & vbCrLf

        code = code & "<td bgcolor=""#CCCCCC"">" & vbCrLf

        code = code & "			<!-- Marker -->" & vbCrLf

        code = code & "			<% " & vbCrLf

        code = code & "				rowMessage = """"" & vbCrLf

        code = code & "				if defaultPosition = rowId then" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "					rowMessage = Message" & vbCrLf

        code = code & "					if len(rowMessage) > 0 then" & vbCrLf

        code = code & "						call showControls_rowMarker" & vbCrLf

        code = code & "					end if	" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "				end if" & vbCrLf

        code = code & "			%>" & vbCrLf

        code = code & "	</td>" & vbCrLf

        code = code & "	<td bgcolor=""#CCCCCC"">" & vbCrLf

        code = code & "			<!-- Row numbering -->" & vbCrLf

        code = code & "			<%" & vbCrLf

        code = code & "				call showControls_rowLabeling" & vbCrLf

        code = code & "			%>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	</td>" & vbCrLf

        For fieldidx = 0 To Rs.Fields.Count - 1

            code = code & "      <td>" & vbCrLf

            code = code & "                    <!-- " & Rs.Fields(fieldidx).Name & "-->" & vbCrLf

            code = code & "<div LANGUAGE=javascript onmouseover=""return record_onfocus(row<%=rowId%>, form<%=rowId%>)"">" & vbCrLf

            setReadOnly = ""

            fieldProperties = ""

            If Rs.Fields(fieldidx).Properties.Item("ISAUTOINCREMENT").Value = True Then

                setReadOnly = " readonly = ""readonly"""

            End If

            Select Case Rs.Fields(fieldidx).Type

                Case 203

                    fieldProperties = " maxlength=""" & Rs.Fields(fieldidx).DefinedSize & """ " & setReadOnly

                Case Else

                    fieldProperties = " maxlength=""" & Rs.Fields(fieldidx).DefinedSize & """ size =""" & Replace(CStr(Rs.Fields(fieldidx).DefinedSize / 10), "0.", "") & """"

            End Select

            fieldProperties = Replace(fieldProperties, """", """""")

            code = code & " <% " & vbCrLf

            code = code & " fieldProperties = """ & fieldProperties & """" & vbCrLf

            code = code & " call showControls_" & Rs.Fields(fieldidx).Name & "(value(" & fieldidx & "), fieldProperties ,controlEvents(rowId," & fieldidx & "))" & vbCrLf

            code = code & " %>" & vbCrLf

            code = code & "<font face=""Arial, Helvetica, sans-serif"" size=""2"" color=""red""><%=rowFieldMessage(" & fieldidx & ")%></font>" & vbCrLf

            code = code & "</div>" & vbCrLf

            code = code & "      </td>" & vbCrLf

        Next

        code = code & "    <td>" & vbCrLf

        code = code & "<div LANGUAGE=javascript onmouseover=""return record_onfocus(row<%=rowid%>, form<%=rowid%>)"">" & vbCrLf

        code = code & "<% " & vbCrLf

        code = code & "	call showBtnupdateNew(Rs)" & vbCrLf

        code = code & "	" & vbCrLf

        code = code & "	if request(""formAction"") = ""updateNew"" and isSaved = false then" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "		call showBtnCancel(rowId)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	end if" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "</div>" & vbCrLf

        code = code & "    </td>" & vbCrLf

        code = code & "<input type=""hidden"" name=""formAction"" value=""updateNew"">" & vbCrLf

        code = code & "                <input type=""hidden"" name = ""Pos""  value=""<%=rowid%>"">" & vbCrLf

        code = code & "    </form>" & vbCrLf

        code = code & "  </tr>" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "	call showRowMessage(rowId, rowMessage)" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	for idx = 0 to " & Rs.Fields.Count - 1 & vbCrLf

        code = code & " 	" & vbCrLf

        code = code & "		rowFieldMessage(idx) = """"" & vbCrLf

        code = code & " 	" & vbCrLf

        code = code & "	next      " & vbCrLf

        code = code & "" & vbCrLf

        code = code & " tempRs.cancelUpdate" & vbCrLf

        code = code & " tempRs.close" & vbCrLf

        code = code & "end sub" & vbCrLf

        code = code & "%>" & vbCrLf

        Call WriteCodeFile(codeFileName, code)

    End Sub
End Module