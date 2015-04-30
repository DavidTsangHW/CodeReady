Option Strict Off
Option Explicit On
Module Bas_ASPPrint
    Private Const ModuleName As String = "Bas_ASPPrint"

    Private Sub BuildPrintRs(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFilename As String = ""

        Dim Fieldidx As Short
        Dim Criteria As String

        Dim Cn As New ADODB.Connection
        Dim ColumnRS As New ADODB.Recordset
        Dim TableRs As New ADODB.Recordset

        Cn.Open(ConnectionString)

        ColumnRS = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        TableRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        Criteria = "TABLE_NAME = '" & Rs.Fields(0).Properties(1).Value & "'"

        TableRs.Filter = Criteria

        codeFilename = Path & "\printRs.asp"

        Code = Code & " <!-- #include File=""pageHeader.asp"" -->" & vbCrLf
        Code = Code & " <!-- #include File=""smenu.asp"" -->" & vbCrLf
        Code = Code & " <!-- #include File=""setDefaultValue.asp"" -->" & vbCrLf

        Code = Code & "<%" & vbCrLf
        Code = Code & "sub printRs(objRs)" & vbCrLf
        Code = Code & "     Dim Rs" & vbCrLf
        Code = Code & "     set Rs = objRs.Clone" & vbCrLf
        Code = Code & "     Rs.sort = objRs.sort" & vbCrLf
        Code = Code & "     Rs.filter = objRs.filter" & vbCrLf
        Code = Code & "if Rs.recordCount > 0 then" & vbCrLf
        Code = Code & "         Rs.MoveFirst()" & vbCrLf
        Code = Code & "end if" & vbCrLf
        Code = Code & "%>" & vbCrLf

        Code = Code & "<html>" & vbCrLf
        Code = Code & "<head>" & vbCrLf

        Code = Code & "<title>" & PrintTableName(TableRs, Rs.Fields(0).Properties(1).Value) & "</title>" & vbCrLf
        Code = Code & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">" & vbCrLf
        Code = Code & "</head>" & vbCrLf

        Code = Code & "<body>"
        Code = Code & "<font face=""arial"">" & vbCrLf
        Code = Code & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
        Code = Code & "<tr>" & vbCrLf
        Code = Code & "<td width=""<%=sideMenuWidth%>"" valign=""top"" border=""0"" >" & vbCrLf

        Code = Code & "<%" & vbCrLf
        Code = Code & "	call showSideMenu" & vbCrLf
        Code = Code & "%>" & vbCrLf

        Code = Code & "</td>" & vbCrLf

        Code = Code & "<td valign=""top"">" & vbCrLf

        Code = Code & "<%" & vbCrLf
        Code = Code & "	showPageHeader" & vbCrLf
        Code = Code & "%>" & vbCrLf

        Code = Code & "<br>" & vbCrLf
        Code = Code & "<b>Print time:</b> <%=formatdatetime(now(), vbshortdate) & "" "" & formatdatetime(now(), vbshorttime)%>" & vbCrLf
        Code = Code & "<br>" & vbCrLf
        Code = Code & "<b>Record count:</b> <%=rs.recordcount%>" & vbCrLf
        Code = Code & "<table width=""100%"" border=""1"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
        Code = Code & "  <tr bgcolor=""#999999""> " & vbCrLf

        For Fieldidx = 0 To Rs.Fields.Count - 1

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(Fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            ColumnRS.Filter = Criteria

            Code = Code & "    <td >" & PrintColumnName(ColumnRS, Rs.Fields(Fieldidx).Name) & "&nbsp</td>" & vbCrLf

        Next

        Code = Code & "  </tr>" & vbCrLf
        Code = Code & "  <%" & vbCrLf
        Code = Code & "    Do Until Rs.eof" & vbCrLf
        Code = Code & "           call setDefaultValue(rs)" & vbCrLf
        Code = Code & "%>" & vbCrLf
        Code = Code & "  <tr> " & vbCrLf
        For Fieldidx = 0 To Rs.Fields.Count - 1

            Code = Code & "    <td><%=Value(" & Fieldidx & ")%>&nbsp</td>" & vbCrLf

        Next
        Code = Code & "  </tr>" & vbCrLf
        Code = Code & "  <%" & vbCrLf
        Code = Code & "        rs.movenext" & vbCrLf
        Code = Code & "    Loop" & vbCrLf


        Code = Code & "%>" & vbCrLf

        Code = Code & "</table>" & vbCrLf


        Code = Code & "</td>" & vbCrLf
        Code = Code & "</tr>" & vbCrLf
        Code = Code & "</table>" & vbCrLf
        Code = Code & "</font>" & vbCrLf
        Code = Code & "</body>" & vbCrLf
        Code = Code & "</html>" & vbCrLf

        Code = Code & "<%" & vbCrLf
        Code = Code & "	end sub" & vbCrLf
        Code = Code & "%>" & vbCrLf

        Call WriteCodeFile(codeFilename, Code)

        ColumnRS.Close()
        Cn.Close()


    End Sub

    Public Sub BuildASPPrint(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)


        Call BuildFramework(ConnectionString, Rs, Path)

        Call BuildPrintRs(ConnectionString, Rs, Path)

        Call CopyDirectory(AppPath("ASP Libraries\print"), Path)

    End Sub

End Module