Module Bas_ASPPageHeader

    Public Sub BuildASPPageHeader(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFilename As String = ""
        Dim TableRs As ADODB.Recordset
        Dim Criteria As String
        Dim Cn As New ADODB.Connection

        Call Cn.Open(ConnectionString)

        TableRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        Criteria = "TABLE_NAME = '" & Rs.Fields(0).Properties(1).Value & "'"

        TableRs.Filter = Criteria

        codeFilename = Path & "\pageHeader.asp"

        Code = Code & "<!-- #include File=""fbar.asp"" -->" & vbCrLf
        Code = Code & "    <%" & vbCrLf
        Code = Code & "    Sub showPageHeader()" & vbCrLf
        Code = Code & "%>" & vbCrLf
        Code = Code & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
        Code = Code & "  <tr>" & vbCrLf
        Code = Code & "    <td><font face=""Arial, Helvetica, sans-serif"" size=""2""><b><font size=""5"">" & PrintTableName(TableRs, Rs.Fields(0).Properties(1).Value) & "</font></b></font><br>" & vbCrLf
        Code = Code & "    </td>" & vbCrLf
        Code = Code & "    <td valign=""top"" align=""right""  width=""40%""> " & vbCrLf
        Code = Code & "<%" & vbCrLf
        Code = Code & "Call showfunctionbar()" & vbCrLf
        Code = Code & "%>" & vbCrLf
        Code = Code & "    </td>" & vbCrLf
        Code = Code & "  </tr>" & vbCrLf
        Code = Code & "</table>" & vbCrLf
        Code = Code & "<%" & vbCrLf
        Code = Code & "	end sub" & vbCrLf
        Code = Code & "%>" & vbCrLf

        Call WriteCodeFile(codeFilename, Code)

        TableRs.Close()

        Cn.Close()

    End Sub

End Module
