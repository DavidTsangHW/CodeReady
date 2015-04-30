Module Bas_ASPSideMenu

    Public Sub BuildSideMenu(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim code As String
        Dim codeFilename As String

        Dim tableRs As New ADODB.Recordset
        Dim Cn As New ADODB.Connection

        codeFilename = Path & "\smenu.asp"

        Cn.Open(ConnectionString)

        tableRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        code = ""

        code = code & "<!-- #include File=""logo.asp"" -->" & vbCrLf

        code = code & "<!-- #include File=""database\accessRights\displayName.asp"" -->" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "	sideMenuWidth = ""20%""" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "	sub showSideMenu" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "%>" & vbCrLf

        code = code & "<% call showLogo%>" & vbCrLf

        code = code & "<br>" & vbCrLf

        code = code & "<font face=""arial"" size=""2"">Hello, <%=credential_displayName%></font>" & vbCrLf

        code = code & "<br>" & vbCrLf

        code = code & "<font face=""arial"" size=""3""><b>Menu</b></font>" & vbCrLf

        code = code & "<br>" & vbCrLf

        code = code & "<font face=""arial"" size=""2"">" & vbCrLf

        code = code & "<br>" & vbCrLf

        code = code & "<ol style=""margin-top:0;"" >" & vbCrLf


        If InStr(1, LCase(Cn.ConnectionString), "excel") > 0 Or InStr(1, LCase(Cn.ConnectionString), ".xls") > 0 Then

            Do Until tableRs.EOF
                code = code & "<li><a href=""../../" & tableRs("TABLE_NAME").Value & "/" & Fso.GetFolder(Path).Name & """ >" & PrintTableName(tableRs, tableRs("TABLE_NAME").Value) & "</a>" & vbCrLf
                tableRs.MoveNext()
            Loop

        Else

            Do Until tableRs.EOF
                'Q300948
                If tableRs.Fields("table_type").Value = "TABLE" Then
                    code = code & "<li><a href=""../../" & tableRs("TABLE_NAME").Value & "/" & Fso.GetFolder(Path).Name & """>" & PrintTableName(tableRs, tableRs("TABLE_NAME").Value) & "</a>" & vbCrLf
                End If

                tableRs.MoveNext()
            Loop

        End If


        code = code & "</ol>" & vbCrLf

        code = code & "</font>" & vbCrLf

        code = code & "" & vbCrLf

        code = code & "<%" & vbCrLf

        code = code & "	end sub" & vbCrLf

        code = code & "%>" & vbCrLf

        Call WriteCodeFile(codeFilename, code)

        tableRs.Close()

        Cn.Close()

    End Sub

End Module
