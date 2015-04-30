Module Bas_ASPExcelImport
    Private Const ModuleName As String = "Bas_ASPExcelImport"

    Public Sub BuildASPImportForm(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Call BuildFramework(ConnectionString, Rs, Path)

        Call CopyDirectory(AppPath("\ASP Libraries\import"), Path)

        If Fso.FolderExists(Path & "\Temp") = False Then
            Call CreatePath(Path & "\Temp")
        End If

        Call BuildUploadsDirVar(ConnectionString, Rs, Path)
        Call BuildTableMapping(ConnectionString, Rs, Path)

    End Sub

    Private Sub BuildTableMapping(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)


        Dim Code As String = ""
        Dim codeFilename As String = ""

        codeFilename = Path & "\tableMapping.asp"

        Code = Code & "<!-- #include File=""readSchema.asp"" -->" & vbCrLf
        Code = Code & "<!-- #include File=""database\database.asp"" -->" & vbCrLf
        Code = Code & "<!-- #include File=""excel.asp"" -->" & vbCrLf
        Code = Code & "<!-- #include File=""pageHeader.asp"" -->" & vbCrLf
        Code = Code & "<!-- #include File=""smenu.asp"" -->" & vbCrLf
        Code = Code & "<!-- #include file=""uploadsDirVar.asp"" -->" & vbCrLf

        Code = Code & "<%	" & vbCrLf

        Code = Code & "        Dim excelCN" & vbCrLf
        Code = Code & "        Dim Rs" & vbCrLf


        Code = Code & "        set excelCN = server.CreateObject(""adodb.connection"")" & vbCrLf

        Code = Code & "        set excelCN = openExcelCN(uploadsDirVar & ""\"" & request(""attach1""))" & vbCrLf

        Code = Code & "        If excelCN.state = 0 Then" & vbCrLf
        Code = Code & "            response.end()" & vbCrLf
        Code = Code & "        End If" & vbCrLf

        Code = Code & "If IsObject(Session(rsName)) = False Then" & vbCrLf

        Code = Code & "Call buildDataSession()" & vbCrLf

        Code = Code & "End If" & vbCrLf

        Code = Code & "        set Rs = Session(rsName)" & vbCrLf

        Code = Code & "%>" & vbCrLf

        Code = Code & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
        Code = Code & "<tr>" & vbCrLf
        Code = Code & "<td width=""<%=sideMenuWidth%>"" valign=""top"" border=""0"" >" & vbCrLf
        Code = Code & "<%" & vbCrLf
        Code = Code & "        Call showSideMenu()" & vbCrLf
        Code = Code & "%>" & vbCrLf
        Code = Code & "</td>" & vbCrLf
        Code = Code & "<td valign=""top"" >" & vbCrLf
        Code = Code & "<%" & vbCrLf
        Code = Code & "        Call showPageHeader()" & vbCrLf
        Code = Code & "%>" & vbCrLf


        Code = Code & "<font face=""arial"">" & vbCrLf

        Code = Code & "<form name=""tableMapping"" method=""post"" action=""import.asp"">" & vbCrLf

        Code = Code & "<table width=""100%"" height=""100"">" & vbCrLf

        Code = Code & "	<tr>" & vbCrLf
        Code = Code & "		<td>" & vbCrLf
        Code = Code & "        Filename()" & vbCrLf
        Code = Code & "		</td>" & vbCrLf
        Code = Code & "		<td>" & vbCrLf
        Code = Code & "        Table()" & vbCrLf
        Code = Code & "		</td>" & vbCrLf
        Code = Code & "	</tr>" & vbCrLf
        Code = Code & "	<tr>" & vbCrLf

        Code = Code & "		<td>" & vbCrLf
        Code = Code & "			<input type=""hidden"" name=""attach1"" value=""<%=request(""attach1"")%>""><%=request(""attach1"")%>" & vbCrLf
        Code = Code & "		</td>" & vbCrLf
        Code = Code & "		<td>" & vbCrLf
        Code = Code & "			<%call showTableSelection(excelCN, ""cboExcel"")%><input type=""hidden"" name=""cbomdb"" value=""" & Rs.Fields(0).Properties(1).Value & """>" & vbCrLf
        Code = Code & "		</td>" & vbCrLf
        Code = Code & "	</tr>" & vbCrLf

        Code = Code & "</table>" & vbCrLf
        Code = Code & "<p>" & vbCrLf

        Code = Code & "	<input type=""submit"" value=""Proceed"">" & vbCrLf
        Code = Code & "<input type=""button"" value=""Cancel"" onClick=""javascript:window.navigate('./')"">" & vbCrLf

        Code = Code & "</p>" & vbCrLf
        Code = Code & "</form>" & vbCrLf

        Code = Code & "</font>" & vbCrLf


        Code = Code & "</td>" & vbCrLf
        Code = Code & "</tr>" & vbCrLf
        Code = Code & "</table>" & vbCrLf

        Code = Code & "<%" & vbCrLf
        Code = Code & "excelCN.Close()" & vbCrLf
        Code = Code & "%>" & vbCrLf

        WriteCodeFile(codeFilename, Code)

    End Sub
    Private Sub BuildUploadsDirVar(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFilename As String = ""

        codeFilename = Path & "\uploadsDirVar.asp"



        Code = "<%" & vbCrLf
        Code = Code & "" & vbCrLf
        Code = Code & "' ****************************************************" & vbCrLf
        Code = Code & "' Change the value of the variable below to the pathname" & vbCrLf
        Code = Code & "' of a directory with write permissions, for example ""C:\Inetpub\wwwroot""" & vbCrLf
        Code = Code & "  Dim uploadsDirVar" & vbCrLf
        Code = Code & "  uploadsDirVar = """ & Path & "\temp""" & vbCrLf
        Code = Code & "' ****************************************************" & vbCrLf
        Code = Code & "" & vbCrLf
        Code = Code & "%>" & vbCrLf

        WriteCodeFile(codeFilename, Code)

    End Sub

End Module
