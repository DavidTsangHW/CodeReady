Option Strict Off
Option Explicit On
Module Bas_ASPProject
	Private Const ModuleName As String = "Bas_ASPProject"
    Public Function CreateASPForms(ByVal InetpubPath As String, ByVal SchemaRs As ADODB.Recordset, ByVal ConnectionString As String, ByVal formType As String) As Boolean

        Dim Cn As New ADODB.Connection
        Dim Rs As New ADODB.Recordset

        Dim SQL As String
        Dim Path As String

        On Error GoTo ErrorHandler

        Cn.Open(ConnectionString)

        SchemaRs.MoveFirst()

        Call Build_AccessRights(ConnectionString, Rs, InetpubPath)

        Do Until SchemaRs.EOF

            If (InStr(1, LCase(ConnectionString), "excel") > 0 Or InStr(1, LCase(ConnectionString), ".xls") > 0) Or SchemaRs.Fields("table_type").Value = "TABLE" Then

                Path = InetpubPath & "\" & LCase(SchemaRs.Fields("table_name").Value)

                If Fso.FolderExists(Path) = False Then

                    Call CreatePath(Path)

                End If

                Path = Path & "\" & formType

                If Fso.FolderExists(Path) = False Then

                    Call CreatePath(Path)

                End If

                SQL = "select * from [" & SchemaRs.Fields("table_name").Value & "]"

                Rs.Open(SQL, Cn, 1, 1)

                Select Case LCase(formType)

                    Case "blankform"
                        Call BuildASPBlankForm(ConnectionString, Rs, Path & "\default.asp")
                        Call BuildSideMenu(ConnectionString, Rs, Path)

                    Case "grid"
                        Call BuildASPGrid(ConnectionString, Rs, Path)
                        Call BuildSideMenu(ConnectionString, Rs, Path)


                    Case "editform"
                        Call BuildASPEditForm(ConnectionString, Rs, Path)
                        Call BuildSideMenu(ConnectionString, Rs, Path)

                    Case "print"
                        Call BuildASPPrint(ConnectionString, Rs, Path)
                        Call BuildSideMenu(ConnectionString, Rs, Path)

                    Case "import"
                        Call BuildASPImportForm(ConnectionString, Rs, Path)
                        Call BuildSideMenu(ConnectionString, Rs, Path)

                    Case "export"
                        Call BuildASPExportForm(ConnectionString, Rs, Path)

                End Select

                Rs.Close()

            End If


            SchemaRs.MoveNext()

        Loop

        BuildASPFormFrame(InetpubPath, SchemaRs, ConnectionString, formType)

        CreateASPForms = True

        Cn.Close()

        'UPGRADE_NOTE: Object Cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Cn = Nothing

        Exit Function

ErrorHandler:

        MsgBox(Err.Description, MsgBoxStyle.Critical)

        CreateASPForms = False

    End Function


    Public Sub BuildASPFormFrame(ByVal InetpubPath As String, ByVal SchemaRs As ADODB.Recordset, ByVal ConnectionString As String, ByVal formType As String)


        Dim Code As String = ""
        Dim codeFilename As String = ""
        Dim Path As String
        Dim Filename As String

        On Error GoTo ErrorHandler

        Path = InetpubPath & "\" & formType
        Filename = "default.asp"

        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If

        codeFilename = Path & "\" & Filename

        Code = "<html>" & vbCrLf
        Code = Code & "<head>" & vbCrLf
        Code = Code & "<title>" & vbCrLf
        Code = Code & "</title>" & vbCrLf
        Code = Code & "</head>" & vbCrLf
        Code = Code & "<body>" & vbCrLf
        Code = Code & "<font face=""arial"">" & vbCrLf
        Code = Code & "<h2>Welcome</h2>" & vbCrLf
        Code = Code & "<table width=""100%"" height=""600"" border=""0"">" & vbCrLf
        Code = Code & "<tr>" & vbCrLf
        Code = Code & "<td width=""20%"" valign=""top"">" & vbCrLf
        Code = Code & "<ol>" & vbCrLf

        SchemaRs.MoveFirst()

        If InStr(1, LCase(ConnectionString), "excel") > 0 Or InStr(1, LCase(ConnectionString), ".xls") > 0 Then

            Do Until SchemaRs.EOF

                Code = Code & "<li><a href=""../" & SchemaRs.Fields("table_name").Value & "/" & formType & """ >" & SchemaRs.Fields("table_name").Value & "</a>" & vbCrLf
                SchemaRs.MoveNext()

            Loop

        Else

            Do Until SchemaRs.EOF

                'Q300948
                If SchemaRs.Fields("table_type").Value = "TABLE" Then

                    Code = Code & "<li><a href=""../" & SchemaRs.Fields("table_name").Value & "/" & formType & """ >" & SchemaRs.Fields("table_name").Value & "</a>" & vbCrLf

                End If

                SchemaRs.MoveNext()

            Loop

        End If

        Code = Code & "</ol>" & vbCrLf
        Code = Code & "</td>" & vbCrLf
        Code = Code & "<td>" & vbCrLf
        Code = Code & "</td>" & vbCrLf
        Code = Code & "</tr>" & vbCrLf
        Code = Code & "</table>" & vbCrLf
        Code = Code & "</font>" & vbCrLf
        Code = Code & "</body>" & vbCrLf
        Code = Code & "</html>" & vbCrLf

        Call WriteCodeFile(codeFilename, Code)



        Exit Sub

ErrorHandler:

        MsgBox(Err.Description, MsgBoxStyle.Critical)

    End Sub

   
End Module