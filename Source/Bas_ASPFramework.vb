Module Bas_ASPFramework


    Public Sub BuildFramework(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Call CopyDirectory(AppPath("ASP Libraries\framework"), Path)

        Call BuildsetDefaultValue(ConnectionString, Rs, Path)
        Call BuildPageTitle(ConnectionString, Rs, Path)
        Call BuildConnectionString(ConnectionString, Rs, Path & "\database")
        Call BuildrsSource(ConnectionString, Rs, Path & "\database")
        Call BuildRsSource_RecordByUserId(ConnectionString, Rs, Path & "\database\recordByUserId")
        Call BuildRsSource_RecordByUserId(ConnectionString, Rs, Path & "\database\recordByUserId\resources\fields")
        Call BuildrsName(ConnectionString, Rs, Path & "\database")
        Call BuildcnName(ConnectionString, Rs, Path & "\database")

        Call BuildEndecrypt_Fields(ConnectionString, Rs, Path & "\endecrypt\resources\endecrypt")
        Call BuildEndecrypt_FieldsAbstract(ConnectionString, Rs, Path & "\endecrypt\resources\abstract")
        Call BuildEndecrypt_FieldsAbstract(ConnectionString, Rs, Path & "\endecrypt")

    End Sub

    Public Sub BuildControlLabels(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Call BuildControlLabel_Text(ConnectionString, Rs, Path & "\controlLabels\resources\text")
        Call BuildControlLabel_LinkToMainTable(ConnectionString, Rs, Path & "\controlLabels\resources\linktomaintable")

        Call BuildControlLabel_Abstract(ConnectionString, Rs, Path & "\controlLabels\resources\abstract")
        Call BuildControlLabel_UserId(ConnectionString, Rs, Path & "\controlLabels\resources\userId")

        Call BuildControlLabel_Text(ConnectionString, Rs, Path & "\controlLabels")
        Call BuildControlLabel_LinkToMainTable(ConnectionString, Rs, Path & "\controlLabels")
        Call BuildControlLabel_UserId(ConnectionString, Rs, Path & "\controlLabels")

    End Sub

    Public Sub BuildControls(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Call BuildControls_Textbox(ConnectionString, Rs, Path & "\controls\resources\textbox")
        Call BuildControls_Hyperlink(ConnectionString, Rs, Path & "\controls\resources\hyperlink")
        Call BuildControls_Text(ConnectionString, Rs, Path & "\controls\resources\textonly")
        Call BuildControls_Pulldown(ConnectionString, Rs, Path & "\controls\resources\pulldown")
        Call BuildControls_RadioButton(ConnectionString, Rs, Path & "\controls\resources\radiobutton")
        Call BuildControls_Abstract(ConnectionString, Rs, Path & "\controls\resources\abstract")
        Call BuildControls_Calendar(ConnectionString, Rs, Path & "\controls\resources\calendar")
        Call BuildControls_UserId(ConnectionString, Rs, Path & "\controls\resources\userID")

        Call BuildControls_Textbox(ConnectionString, Rs, Path & "\controls")
        Call BuildControls_RichTextBox(ConnectionString, Rs, Path & "\controls")
        Call BuildControls_Pulldown(ConnectionString, Rs, Path & "\controls")
        Call BuildControls_RadioButton(ConnectionString, Rs, Path & "\controls")
        Call BuildControls_Calendar(ConnectionString, Rs, Path & "\controls")
        Call BuildControls_UserId(ConnectionString, Rs, Path & "\controls")

    End Sub

    Public Sub BuildControls_Pulldown(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFileName As String = ""
        Dim Cn As New ADODB.Connection

        Dim FieldRelationRs As New ADODB.Recordset
        Dim ColumnRS As New ADODB.Recordset

        Dim Fieldidx As Short

        Dim Criteria As String

        Cn.Open(ConnectionString)

        FieldRelationRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaForeignKeys)
        ColumnRS = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        For Fieldidx = 0 To Rs.Fields.Count - 1

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(Fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            ColumnRS.Filter = Criteria

            'If it is a foreign key
            '12 Oct 2009
            Criteria = "FK_TABLE_NAME= '" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            Criteria = Criteria & " AND FK_COLUMN_NAME= '" & Rs.Fields(Fieldidx).Name & "'"

            FieldRelationRs.Filter = ""
            FieldRelationRs.Filter = Criteria

            Select Case False

                Case FieldRelationRs.EOF

                    codeFileName = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

                    Code = "<%" & vbCrLf

                    Code = Code & "sub showControls_" & Rs.Fields(Fieldidx).Name & "(defaultValue, properties, javascript)" & vbCrLf

                    Code = Code & " Dim lookupRs" & vbCrLf

                    Code = Code & " set lookupRs = server.createObject(""adodb.recordset"")" & vbCrLf

                    Code = Code & "%>" & vbCrLf

                    Code = Code & "<select name=""txtfield_" & Rs.Fields(Fieldidx).Name & """>" & vbCrLf

                    Code = Code & "<option value=""""> -- Select -- </option>" & vbCrLf

                    Do Until FieldRelationRs.EOF

                        Code = Code & "<%" & vbCrLf
                        Code = Code & "SQL = ""select " & FieldRelationRs.Fields("PK_COLUMN_NAME").Value & " from [" & FieldRelationRs.Fields("PK_TABLE_NAME").Value & "]""" & vbCrLf
                        Code = Code & "%>" & vbCrLf

                        Code = Code & "<%"
                        Code = Code & "        lookupRs.Open SQL,CN,1,1" & vbCrLf
                        Code = Code & "        Do Until lookupRs.EOF" & vbCrLf
                        Code = Code & "            Response.Write ""<option value="""""" & lookupRs(0) & """"""""" & vbCrLf

                        Code = Code & " if lookupRs(0) = defaultValue then" & vbCrLf
                        Code = Code & "     Response.Write "" selected"" " & vbCrLf
                        Code = Code & " end if" & vbCrLf
                        Code = Code & "" & vbCrLf

                        Code = Code & "  Response.Write "">"" & lookupRs(0)" & " & ""</option>""" & vbCrLf

                        Code = Code & "            lookupRs.MoveNext" & vbCrLf
                        Code = Code & "        Loop" & vbCrLf
                        Code = Code & "        lookupRs.Close" & vbCrLf
                        Code = Code & "%>" & vbCrLf

                        FieldRelationRs.MoveNext()

                    Loop

                    Code = Code & "        </select>" & vbCrLf

                    Code = Code & "<%" & vbCrLf

                    Code = Code & " End sub" & vbCrLf

                    Code = Code & "%>" & vbCrLf

                    WriteCodeFile(codeFileName, Code)


            End Select

        Next


        FieldRelationRs.Close()

        ColumnRS.Close()

        Cn.Close()

    End Sub
    Public Sub BuildFrmFind(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFileName As String = ""
        Dim FieldIdx As Short
        Dim Criteria As String

        Dim ColumnRS As New ADODB.Recordset
        Dim CN As New ADODB.Connection

        CN.Open(ConnectionString)

        ColumnRS = CN.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        codeFileName = Path & "\frmFind.asp"

        Code = "<!-- #include File=""buttons\btnFind.asp"" -->" & vbCrLf

        Code = Code & "<!-- #include File=""buttons\btnReload.asp"" -->" & vbCrLf

        Code = Code & "<%" & vbCrLf

        Code = Code & "sub showFrmFind(Rs)" & vbCrLf

        Code = Code & "	" & vbCrLf

        Code = Code & "	dim criteria" & vbCrLf

        Code = Code & "	dim fieldName" & vbCrLf

        Code = Code & "	dim fieldIdx" & vbCrLf

        Code = Code & "	dim idx" & vbCrLf

        Code = Code & "	" & vbCrLf

        Code = Code & "	if not cstr(Rs.filter) = ""0"" then" & vbCrLf

        Code = Code & "			" & vbCrLf

        Code = Code & "		criteria = Rs.filter" & vbCrLf

        Code = Code & "	" & vbCrLf

        Code = Code & "		fieldname =  mid(criteria, 1,  instr(1, criteria, "" "")-1)" & vbCrLf

        Code = Code & "		" & vbCrLf

        Code = Code & "		criteria = mid(criteria, instr(1, criteria, ""like"") + 5)" & vbCrLf

        Code = Code & "		" & vbCrLf

        Code = Code & "		criteria = replace(criteria, ""'"", """")" & vbCrLf

        Code = Code & "		" & vbCrLf

        Code = Code & "		for idx = 0 to rs.fields.count -1 " & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "			if lcase(fieldname) = lcase(rs(idx).name) then" & vbCrLf

        Code = Code & "				" & vbCrLf

        Code = Code & "				fieldidx = idx" & vbCrLf

        Code = Code & "				exit for" & vbCrLf

        Code = Code & "				" & vbCrLf

        Code = Code & "			end if		" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "		next" & vbCrLf

        Code = Code & "		" & vbCrLf

        Code = Code & "	end if" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "%>" & vbCrLf

        Code = Code & " <form name=""frmfind"" method=""post"" action=""edit.asp"" style=""display: inline; margin: 0;"">" & vbCrLf

        Code = Code & "        <p><font face=""Arial, Helvetica, sans-serif"" size=""2"">Find: " & vbCrLf

        Code = Code & "          <input type=""text"" name=""findstring"" value=""<%=criteria%>"" Language=JAVASCRIPT onkeydown=""return record_onkeydown(<%=startposition-1%>,0)"">" & vbCrLf

        Code = Code & "          in" & vbCrLf

        Code = Code & "<select name=""fieldidx"">" & vbCrLf

        For FieldIdx = 0 To Rs.Fields.Count - 1

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(FieldIdx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(FieldIdx).Properties(1).Value & "'"

            ColumnRS.Filter = Criteria

            Code = Code & "<option value=""" & FieldIdx & """" & vbCrLf

            Code = Code & "<% if fieldidx = " & FieldIdx & " then%>" & vbCrLf

            Code = Code & "	selected" & vbCrLf

            Code = Code & "<%end if%>" & vbCrLf

            Code = Code & ">" & PrintColumnName(ColumnRS, Rs.Fields(FieldIdx).Name) & "</option>" & vbCrLf

        Next

        Code = Code & "</select>" & vbCrLf

        Code = Code & "          <%call showBtnFind()%>" & vbCrLf

        Code = Code & "          <%call showBtnReload()%>" & vbCrLf

        Code = Code & "         </font>" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "<input type=""hidden"" name=""formAction"" value=""Find"">" & vbCrLf

        Code = Code & "</form>" & vbCrLf

        Code = Code & "<%" & vbCrLf

        Code = Code & " end sub" & vbCrLf

        Code = Code & "%>" & vbCrLf

        WriteCodeFile(codeFileName, Code)

        ColumnRS.Close()

        CN.Close()


    End Sub

    Private Sub BuildsetDefaultValue(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFileName As String = ""
        Dim FieldIdx As Short

        Dim ColumnRS As New ADODB.Recordset
        Dim CN As New ADODB.Connection

        CN.Open(ConnectionString)

        ColumnRS = CN.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        codeFileName = Path & "\setDefaultValue.asp"

        For FieldIdx = 0 To Rs.Fields.Count - 1

            Code = Code & "<!-- #include File=""endecrypt\endecrypt_" & Rs.Fields(FieldIdx).Name & ".asp"" -->" & vbCrLf

        Next

        Code = Code & " <%" & vbCrLf

        Code = Code & "Dim value(" & Rs.Fields.Count - 1 & ")" & vbCrLf

        Code = Code & " Sub setDefaultValue(Rs)" & vbCrLf

        Code = Code & " " & vbCrLf

        Code = Code & " 	If Rs.EOF = False And Rs.BOF = False Then" & vbCrLf

        For FieldIdx = 0 To Rs.Fields.Count - 1

            Code = Code & " 		value(" & FieldIdx & ") = Decrypt_" & Rs.Fields(FieldIdx).Name & "(Rs(""" & Rs(FieldIdx).Name & """))" & vbCrLf

        Next
        Code = Code & " 	End if" & vbCrLf

        Code = Code & " " & vbCrLf

        Code = Code & " End Sub" & vbCrLf

        Code = Code & " %>" & vbCrLf

        WriteCodeFile(codeFileName, Code)

        ColumnRS.Close()

        CN.Close()

    End Sub

    Private Sub BuildControls_Calendar(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFilename As String = ""
        Dim Fieldidx As Short
        Dim columnRs As New ADODB.Recordset
        Dim Cn As New ADODB.Connection
        Dim Criteria As String

        Cn.Open(ConnectionString)

        columnRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If


        For Fieldidx = 0 To Rs.Fields.Count - 1

            Select Case Rs.Fields(Fieldidx).Type

                Case ADODB.DataTypeEnum.adDate, ADODB.DataTypeEnum.adDBDate, ADODB.DataTypeEnum.adDBDate

                    'Search description of field
                    '25 August 2010
                    Criteria = "COLUMN_NAME = '" & Rs.Fields(Fieldidx).Name & "'"
                    Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

                    columnRs.Filter = Criteria

                    codeFilename = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

                    Code = "<%" & vbCrLf

                    Code = Code & "sub showControls_" & Rs.Fields(Fieldidx).Name & "(defaultValue, properties, javascript)" & vbCrLf

                    Code = Code & "	dim defaultValue_day" & vbCrLf

                    Code = Code & "	dim defaultValue_month" & vbCrLf

                    Code = Code & "	dim defaultValue_year" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "	if isDate(defaultValue) = true then" & vbCrLf

                    Code = Code & "		defaultValue_day = day(defaultValue)" & vbCrLf

                    Code = Code & "		defaultValue_month = month(defaultValue)" & vbCrLf

                    Code = Code & "		defaultValue_year = year(defaultValue)" & vbCrLf

                    Code = Code & "	end if" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "%>" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "<link rel=""stylesheet"" type=""text/css"" href=""../../yui/build/fonts/fonts-min.css"" />" & vbCrLf

                    Code = Code & "<link rel=""stylesheet"" type=""text/css"" href=""../../yui/build/button/assets/skins/sam/button.css"" />" & vbCrLf

                    Code = Code & "<link rel=""stylesheet"" type=""text/css"" href=""../../yui/build/container/assets/skins/sam/container.css"" />" & vbCrLf

                    Code = Code & "<link rel=""stylesheet"" type=""text/css"" href=""../../yui/build/calendar/assets/skins/sam/calendar.css"" />" & vbCrLf

                    Code = Code & "<script type=""text/javascript"" src=""../../yui/build/yahoo-dom-event/yahoo-dom-event.js""></script>" & vbCrLf

                    Code = Code & "<script type=""text/javascript"" src=""../../yui/build/dragdrop/dragdrop-min.js""></script>" & vbCrLf

                    Code = Code & "<script type=""text/javascript"" src=""../../yui/build/element/element-min.js""></script>" & vbCrLf

                    Code = Code & "<script type=""text/javascript"" src=""../../yui/build/button/button-min.js""></script>" & vbCrLf

                    Code = Code & "<script type=""text/javascript"" src=""../../yui/build/container/container-min.js""></script>" & vbCrLf

                    Code = Code & "<script type=""text/javascript"" src=""../../yui/build/calendar/calendar-min.js""></script>" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "<style type=""text/css"">" & vbCrLf

                    Code = Code & "    /* Clear calendar's float, using dialog inbuilt form element */" & vbCrLf

                    Code = Code & "    #" & controlName(Rs, Fieldidx) & "container .bd form {" & vbCrLf

                    Code = Code & "        clear:left;" & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "    /* Have calendar squeeze upto bd bounding box */" & vbCrLf

                    Code = Code & "    #" & controlName(Rs, Fieldidx) & "container .bd {" & vbCrLf

                    Code = Code & "        padding:0;" & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "    #" & controlName(Rs, Fieldidx) & "container .hd {" & vbCrLf

                    Code = Code & "        text-align:left;" & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "    /* Center buttons in the footer */" & vbCrLf

                    Code = Code & "    #" & controlName(Rs, Fieldidx) & "container .ft .button-group {" & vbCrLf

                    Code = Code & "        text-align:center;" & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "    /* Prevent border-collapse:collapse from bleeding through in IE6, IE7 */" & vbCrLf

                    Code = Code & "    #" & controlName(Rs, Fieldidx) & "container_c.yui-overlay-hidden table {" & vbCrLf

                    Code = Code & "        *display:none;" & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "    /* Remove calendar's border and set padding in ems instead of px, so we can specify an width in ems for the container */" & vbCrLf

                    Code = Code & "    #cal {" & vbCrLf

                    Code = Code & "        border:none;" & vbCrLf

                    Code = Code & "        padding:1em;" & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "    /* Datefield look/feel */" & vbCrLf

                    Code = Code & "    .datefield {" & vbCrLf

                    Code = Code & "        position:relative;" & vbCrLf

                    Code = Code & "        white-space:nowrap;" & vbCrLf

                    Code = Code & "        border:0px solid black;" & vbCrLf

                    Code = Code & "        background-color:#fff;" & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "    .datefield input," & vbCrLf

                    Code = Code & "    .datefield button," & vbCrLf

                    Code = Code & "    .datefield label  {" & vbCrLf

                    Code = Code & "        vertical-align:middle;" & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "    #" & controlName(Rs, Fieldidx) & "month_field," & vbCrLf

                    Code = Code & "    #" & controlName(Rs, Fieldidx) & "date_field {" & vbCrLf

                    Code = Code & "    " & vbCrLf

                    Code = Code & "        width: 2em;" & vbCrLf

                    Code = Code & "    " & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "    " & vbCrLf

                    Code = Code & "    #" & controlName(Rs, Fieldidx) & "year_field {" & vbCrLf

                    Code = Code & "    " & vbCrLf

                    Code = Code & "        width: 3em;" & vbCrLf

                    Code = Code & "    " & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "    .datefield input  {" & vbCrLf

                    Code = Code & "        width:15em;" & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "    .datefield button  {" & vbCrLf

                    Code = Code & "        padding:0 5px 0 5px;" & vbCrLf

                    Code = Code & "        margin-left:2px;" & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "    .datefield button img {" & vbCrLf

                    Code = Code & "        padding:0;" & vbCrLf

                    Code = Code & "        margin:0;" & vbCrLf

                    Code = Code & "        vertical-align:middle;" & vbCrLf

                    Code = Code & "    }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "</style>" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "<script type=""text/javascript"">" & vbCrLf

                    Code = Code & "    YAHOO.util.Event.onDOMReady(function(){" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "        var Event = YAHOO.util.Event," & vbCrLf

                    Code = Code & "            Dom = YAHOO.util.Dom," & vbCrLf

                    Code = Code & "            " & controlName(Rs, Fieldidx) & "dialog," & vbCrLf

                    Code = Code & "            " & controlName(Rs, Fieldidx) & "calendar;" & vbCrLf

                    Code = Code & "	" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "	   Event.on( Dom.get(""" & controlName(Rs, Fieldidx) & "date_field""), ""change"", function() {" & vbCrLf

                    Code = Code & "			Dom.get(""" & controlName(Rs, Fieldidx) & """).value  = Dom.get(""" & controlName(Rs, Fieldidx) & "date_field"").value + '/' + Dom.get(""" & controlName(Rs, Fieldidx) & "month_field"").value + '/' + Dom.get(""" & controlName(Rs, Fieldidx) & "year_field"").value" & vbCrLf

                    Code = Code & "		});" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "	Event.on( Dom.get(""" & controlName(Rs, Fieldidx) & "month_field""), ""change"", function() {" & vbCrLf

                    Code = Code & "			Dom.get(""" & controlName(Rs, Fieldidx) & """).value  = Dom.get(""" & controlName(Rs, Fieldidx) & "date_field"").value + '/' + Dom.get(""" & controlName(Rs, Fieldidx) & "month_field"").value + '/' + Dom.get(""" & controlName(Rs, Fieldidx) & "year_field"").value" & vbCrLf

                    Code = Code & "		});" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "	Event.on( Dom.get(""" & controlName(Rs, Fieldidx) & "year_field""), ""change"", function() {" & vbCrLf

                    Code = Code & "			Dom.get(""" & controlName(Rs, Fieldidx) & """).value  = Dom.get(""" & controlName(Rs, Fieldidx) & "date_field"").value + '/' + Dom.get(""" & controlName(Rs, Fieldidx) & "month_field"").value + '/' + Dom.get(""" & controlName(Rs, Fieldidx) & "year_field"").value" & vbCrLf

                    Code = Code & "		});" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "        var " & controlName(Rs, Fieldidx) & "showBtn = Dom.get(""" & controlName(Rs, Fieldidx) & "show"");" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "        Event.on(" & controlName(Rs, Fieldidx) & "showBtn, ""click"", function() {" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "            // Lazy Dialog Creation - Wait to create the Dialog, and setup document click listeners, until the first time the button is clicked." & vbCrLf

                    Code = Code & "            if (!" & controlName(Rs, Fieldidx) & "dialog) {" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "                // Hide Calendar if we click anywhere in the document other than the calendar" & vbCrLf

                    Code = Code & "                Event.on(document, ""click"", function(e) {" & vbCrLf

                    Code = Code & "                    var el = Event.getTarget(e);" & vbCrLf

                    Code = Code & "                    var dialogEl = " & controlName(Rs, Fieldidx) & "dialog.element;" & vbCrLf

                    Code = Code & "                    if (el != dialogEl && !Dom.isAncestor(dialogEl, el) && el != " & controlName(Rs, Fieldidx) & "showBtn && !Dom.isAncestor(" & controlName(Rs, Fieldidx) & "showBtn, el)) {" & vbCrLf

                    Code = Code & "                        " & controlName(Rs, Fieldidx) & "dialog.hide();" & vbCrLf

                    Code = Code & "                    }" & vbCrLf

                    Code = Code & "                });" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "                function resetHandler() {" & vbCrLf

                    Code = Code & "                    // Reset the current calendar page to the select date, or " & vbCrLf

                    Code = Code & "                    // to today if nothing is selected." & vbCrLf

                    Code = Code & "                    var selDates = " & controlName(Rs, Fieldidx) & "calendar.getSelectedDates();" & vbCrLf

                    Code = Code & "                    var resetDate;" & vbCrLf

                    Code = Code & "        " & vbCrLf

                    Code = Code & "                    if (selDates.length > 0) {" & vbCrLf

                    Code = Code & "                        resetDate = selDates[0];" & vbCrLf

                    Code = Code & "                    } else {" & vbCrLf

                    Code = Code & "                        resetDate = " & controlName(Rs, Fieldidx) & "calendar.today;" & vbCrLf

                    Code = Code & "                    }" & vbCrLf

                    Code = Code & "        " & vbCrLf

                    Code = Code & "                    " & controlName(Rs, Fieldidx) & "calendar.cfg.setProperty(""pagedate"", resetDate);" & vbCrLf

                    Code = Code & "                    " & controlName(Rs, Fieldidx) & "calendar.render();" & vbCrLf

                    Code = Code & "                }" & vbCrLf

                    Code = Code & "        " & vbCrLf

                    Code = Code & "               function clearHandler() {" & vbCrLf

                    Code = Code & "                    	Dom.get(""" & controlName(Rs, Fieldidx) & "date_field"").value = """";" & vbCrLf

                    Code = Code & "		Dom.get(""" & controlName(Rs, Fieldidx) & "month_field"").value = """";" & vbCrLf

                    Code = Code & "		Dom.get(""" & controlName(Rs, Fieldidx) & "year_field"").value = """";" & vbCrLf

                    Code = Code & "		Dom.get(""" & controlName(Rs, Fieldidx) & """).value  = """";" & vbCrLf

                    Code = Code & "		" & controlName(Rs, Fieldidx) & "dialog.hide();" & vbCrLf

                    Code = Code & "                }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "                " & controlName(Rs, Fieldidx) & "dialog = new YAHOO.widget.Dialog(""" & controlName(Rs, Fieldidx) & "container"", {" & vbCrLf

                    Code = Code & "                    visible:false," & vbCrLf

                    Code = Code & "                    context:[""" & controlName(Rs, Fieldidx) & "show"", ""tl"", ""bl""]," & vbCrLf

                    Code = Code & "                    buttons:[ {text:""Reset"", handler: resetHandler, isDefault:true}, {text:""Clear"", handler: clearHandler}]," & vbCrLf

                    Code = Code & "                    draggable:true," & vbCrLf

                    Code = Code & "                    close:true" & vbCrLf

                    Code = Code & "                });" & vbCrLf

                    Code = Code & "                " & controlName(Rs, Fieldidx) & "dialog.setHeader('" & PrintColumnName(columnRs, Rs.Fields(Fieldidx).Name) & "');" & vbCrLf

                    Code = Code & "                " & controlName(Rs, Fieldidx) & "dialog.setBody('<div id=""cal""></div>');" & vbCrLf

                    Code = Code & "                " & controlName(Rs, Fieldidx) & "dialog.render(document.body);" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "                " & controlName(Rs, Fieldidx) & "dialog.showEvent.subscribe(function() {" & vbCrLf

                    Code = Code & "                    if (YAHOO.env.ua.ie) {" & vbCrLf

                    Code = Code & "                        // Since we're hiding the table using yui-overlay-hidden, we " & vbCrLf

                    Code = Code & "                        // want to let the dialog know that the content size has changed, when" & vbCrLf

                    Code = Code & "                        // shown" & vbCrLf

                    Code = Code & "                        " & controlName(Rs, Fieldidx) & "dialog.fireEvent(""changeContent"");" & vbCrLf

                    Code = Code & "                    }" & vbCrLf

                    Code = Code & "                });" & vbCrLf

                    Code = Code & "            }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "            // Lazy Calendar Creation - Wait to create the Calendar until the first time the button is clicked." & vbCrLf

                    Code = Code & "            if (!" & controlName(Rs, Fieldidx) & "calendar) {" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "                " & controlName(Rs, Fieldidx) & "calendar = new YAHOO.widget.Calendar(""cal"", {" & vbCrLf

                    Code = Code & "                    iframe:false,          // Turn iframe off, since " & controlName(Rs, Fieldidx) & "container has iframe support." & vbCrLf

                    Code = Code & "                    hide_blank_weeks:true  // Enable, to demonstrate how we handle changing height, using changeContent" & vbCrLf

                    Code = Code & "                });" & vbCrLf

                    Code = Code & "                " & controlName(Rs, Fieldidx) & "calendar.render();" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "                " & controlName(Rs, Fieldidx) & "calendar.selectEvent.subscribe(function (p_sType, p_aArgs) {" & vbCrLf

                    Code = Code & "                    if (" & controlName(Rs, Fieldidx) & "calendar.getSelectedDates().length > 0) {" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "                        	var selDate = " & controlName(Rs, Fieldidx) & "calendar.getSelectedDates()[0];" & vbCrLf

                    Code = Code & "		if (p_aArgs) {" & vbCrLf

                    Code = Code & "						" & vbCrLf

                    Code = Code & "			Dom.get(""" & controlName(Rs, Fieldidx) & "month_field"").value = selDate.getMonth() + 1;" & vbCrLf

                    Code = Code & "			Dom.get(""" & controlName(Rs, Fieldidx) & "date_field"").value = selDate.getDate();" & vbCrLf

                    Code = Code & "			Dom.get(""" & controlName(Rs, Fieldidx) & "year_field"").value = selDate.getFullYear();" & vbCrLf

                    Code = Code & "			Dom.get(""" & controlName(Rs, Fieldidx) & """).value = Dom.get(""" & controlName(Rs, Fieldidx) & "date_field"").value + '/' + Dom.get(""" & controlName(Rs, Fieldidx) & "month_field"").value + '/' + Dom.get(""" & controlName(Rs, Fieldidx) & "year_field"").value;" & vbCrLf

                    Code = Code & "			}" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "                    } else {" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "                        	Dom.get(""" & controlName(Rs, Fieldidx) & "date_field"").value = """";" & vbCrLf

                    Code = Code & "		Dom.get(""" & controlName(Rs, Fieldidx) & "month_field"").value = """";" & vbCrLf

                    Code = Code & "		Dom.get(""" & controlName(Rs, Fieldidx) & "year_field"").value = """";" & vbCrLf

                    Code = Code & "		Dom.get(""" & controlName(Rs, Fieldidx) & """).value = """";" & vbCrLf

                    Code = Code & "	" & vbCrLf

                    Code = Code & "                    }" & vbCrLf

                    Code = Code & "                    " & controlName(Rs, Fieldidx) & "dialog.hide();" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "                });" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "                " & controlName(Rs, Fieldidx) & "calendar.renderEvent.subscribe(function() {" & vbCrLf

                    Code = Code & "                    // Tell Dialog it's contents have changed, which allows " & vbCrLf

                    Code = Code & "                    // " & controlName(Rs, Fieldidx) & "container to redraw the underlay (for IE6/Safari2)" & vbCrLf

                    Code = Code & "                    " & controlName(Rs, Fieldidx) & "dialog.fireEvent(""changeContent"");" & vbCrLf

                    Code = Code & "                });" & vbCrLf

                    Code = Code & "            }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "            var seldate = " & controlName(Rs, Fieldidx) & "calendar.getSelectedDates();" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "            if (seldate.length > 0) {" & vbCrLf

                    Code = Code & "                // Set the pagedate to show the selected date if it exists" & vbCrLf

                    Code = Code & "                " & controlName(Rs, Fieldidx) & "calendar.cfg.setProperty(""pagedate"", seldate[0]);" & vbCrLf

                    Code = Code & "                " & controlName(Rs, Fieldidx) & "calendar.render();" & vbCrLf

                    Code = Code & "            }" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "            " & controlName(Rs, Fieldidx) & "dialog.show();" & vbCrLf

                    Code = Code & "        });" & vbCrLf

                    Code = Code & "    });" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "</script>" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & " <input id=""" & controlName(Rs, Fieldidx) & "date_field"" type=""text"" name=""date"" maxLength=""2"" value=""<%=defaultValue_day%>""> / <input id=""" & controlName(Rs, Fieldidx) & "month_field"" type=""text"" name=""month"" maxLength=""2"" value=""<%=defaultValue_month%>""> / <input id=""" & controlName(Rs, Fieldidx) & "year_field"" type=""text"" name=""year"" maxLength=""4"" value=""<%=defaultValue_year%>"">" & vbCrLf

                    Code = Code & "<button type=""button"" id=""" & controlName(Rs, Fieldidx) & "show"" title=""Show Calendar""><img src=""../../yui/examples/calendar/assets/calbtn.gif"" width=""18"" height=""18"" alt=""Calendar"" ></button> dd/mm/yyyy" & vbCrLf

                    Code = Code & "<input id=""" & controlName(Rs, Fieldidx) & """ type=""hidden"" name=""" & controlName(Rs, Fieldidx) & """ value=""<%=defaultValue%>"">" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "<%" & vbCrLf

                    Code = Code & "end sub" & vbCrLf

                    Code = Code & "%>" & vbCrLf

                    Code = Code & "" & vbCrLf

                    WriteCodeFile(codeFilename, Code)

            End Select

        Next

        columnRs.Close()
        Cn.Close()


    End Sub

    Private Sub BuildControls_UserId(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFileName As String = ""
        Dim Cn As New ADODB.Connection

        Dim FieldRelationRs As New ADODB.Recordset
        Dim ColumnRS As New ADODB.Recordset

        Dim Fieldidx As Short

        Dim Criteria As String

        Cn.Open(ConnectionString)

        FieldRelationRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaForeignKeys)
        ColumnRS = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        For Fieldidx = 0 To Rs.Fields.Count - 1

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(Fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            ColumnRS.Filter = Criteria

            'If it is a foreign key
            '12 Oct 2009
            Criteria = "FK_TABLE_NAME= '" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            Criteria = Criteria & " AND FK_COLUMN_NAME= '" & Rs.Fields(Fieldidx).Name & "'"

            FieldRelationRs.Filter = ""
            FieldRelationRs.Filter = Criteria

            Select Case False

                Case FieldRelationRs.EOF

                    If LCase(FieldRelationRs.Fields("PK_TABLE_NAME").Value) = LCase("sys_useraccount") And _
                            LCase(FieldRelationRs.Fields("PK_COLUMN_NAME").Value) = LCase("sys_userId") Then

                        codeFileName = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

                        Code = "<%" & vbCrLf

                        Code = Code & "sub showControls_" & Rs.Fields(Fieldidx).Name & "(defaultValue, properties, javascript)" & vbCrLf

                        Code = Code & "end sub" & vbCrLf

                        Code = Code & "%>" & vbCrLf

                        Call WriteCodeFile(codeFileName, Code)

                    End If

            End Select


        Next


        FieldRelationRs.Close()

        ColumnRS.Close()

        Cn.Close()


    End Sub


    Private Sub BuildControlLabel_UserId(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFileName As String = ""
        Dim Cn As New ADODB.Connection

        Dim FieldRelationRs As New ADODB.Recordset
        Dim ColumnRS As New ADODB.Recordset

        Dim Fieldidx As Short

        Dim Criteria As String

        Cn.Open(ConnectionString)

        FieldRelationRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaForeignKeys)
        ColumnRS = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        For Fieldidx = 0 To Rs.Fields.Count - 1

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(Fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            ColumnRS.Filter = Criteria

            'If it is a foreign key
            '12 Oct 2009
            Criteria = "FK_TABLE_NAME= '" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            Criteria = Criteria & " AND FK_COLUMN_NAME= '" & Rs.Fields(Fieldidx).Name & "'"

            FieldRelationRs.Filter = ""
            FieldRelationRs.Filter = Criteria

            Select Case False

                Case FieldRelationRs.EOF

                    If LCase(FieldRelationRs.Fields("PK_TABLE_NAME").Value) = LCase("sys_useraccount") And _
                            LCase(FieldRelationRs.Fields("PK_COLUMN_NAME").Value) = LCase("sys_userId") Then

                        codeFileName = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

                        Code = "<%" & vbCrLf

                        Code = Code & "sub showControlLabel_" & Rs.Fields(Fieldidx).Name & vbCrLf

                        Code = Code & "end sub" & vbCrLf

                        Code = Code & "%>" & vbCrLf

                        Call WriteCodeFile(codeFileName, Code)

                    End If

            End Select


        Next


        FieldRelationRs.Close()

        ColumnRS.Close()

        Cn.Close()


    End Sub

    Private Sub BuildRsSource_RecordByUserId(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFileName As String = ""
        Dim Cn As New ADODB.Connection

        Dim FieldRelationRs As New ADODB.Recordset
        Dim ColumnRS As New ADODB.Recordset

        Dim Fieldidx As Short

        Dim Criteria As String

        Cn.Open(ConnectionString)

        FieldRelationRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaForeignKeys)
        ColumnRS = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        codeFileName = Path & "\recordByUserId.asp"

        For Fieldidx = 0 To Rs.Fields.Count - 1

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(Fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            ColumnRS.Filter = Criteria

            'If it is a foreign key
            '12 Oct 2009
            Criteria = "FK_TABLE_NAME= '" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            Criteria = Criteria & " AND FK_COLUMN_NAME= '" & Rs.Fields(Fieldidx).Name & "'"

            FieldRelationRs.Filter = ""
            FieldRelationRs.Filter = Criteria


            Select Case False

                Case FieldRelationRs.EOF

                    If LCase(FieldRelationRs.Fields("PK_TABLE_NAME").Value) = LCase("sys_useraccount") And _
                            LCase(FieldRelationRs.Fields("PK_COLUMN_NAME").Value) = LCase("sys_userId") Then

                        Code = "<%		" & vbCrLf

                        Code = Code & "	function recordByUserId" & vbCrLf

                        Code = Code & "" & vbCrLf

                        Code = Code & "		recordByUserId = "" where " & Rs.Fields(Fieldidx).Name & " = '"" & Session(TokenName).fields(""" & FieldRelationRs.Fields("PK_COLUMN_NAME").Value & """) & ""'""" & vbCrLf

                        Code = Code & "" & vbCrLf

                        Code = Code & "	end function" & vbCrLf

                        Code = Code & "%>" & vbCrLf

                        Call WriteCodeFile(codeFileName, Code)

                    End If

            End Select


        Next


        FieldRelationRs.Close()

        ColumnRS.Close()

        Cn.Close()


    End Sub
    Private Sub BuildPageTitle(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFileName As String = ""
        Dim CN As New ADODB.Connection

        Dim TableRs As New ADODB.Recordset
        Dim Value As String = ""

        Dim Criteria As String

        CN.Open(ConnectionString)

        TableRs = CN.OpenSchema(ADODB.SchemaEnum.adSchemaTables)

        If Not IsDBNull(Rs.Fields(0).Properties(1).Value) Then

            Value = Rs.Fields(0).Properties(1).Value

        End If

        Criteria = "TABLE_NAME = '" & Value & "'"

        TableRs.Filter = Criteria

        codeFileName = Path & "\pageTitle.asp"

        Code = "<%" & vbCrLf

        Code = Code & "	function getPageTitle" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "		getPageTitle = """ & PrintTableName(TableRs, Value) & """" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "	end function" & vbCrLf

        Code = Code & "%>" & vbCrLf

        WriteCodeFile(codeFileName, Code)

        TableRs.Close()

    End Sub
    Public Sub BuildConnectionString(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFileName As String = ""
        codeFileName = Path & "\connectionString.asp"

        Code = "<%" & vbCrLf

        Code = Code & "	function getConnectionString" & vbCrLf

        Code = Code & "	" & vbCrLf

        Code = Code & "		getConnectionString = """ & Replace(ConnectionString, """", """""") & """" & vbCrLf

        Code = Code & " " & vbCrLf

        Code = Code & "	end function" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "%>" & vbCrLf

        WriteCodeFile(codeFileName, Code)

    End Sub
    Private Sub BuildrsSource(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFilename As String = ""

        codeFilename = Path & "\rsSource.asp"

        Code = "<!-- #include File=""recordByUserId\recordByUserId.asp"" -->" & vbCrLf

        Code = Code & "<%" & vbCrLf

        Code = Code & "	function rsSource" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "		rsSource = """ & Rs.Source & """" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "		rsSource = rsSource & recordByUserId" & vbCrLf

        Code = Code & "		" & vbCrLf

        Code = Code & "	end function		" & vbCrLf

        Code = Code & "	" & vbCrLf

        Code = Code & "%>" & vbCrLf

        Code = Code & "" & vbCrLf

        WriteCodeFile(codeFilename, Code)

    End Sub
    Private Sub BuildrsName(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFileName As String = ""

        codeFileName = Path & "\rsName.asp"

        Code = "<!-- #include File=""cnName.asp"" -->" & vbCrLf

        Code = Code & "<%" & vbCrLf

        Code = Code & "	function rsName" & vbCrLf

        Code = Code & "	" & vbCrLf

        Code = Code & "		rsName = cnName & ""_" & Rs.Fields(0).Properties(1).Value & """" & vbCrLf

        Code = Code & " " & vbCrLf

        Code = Code & "	end function" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "%>" & vbCrLf

        WriteCodeFile(codeFileName, Code)

    End Sub

    Public Sub BuildcnName(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFileName As String = ""

        codeFileName = Path & "\cnName.asp"

        Code = Code & "<%" & vbCrLf

        Code = Code & "	function cnName" & vbCrLf

        Code = Code & "	" & vbCrLf

        Code = Code & "		cnName =  """ & genCnName(ConnectionString) & """" & vbCrLf

        Code = Code & " " & vbCrLf

        Code = Code & "	end function" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "%>" & vbCrLf

        WriteCodeFile(codeFileName, Code)

    End Sub

    Private Function genCnName(ByVal connectionString As String) As String

        Dim strLen As String
        Dim code As Double
        Dim c As Integer
        Dim returnString As String

        strLen = Len(connectionString)

        returnString = ""

        code = 0

        For c = 1 To strLen

            code = code + (Asc(Mid(connectionString, c, 1)) + c)

        Next

        returnString = Hex(code) & Hex(Len(connectionString))

        GenCNName = returnString

    End Function

    Private Sub BuildSave_Fields(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim Fieldidx As Short
        Dim codeFileName As String = ""

        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If

        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFileName = Path & "\save_" & Rs.Fields(Fieldidx).Name & ".asp"

            Code = "<!-- #include File=""..\endecrypt\endecrypt_" & Rs.Fields(Fieldidx).Name & ".asp"" -->" & vbCrLf

            Code = Code & "<%" & vbCrLf

            Code = Code & "sub save_" & Rs.Fields(Fieldidx).Name & "(errorFlag)" & vbCrLf

            Code = Code & "" & vbCrLf

            Code = Code & "	On Error Resume Next" & vbCrLf

            Code = Code & "	Value(" & Fieldidx & ")= getFieldValue(""txtfield_" & Rs.Fields(Fieldidx).Name & """, Cn, Rs, """ & Rs.Fields(Fieldidx).Name & """)" & vbCrLf

            Code = Code & "	Rs(""" & Rs.Fields(Fieldidx).Name & """) = encrypt_" & Rs.Fields(Fieldidx).Name & "(Value(" & Fieldidx & "))" & vbCrLf

            Code = Code & "	if err.number <> 0 then" & vbCrLf

            Code = Code & " 		fieldMessage(" & Fieldidx & ") = err.description" & vbCrLf

            Code = Code & " 		errorFlag = true" & vbCrLf

            Code = Code & " 		err.clear" & vbCrLf

            Code = Code & "	end if" & vbCrLf

            Code = Code & "" & vbCrLf

            Code = Code & "end sub" & vbCrLf

            Code = Code & "%>" & vbCrLf

            WriteCodeFile(codeFileName, Code)

        Next

    End Sub
    Public Sub BuildSave(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim Fieldidx As Short
        Dim codeFileName As String = ""
        Call BuildSave_Fields(ConnectionString, Rs, Path)
        Call BuildSave_UserId(ConnectionString, Rs, Path)
        Call BuildSave_UserId(ConnectionString, Rs, Path & "\resources\userId")

        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If

        codeFileName = Path & "\save.asp"

        Code = "<!-- #include File=""getFieldValue.asp"" -->" & vbCrLf

        For Fieldidx = 0 To Rs.Fields.Count - 1

            Code = Code & "<!-- #include File=""save_" & Rs.Fields(Fieldidx).Name & ".asp"" -->" & vbCrLf

        Next

        Code = Code & "<%" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "Dim fieldMessage(" & Rs.Fields.Count - 1 & ")" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "Sub saveRecord(Cn, Rs)" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "Dim errorFlag" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "select case rs.editmode" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "case 2	'Addnew" & vbCrLf

        Code = Code & "	message = ""Record added successfully""" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "case else" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "	message = ""Record updated successfully""" & vbCrLf

        Code = Code & "	" & vbCrLf

        Code = Code & "end select" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "On Error Resume Next" & vbCrLf

        Code = Code & " errorFlag = false" & vbCrLf

        For Fieldidx = 0 To Rs.Fields.Count - 1

            Code = Code & "call save_" & Rs.Fields(Fieldidx).Name & "(errorFlag)" & vbCrLf

        Next

        Code = Code & "" & vbCrLf

        Code = Code & "if errorFlag = false then" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "    Rs.Update" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "     if err.number <> 0 then" & vbCrLf

        Code = Code & "	call restoreValue" & vbCrLf

        Code = Code & "         	Message = err.description" & vbCrLf

        Code = Code & "         	err.clear" & vbCrLf

        Code = Code & "     else" & vbCrLf

        Code = Code & "         	call updateNumberSequence(Cn, Rs)" & vbCrLf

        Code = Code & "         	isSaved = True" & vbCrLf

        Code = Code & "     end if" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "else" & vbCrLf

        Code = Code & "	call restoreValue" & vbCrLf

        Code = Code & "	Message = """"" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "end if" & vbCrLf

        Code = Code & "end sub" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "sub restoreValue" & vbCrLf

        Code = Code & "	" & vbCrLf

        For Fieldidx = 0 To Rs.Fields.Count - 1

            Code = Code & "	value(" & Fieldidx & ") = Request(""txtfield_" & Rs(Fieldidx).Name & """)" & vbCrLf

        Next
        Code = Code & "" & vbCrLf

        Code = Code & "end sub" & vbCrLf

        Code = Code & "" & vbCrLf

        Code = Code & "%>" & vbCrLf

        WriteCodeFile(codeFileName, Code)

    End Sub

    Private Sub BuildSave_UserId(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFileName As String = ""
        Dim Cn As New ADODB.Connection

        Dim Fieldidx As Short

        Cn.Open(ConnectionString)

        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If

        For Fieldidx = 0 To Rs.Fields.Count - 1

            If LCase(Rs.Fields(Fieldidx).Name) = LCase("sys_userId") Then

                codeFileName = Path & "\save_" & Rs.Fields(Fieldidx).Name & ".asp"

                Code = "<!-- #include File=""..\endecrypt\endecrypt_" & Rs.Fields(Fieldidx).Name & ".asp"" -->" & vbCrLf

                Code = Code & "<%" & vbCrLf

                Code = Code & "sub save_" & Rs.Fields(Fieldidx).Name & "(errorFlag)" & vbCrLf

                Code = Code & "" & vbCrLf

                Code = Code & "	On Error Resume Next" & vbCrLf

                Code = Code & "	Value(" & Fieldidx & ") = Session(TokenName).fields(""" & Rs.Fields(Fieldidx).Name & """)" & vbCrLf

                Code = Code & "	Rs(""" & Rs.Fields(Fieldidx).Name & """) = encrypt_" & Rs.Fields(Fieldidx).Name & "(Value(" & Fieldidx & "))" & vbCrLf

                Code = Code & "	if err.number <> 0 then" & vbCrLf

                Code = Code & " 		fieldMessage(" & Fieldidx & ") = err.description" & vbCrLf

                Code = Code & " 		errorFlag = true" & vbCrLf

                Code = Code & " 		err.clear" & vbCrLf

                Code = Code & "	end if" & vbCrLf

                Code = Code & "" & vbCrLf

                Code = Code & "end sub" & vbCrLf

                Code = Code & "%>" & vbCrLf

                WriteCodeFile(codeFileName, Code)

            End If

        Next

    End Sub

    Private Sub BuildControls_Textbox(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFilename As String = ""
        Dim Fieldidx As Short

        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If


        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFilename = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

            Code = "<%" & vbCrLf

            Code = Code & "sub showControls_" & Rs.Fields(Fieldidx).Name & "(defaultValue, properties, javascript)" & vbCrLf

            Code = Code & "%>" & vbCrLf

            Select Case Rs.Fields(Fieldidx).Type

                Case 203

                    Code = Code & "        <textarea wrap=""OFF"" name=""txtfield_" & Rs.Fields(Fieldidx).Name & """ <%=properties%> ""<%=javascript%>""><%=defaultValue%></textarea>" & vbCrLf

                Case Else

                    Code = Code & "        <input  class=""control"" type=""text"" name=""txtfield_" & Rs.Fields(Fieldidx).Name & """  value=""<%=defaultValue%>""  <%=properties%> ""<%=javascript%>"">" & vbCrLf

            End Select


            Code = Code & "<%" & vbCrLf

            Code = Code & "end sub" & vbCrLf

            Code = Code & "%>" & vbCrLf

            WriteCodeFile(codeFilename, Code)

        Next

    End Sub


    Private Sub BuildControls_RichTextBox(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim codeFilename As String = ""
        Dim Fieldidx As Short

        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If


        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFilename = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

            Select Case Rs.Fields(Fieldidx).Type

                Case 203

                    Code = ""

                    Code = Code & "<%" & vbCrLf

                    Code = Code & "sub showControls_" & Rs.Fields(Fieldidx).Name & "(defaultValue, properties, javascript)" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & "%>" & vbCrLf

                    Code = Code & "      " & vbCrLf

                    Code = Code & " " & vbCrLf

                    Code = Code & "<div class=""yui-skin-sam"">" & vbCrLf

                    Code = Code & " <textarea id=""txtfield_" & Rs.Fields(Fieldidx).Name & """ name=""txtfield_" & Rs.Fields(Fieldidx).Name & """ rows=""20"" cols=""75"" ><%=defaultValue%></textarea>" & vbCrLf

                    Code = Code & "</div>" & vbCrLf

                    Code = Code & "<script>" & vbCrLf

                    Code = Code & "" & vbCrLf

                    Code = Code & " " & vbCrLf

                    Code = Code & "(function() {" & vbCrLf

                    Code = Code & "    var Dom = YAHOO.util.Dom," & vbCrLf

                    Code = Code & "        Event = YAHOO.util.Event;" & vbCrLf

                    Code = Code & "    " & vbCrLf

                    Code = Code & "    var myConfig = {" & vbCrLf

                    Code = Code & "        height: '300px'," & vbCrLf

                    Code = Code & "        width: '600px'," & vbCrLf

                    Code = Code & "        dompath: true," & vbCrLf

                    Code = Code & "        focusAtStart: false," & vbCrLf

                    Code = Code & "        handleSubmit: true" & vbCrLf

                    Code = Code & "    };" & vbCrLf

                    Code = Code & " " & vbCrLf

                    Code = Code & "    //YAHOO.log('Create the Editor..', 'info', 'example');" & vbCrLf

                    Code = Code & "    var txtfield_" & Rs.Fields(Fieldidx).Name & "_editor = new YAHOO.widget.SimpleEditor('txtfield_" & Rs.Fields(Fieldidx).Name & "', myConfig);" & vbCrLf

                    Code = Code & "    txtfield_" & Rs.Fields(Fieldidx).Name & "_editor._defaultToolbar.buttonType = 'advanced';    " & vbCrLf

                    Code = Code & "    txtfield_" & Rs.Fields(Fieldidx).Name & "_editor.render();" & vbCrLf

                    Code = Code & " " & vbCrLf

                    Code = Code & "})();" & vbCrLf

                    Code = Code & "</script>" & vbCrLf

                    Code = Code & " " & vbCrLf

                    Code = Code & "<%" & vbCrLf

                    Code = Code & "end sub" & vbCrLf

                    Code = Code & "%>" & vbCrLf

                    Code = Code & "" & vbCrLf

                    WriteCodeFile(codeFilename, Code)


            End Select




        Next

    End Sub
    Private Sub BuildControls_Hyperlink(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim Fieldidx As Short
        Dim codeFileName As String = ""
        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If


        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFileName = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

            Code = "<%" & vbCrLf

            Code = Code & "sub showControls_" & Rs.Fields(Fieldidx).Name & "(defaultValue, properties, javascript)" & vbCrLf

            Code = Code & "%>" & vbCrLf

            Code = Code & "        <a href=""javascript:details(form<%=Rs.AbsolutePosition%>)"" name=""txtfield_" & Rs.Fields(Fieldidx).Name & """><%=defaultValue%></a>&nbsp" & vbCrLf

            Code = Code & "<%" & vbCrLf

            Code = Code & "end sub" & vbCrLf

            Code = Code & "%>" & vbCrLf

            WriteCodeFile(codeFileName, Code)

        Next


    End Sub
    Private Sub BuildControls_Text(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim Fieldidx As Short
        Dim codeFileName As String = ""
        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If


        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFileName = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

            Code = "<%" & vbCrLf

            Code = Code & "sub showControls_" & Rs.Fields(Fieldidx).Name & "(defaultValue, properties, javascript)" & vbCrLf

            Code = Code & "%>" & vbCrLf

            Code = Code & "        <%=defaultValue%>&nbsp" & vbCrLf

            Code = Code & "<%" & vbCrLf

            Code = Code & "end sub" & vbCrLf

            Code = Code & "%>" & vbCrLf

            WriteCodeFile(codeFileName, Code)

        Next


    End Sub

    Private Sub BuildControlRemark_Required(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim Fieldidx As Short
        Dim codeFileName As String = ""

        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If


        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFileName = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

            Code = "<%" & vbCrLf

            Code = Code & "sub showControlsRemark_" & Rs.Fields(Fieldidx).Name & vbCrLf

            Code = Code & "%>" & vbCrLf

            Code = Code & "	<font face=""Arial, Helvetica, sans-serif"" size=""2"">* REQUIRED</font>" & vbCrLf

            Code = Code & "<br><font face=""Arial, Helvetica, sans-serif"" size=""2"" color=""red""><%=fieldMessage(" & Fieldidx & ")%></font>" & vbCrLf

            Code = Code & "<%" & vbCrLf

            Code = Code & "end sub" & vbCrLf

            Code = Code & "%>" & vbCrLf

            WriteCodeFile(codeFileName, Code)

        Next


    End Sub

    Private Sub BuildControlRemark_Abstract(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim Fieldidx As Short
        Dim codeFileName As String = ""

        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If

        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFileName = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

            Code = "<%" & vbCrLf

            Code = Code & "sub showControlsRemark_" & Rs.Fields(Fieldidx).Name & vbCrLf

            Code = Code & "end sub" & vbCrLf

            Code = Code & "%>" & vbCrLf

            WriteCodeFile(codeFileName, Code)

        Next

    End Sub

    Private Sub BuildControls_Abstract(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim Fieldidx As Short
        Dim codeFileName As String = ""
        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If


        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFileName = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

            Code = "<%" & vbCrLf

            Code = Code & "sub showControls_" & Rs.Fields(Fieldidx).Name & "(defaultValue, properties, javascript)" & vbCrLf

            Code = Code & "end sub" & vbCrLf

            Code = Code & "%>" & vbCrLf

            WriteCodeFile(codeFileName, Code)

        Next


    End Sub

    Private Sub BuildControlLabel_Text(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim fieldidx As Short
        Dim Code As String
        Dim Criteria As String
        Dim codeFileName As String = ""
        Dim columnRs As New ADODB.Recordset
        Dim Cn As New ADODB.Connection

        Cn.Open(ConnectionString)

        columnRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)


        For fieldidx = 0 To Rs.Fields.Count - 1

            codeFileName = Path & "\" & Rs.Fields(fieldidx).Name & ".asp"

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(fieldidx).Properties(1).Value & "'"

            columnRs.Filter = Criteria

            Code = "<%" & vbCrLf

            Code = Code & "sub showControlLabel_" & Rs.Fields(fieldidx).Name & vbCrLf

            Code = Code & "%>" & vbCrLf

            Code = Code & PrintColumnName(columnRs, Rs.Fields(fieldidx).Name) & ":" & vbCrLf

            Code = Code & "<%" & vbCrLf

            Code = Code & "end sub" & vbCrLf

            Code = Code & "%>" & vbCrLf

            WriteCodeFile(codeFileName, Code)

        Next

        columnRs.Close()

        Cn.Close()


    End Sub


    Private Sub BuildControlLabel_Abstract(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim fieldidx As Short
        Dim Code As String
        Dim Criteria As String
        Dim codeFileName As String = ""
        Dim columnRs As New ADODB.Recordset
        Dim Cn As New ADODB.Connection

        Cn.Open(ConnectionString)

        columnRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)


        For fieldidx = 0 To Rs.Fields.Count - 1

            codeFileName = Path & "\" & Rs.Fields(fieldidx).Name & ".asp"

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(fieldidx).Properties(1).Value & "'"

            columnRs.Filter = Criteria

            Code = "<%" & vbCrLf

            Code = Code & "sub showControlLabel_" & Rs.Fields(fieldidx).Name & vbCrLf

            Code = Code & "end sub" & vbCrLf

            Code = Code & "%>" & vbCrLf

            WriteCodeFile(codeFileName, Code)


        Next

        columnRs.Close()

        Cn.Close()


    End Sub

    Private Sub BuildControlLabel_LinkToMainTable(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim Cn As New ADODB.Connection
        Dim codeFileName As String = ""
        Dim FieldRelationRs As New ADODB.Recordset
        Dim ColumnRS As New ADODB.Recordset

        Dim Fieldidx As Short

        Dim Criteria As String

        Cn.Open(ConnectionString)

        FieldRelationRs = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaForeignKeys)
        ColumnRS = Cn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

        For Fieldidx = 0 To Rs.Fields.Count - 1

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(Fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            ColumnRS.Filter = Criteria

            'If it is a foreign key
            '12 Oct 2009
            Criteria = "FK_TABLE_NAME= '" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            Criteria = Criteria & " AND FK_COLUMN_NAME= '" & Rs.Fields(Fieldidx).Name & "'"

            FieldRelationRs.Filter = ""
            FieldRelationRs.Filter = Criteria

            'Search description of field
            '25 August 2010
            Criteria = "COLUMN_NAME = '" & Rs.Fields(Fieldidx).Name & "'"
            Criteria = Criteria & " AND TABLE_NAME  ='" & Rs.Fields(Fieldidx).Properties(1).Value & "'"

            ColumnRS.Filter = Criteria

            Select Case False

                Case FieldRelationRs.EOF

                    codeFileName = Path & "\" & Rs.Fields(Fieldidx).Name & ".asp"

                    Code = "<%" & vbCrLf

                    Code = Code & "sub showControlLabel_" & Rs.Fields(Fieldidx).Name & vbCrLf

                    Code = Code & "%>" & vbCrLf

                    Code = Code & "<a href=""../../" & FieldRelationRs.Fields("PK_TABLE_NAME").Value & "/grid"">" & PrintColumnName(ColumnRS, Rs.Fields(Fieldidx).Name) & "</a>:" & vbCrLf

                    Code = Code & "<%" & vbCrLf

                    Code = Code & "end sub" & vbCrLf

                    Code = Code & "%>" & vbCrLf

                    WriteCodeFile(codeFileName, Code)

            End Select


        Next

        ColumnRS.Close()

        Cn.Close()


    End Sub

    Private Function controlName(ByVal Rs As ADODB.Recordset, ByVal fieldIdx As Integer) As String

        controlName = "txt_" & Rs.Fields(fieldIdx).Properties(1).Value & "_" & Rs.Fields(fieldIdx).Name
        controlName = "txtfield_" & Rs.Fields(fieldIdx).Name

    End Function

End Module
