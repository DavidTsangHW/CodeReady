Module Bas_ASPEndecrypt
    Public Sub BuildEndecrypt_Fields(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim Fieldidx As Short
        Dim codeFileName As String = ""
        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If

        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFileName = Path & "\endecrypt_" & Rs.Fields(Fieldidx).Name & ".asp"

            Code = ""

            Code = Code & "<%" & vbCrLf

            Code = Code & "function encrypt_" & Rs.Fields(Fieldidx).Name & "(value)" & vbCrLf

            Code = Code & "	" & vbCrLf

            Code = Code & "	On Error Resume Next" & vbCrLf

            Code = Code & "		" & vbCrLf

            Code = Code & "	encrypt_" & Rs.Fields(Fieldidx).Name & " = EnDeCrypt(value, """ & Rs.Fields(Fieldidx).Name & """)" & vbCrLf

            Code = Code & "		" & vbCrLf

            Code = Code & "	" & vbCrLf

            Code = Code & "end function" & vbCrLf

            Code = Code & "" & vbCrLf

            Code = Code & "function decrypt_" & Rs.Fields(Fieldidx).Name & "(value)" & vbCrLf

            Code = Code & "	" & vbCrLf

            Code = Code & "	On Error Resume Next" & vbCrLf

            Code = Code & "		" & vbCrLf

            Code = Code & "	decrypt_" & Rs.Fields(Fieldidx).Name & " = EnDeCrypt(value, """ & Rs.Fields(Fieldidx).Name & """)" & vbCrLf

            Code = Code & "" & vbCrLf

            Code = Code & "end function" & vbCrLf

            Code = Code & "" & vbCrLf

            Code = Code & "%>" & vbCrLf

            Code = Code & "" & vbCrLf

            WriteCodeFile(codeFileName, Code)

        Next

    End Sub

    Public Sub BuildEndecrypt_FieldsAbstract(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim Code As String = ""
        Dim Fieldidx As Short
        Dim codeFileName As String = ""
        If Fso.FolderExists(Path) = False Then

            Call CreatePath(Path)

        End If

        For Fieldidx = 0 To Rs.Fields.Count - 1

            codeFileName = Path & "\endecrypt_" & Rs.Fields(Fieldidx).Name & ".asp"

            Code = ""

            Code = Code & "<%" & vbCrLf

            Code = Code & "function encrypt_" & Rs.Fields(Fieldidx).Name & "(value)" & vbCrLf

            Code = Code & "	" & vbCrLf

            Code = Code & "	On Error Resume Next" & vbCrLf

            Code = Code & "		" & vbCrLf

            Code = Code & "	encrypt_" & Rs.Fields(Fieldidx).Name & " = value" & vbCrLf

            Code = Code & "		" & vbCrLf

            Code = Code & "	" & vbCrLf

            Code = Code & "end function" & vbCrLf

            Code = Code & "" & vbCrLf

            Code = Code & "function decrypt_" & Rs.Fields(Fieldidx).Name & "(value)" & vbCrLf

            Code = Code & "	" & vbCrLf

            Code = Code & "	On Error Resume Next" & vbCrLf

            Code = Code & "		" & vbCrLf

            Code = Code & "	decrypt_" & Rs.Fields(Fieldidx).Name & " = value " & vbCrLf

            Code = Code & "" & vbCrLf

            Code = Code & "end function" & vbCrLf

            Code = Code & "" & vbCrLf

            Code = Code & "%>" & vbCrLf

            Code = Code & "" & vbCrLf

            WriteCodeFile(codeFileName, Code)

        Next

    End Sub
End Module
