Module Bas_adSchemaColumns

    Private Const ModuleName As String = "Bas_asSchemaColumn"

    Public Function PrintColumnName(ByRef adSchemaColumnsRs As ADODB.Recordset, Optional ByVal defaultName As String = "") As String

        Dim returnString As String

        If adSchemaColumnsRs.EOF = True And adSchemaColumnsRs.BOF = True Then

            returnString = defaultName

        Else

            returnString = adSchemaColumnsRs("COLUMN_NAME").Value

            If Not adSchemaColumnsRs("DESCRIPTION").Value Is DBNull.Value Then

                returnString = adSchemaColumnsRs("DESCRIPTION").Value

            End If

        End If

        PrintColumnName = returnString

    End Function

End Module
