Module Bas_adSchemaTables
    Private Const ModuleName As String = "Bas_asSchemaTables"

    Public Function PrintTableName(ByRef adSchemaTablesRs As ADODB.Recordset, Optional ByVal defaultName As String = "") As String

        Dim returnString As String

        If adSchemaTablesRs.EOF = True And adSchemaTablesRs.BOF = True Then

            returnString = defaultName

        Else

            returnString = adSchemaTablesRs("TABLE_NAME").Value


            If Not adSchemaTablesRs("DESCRIPTION").Value Is DBNull.Value Then

                returnString = adSchemaTablesRs("DESCRIPTION").Value


            End If

        End If

        PrintTableName = returnString

    End Function
End Module
