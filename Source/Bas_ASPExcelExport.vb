Module Bas_ASPExcelExport
    Private Const ModuleName As String = "Bas_ASPExcelExport"

    Public Sub BuildASPExportForm(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Call BuildFramework(ConnectionString, Rs, Path)

        Call CopyDirectory(AppPath("ASP Libraries\export"), Path)

    End Sub


End Module
