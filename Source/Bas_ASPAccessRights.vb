
Module Bas_ASPAccessRights

    Public Sub Build_AccessRights(ByVal ConnectionString As String, ByVal Rs As ADODB.Recordset, ByVal Path As String)

        Dim targetPath As String

        targetPath = Path

        Call CopyDirectory(AppPath("ASP Libraries\root"), targetPath)

        targetPath = Path & "\accessRights\database"

        Call CopyDirectory(AppPath("ASP Libraries\framework\database"), targetPath)

        Call BuildConnectionString(ConnectionString, Rs, targetPath)

        Call BuildcnName(ConnectionString, Rs, targetPath)

    End Sub

End Module
