Option Strict Off
Option Explicit On
Module Bas_ODBC
	Private Const ModuleName As String = "Bas_ODBC"
    Public Function GetDataLinks() As String


        Dim MSDASCObj As MSDASC.DataLinks
        Dim cn As New ADODB.Connection

        MSDASCObj = New MSDASC.DataLinks
        MSDASCObj.PromptEdit(cn)

        GetDataLinks = cn.ConnectionString


    End Function
End Module