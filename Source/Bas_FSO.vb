Option Strict Off
Option Explicit On
Module Bas_Fso
	Private Const ModuleName As String = "Bas_Fso"
	Public Fso As New Scripting.FileSystemObject

    Public Function AppendFile(ByVal Path As String, ByVal Content As String) As Boolean
        On Error GoTo ErrorHandler

        AppendFile = False


        Dim Ts As Scripting.TextStream

        Ts = Fso.OpenTextFile(Path, Scripting.IOMode.ForAppending, True)
        Ts.WriteLine(Content)
        Ts.WriteBlankLines((1))
        Ts.Close()

        AppendFile = True

        Exit Function

ErrorHandler:
        AppendFile = False

    End Function
	
	Public Function WriteNewFile(ByVal Path As String, ByVal Content As String) As Boolean
		Dim WriteFile As Object
		
		On Error GoTo ErrorHandler
		
		
		Dim Ts As Scripting.TextStream
		
		Ts = Fso.CreateTextFile(Path, True)
		
		Ts.Write(Content)
		
		Ts.Close()
		
		'UPGRADE_NOTE: Object Ts may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Ts = Nothing
		
		WriteNewFile = True
		
		Exit Function
		
ErrorHandler: 
		'UPGRADE_WARNING: Couldn't resolve default property of object WriteFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		WriteFile = False
		
	End Function
	
	Public Function FileFindReplace(ByVal SourceFile As String, ByVal TargetFile As String, ByVal FindString As String, ByVal ReplaceString As String) As Boolean

        '24 March 2006
		
		Dim SourceTs As Scripting.TextStream
		Dim TargetTs As Scripting.TextStream
		
		Dim Temp As String
		Dim FileContent As String
		
		On Error GoTo ErrorHandler
		
		SourceTs = Fso.OpenTextFile(SourceFile, Scripting.IOMode.ForReading)
		
		TargetTs = Fso.OpenTextFile(SourceFile, Scripting.IOMode.ForWriting)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Ts.AtEndOfStream. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do Until SourceTs.AtEndOfStream

            'UPGRADE_WARNING: Couldn't resolve default property of object Ts.ReadLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Temp = SourceTs.ReadLine

            Temp = Replace(Temp, FindString, ReplaceString)

            TargetTs.WriteLine(Temp)

        Loop
		
		SourceTs.Close()
		
		TargetTs.Close()
		
		'UPGRADE_NOTE: Object SourceTs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		SourceTs = Nothing

		
		FileFindReplace = True
		
		Exit Function
		
ErrorHandler: 
		
		FileFindReplace = False
		
    End Function


    Public Function CopyFolderWithFiles(ByVal sourcePath As String, ByVal targetPath As String) As Boolean

        If Fso.FolderExists(targetPath) = False Then

            Call Fso.CreateFolder(targetPath)

        End If

        Try

            Call Fso.CopyFile(sourcePath & "\*.*", targetPath, True)

            CopyFolderWithFiles = True

        Catch ex As Exception

            CopyFolderWithFiles = False

        End Try

    End Function

    Public Sub CreatePath(ByVal path As String)

        '15 April 2011
        'David Tsang

        Dim parentPath As String = ""

        If Fso.FolderExists(path) = True Then

            Exit Sub

        End If

        parentPath = Fso.GetParentFolderName(path)

        If Fso.FolderExists(parentPath) = False Then

            Call CreatePath(parentPath)

        End If

        Call Fso.CreateFolder(path)

    End Sub

End Module