Module Bas_customCodeHandling
    '28 March 2011
    'This module is to manage custom code.
    'The program writes time stamp and file size to a file log.
    'By comparing file time stamp and file size in the file log before writing code to file,
    'if current file time stamp or file size do not match with its value in the file log, the file is considered as modified, and will not be updated.

    Private Const Const_FileLogName As String = "files.log"
    Private Const Const_FileSeparator As String = vbTab

    Public Function isCodeModified(ByVal filepath As String) As Boolean

        Dim readTs As Scripting.TextStream
        Dim readLine As String = ""
        Dim parentPath As String = ""
        Dim logfilename As String = ""

        isCodeModified = True

        parentPath = Fso.GetParentFolderName(filepath)

        logfilename = parentPath & "\" & Const_FileLogName

        'If the code file is not exist
        If Fso.FileExists(filepath) = False Then

            isCodeModified = False
            Exit Function

        End If

        'If the log file is not exist, all code files in the folder will be considered as modified
        If Fso.FileExists(logfilename) = False Then

            isCodeModified = True
            Exit Function

        End If

        readTs = Fso.OpenTextFile(logfilename, Scripting.IOMode.ForReading, False)

        Do Until readTs.AtEndOfStream

            readLine = readTs.ReadLine

            'Check filename
            If Read_Separated_Text(readLine, Const_FileSeparator, 0) = Fso.GetFile(filepath).Name Then

                'Check date last modified
                If Fso.GetFile(filepath).DateLastModified = Read_Separated_Text(readLine, Const_FileSeparator, 1) Then

                    'Chech file size
                    If Fso.GetFile(filepath).Size = Read_Separated_Text(readLine, Const_FileSeparator, 2) Then

                        isCodeModified = False
                        Exit Do

                    End If

                    Exit Do

                End If

            End If

        Loop

        readTs.Close()

    End Function

    Public Function WriteFileLog(ByVal filepath As String) As Boolean
        '28 March 2011

        Dim headerString As String = ""

        Dim writeLine As String = ""
        Dim readLine As String = ""

        Dim parentPath As String
        Dim logfilename As String

        Dim readTs As Scripting.TextStream
        Dim writeTs As Scripting.TextStream

        Dim fileInfo As String = ""

        parentPath = Fso.GetParentFolderName(filepath)
        logfilename = parentPath & "\" & Const_FileLogName

        fileInfo = Fso.GetFileName(filepath) & Const_FileSeparator & Fso.GetFile(filepath).DateLastModified & Const_FileSeparator & Fso.GetFile(filepath).Size

        headerString = "This file log is built by CodeReady" & vbCrLf
        headerString = headerString & "CodeReady searches the file information stored in this file log before updating file. If the current file information does not match with the file information stored in this file, the file is considered as a modified file." & vbCrLf
        headerString = headerString & "CodeReady will not update for any modified files." & vbCrLf
        headerString = headerString & "Do not rename, remove, or edit this file" & vbCrLf & vbCrLf
        headerString = headerString & "Filename" & vbTab & "Date last modified" & vbCrLf

        If Fso.FileExists(logfilename) = True Then

            headerString = ""

            readTs = Fso.OpenTextFile(logfilename, Scripting.IOMode.ForReading, False)

            Do Until readTs.AtEndOfStream

                readLine = readTs.ReadLine & vbCrLf

                If Read_Separated_Text(readLine, Const_FileSeparator, 0) = Fso.GetFileName(filepath) Then

                    readLine = ""

                End If

                writeLine = writeLine & readLine

            Loop

            readTs.Close()

        End If

        writeLine = headerString & writeLine & fileInfo

        writeTs = Fso.OpenTextFile(logfilename, Scripting.IOMode.ForWriting, True)

        writeTs.WriteLine(writeLine)

        writeTs.Close()

    End Function

    Public Function WriteCodeFile(ByVal filename As String, ByVal code As String) As Boolean

        Dim WriteTs As Scripting.TextStream

        WriteCodeFile = False

        If isCodeModified(filename) = True Then

            Exit Function

        End If

        Call CreatePath(Fso.GetParentFolderName(filename))

        code = "<!-- " & Fso.GetFileName(filename) & " - " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision & " " & VB6.Format(Now, "yyyy/MM/dd HH:mm:ss") & "-->" & vbCrLf & code

        WriteTs = Fso.OpenTextFile(filename, Scripting.IOMode.ForWriting, True)

        WriteTs.WriteLine(code)

        WriteTs.Close()

        Call WriteFileLog(filename)

        WriteCodeFile = True

    End Function

End Module
