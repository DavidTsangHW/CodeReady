
'Source from http://bytes.com/topic/visual-basic-net/answers/365349-copying-all-contents-one-folder-into-another
'Author: Mike McClaran
'Date: 05 Oct 2010

Module Bas_CopyDirectory

    Public Sub CopyDirectory(ByVal strSrc As String, ByVal strDest As String)
        Dim dirInfo As New System.IO.DirectoryInfo(strSrc)
        Dim fsInfo As System.IO.FileSystemInfo

        If Not System.IO.Directory.Exists(strDest) Then
            System.IO.Directory.CreateDirectory(strDest)
        End If

        For Each fsInfo In dirInfo.GetFileSystemInfos
            Dim strDestFileName As String = System.IO.Path.Combine(strDest, fsInfo.Name)

            If TypeOf fsInfo Is System.IO.FileInfo Then

                'Custom code handling
                If isCodeModified(strDestFileName) = False Then

                    System.IO.File.Copy(fsInfo.FullName, strDestFileName, True)
                    Call WriteFileLog(strDestFileName)

                End If

            Else

                CopyDirectory(fsInfo.FullName, strDestFileName)

            End If

        Next

    End Sub

End Module
