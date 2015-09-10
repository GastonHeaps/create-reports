Imports System.IO


Module Main

    Private scriptFilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Adobe\Acrobat\Privileged\11.0\JavaScripts\"
    Private scriptFileName As String = "RemovePages.js"

    Public downloadFilePath As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\Reports\"

    Private gApp As AcroApp

    Sub Main()

        UnloadResources()

        gApp = CreateObject("AcroExch.App")

        'daysSinceLastDownload = DateDiff(DateInterval.Day, Directory.GetLastWriteTime(folderLocation & "Output"), Now)
        'daysSinceLastDownload = 0

        Reporting.DownloadAllReports()

        Reporting.CreatePacketPerson1()
        Reporting.CreatePacketPerson2()

        gApp.Exit()
        gApp = Nothing

    End Sub

    Private Sub UnloadResources()

        If Not My.Computer.FileSystem.DirectoryExists(scriptFilePath) Then
            My.Computer.FileSystem.CreateDirectory(scriptFilePath)
        End If

        If Not My.Computer.FileSystem.FileExists(scriptFilePath & scriptFileName) Then
            My.Computer.FileSystem.WriteAllBytes(scriptFilePath & scriptFileName, My.Resources.RemovePages, False)
        ElseIf File.ReadAllBytes(scriptFilePath & scriptFileName).Length <> My.Resources.RemovePages.Length Then
            My.Computer.FileSystem.WriteAllBytes(scriptFilePath & scriptFileName, My.Resources.RemovePages, False)
        End If

    End Sub

End Module
