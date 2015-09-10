Imports System.IO
Imports System.Net
Imports Acrobat

Module Reporting
    Public ReadOnly dateLastDownloaded As Date = Now.Date.AddDays(-1) 'Directory.GetLastWriteTime(folderLocation & "Output").Date

    Public Sub DownloadAllReports()
        Report.Report1.DownloadAllReports()
        Report.Report2.DownloadAllReports()
        Report.Report3.DownloadAllReports()
        Report.Report4.DownloadAllReports()
        Report.Report5.DownloadAllReports()
        Report.Report6.DownloadAllReports()
        Report.Report7.DownloadAllReports()
        Report.Report8.DownloadAllReports()
        Report.Report9.DownloadAllReports()
        Report.Report10.DownloadAllReports()
        Report.Report11.DownloadAllReports()
    End Sub

    Public Function CreatePacketPerson1() As Packet

        Dim packetPerson1 As Packet = New Packet("Person1")

        With packetPerson1
            .InsertPagesWithSection(Report.Report1.GetAllReports, "SECTION 1")
            .InsertPagesWithSection(Report.Report2.GetAllReports, "SECTION 1")
            .InsertPagesWithSection(Report.Report4.GetAllReports, "SECTION 1")
            .InsertPagesWithSections(Report.Report11.GetAllReports, {"SECTION 1", "SECTION 2"})
            .InsertAllPages(Report.Report9.GetAllReports)
            .Save()
            .Close()
        End With

        Return packetPerson1

    End Function

    Public Function CreatePacketPerson2() As Packet

        Dim packetPerson2 As New Packet("Person2")

        With packetPerson2
            .InsertPagesWithExactSections(Report.Report1.GetAllReports, {"SECTION 7", "SECTION 8", "SECTION 9"})
            .InsertAllPages(Report.Report8.GetAllReports)
            .InsertAllPages(Report.Report6.GetAllReports)
            .InsertAllPages(Report.Report9.GetAllReports)
            .InsertAllPages(Report.Report7.GetAllReports)
            .Save()
            .Close()
        End With

        Return packetPerson2

    End Function

End Module
