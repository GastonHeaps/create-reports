Imports System.Net
Imports System.IO

Imports Acrobat

Public MustInherit Class Report
    Private Shared client As CookieAwareWebClient = SiteMinder.GetAuthenticatedWebClient

    Protected MustOverride Property mURLBase As String
    Protected MustOverride Property mURLName As String
    Protected MustOverride Property mReportName As String

    Protected mReportDate As Date
    Protected mReportURL As String
    Protected mFileName As String

    Protected mReportFileInfo As FileInfo

    Protected Sub New(reportDate As Date)
        Me.mReportDate = reportDate
    End Sub

    Protected MustOverride Function FormatURLDate(reportDate As Date) As String
    Protected MustOverride Function FormatFileDate(reportDate As Date) As String
    Public MustOverride Function GetPagesWithSection(searchString As String) As CAcroPDDoc
    Public MustOverride Function GetPagesWithSections(searchStrings As Array) As CAcroPDDoc
    Public MustOverride Function GetPagesWithExactSection(searchString As String) As CAcroPDDoc
    Public MustOverride Function GetPagesWithExactSections(searchStrings As Array) As CAcroPDDoc
    Public Function GetAllPages() As CAcroPDDoc
        Dim reportPDDoc As CAcroPDDoc

        reportPDDoc = CreateObject("AcroExch.PDDoc")

        reportPDDoc.Open(mReportFileInfo.FullName)
        Return reportPDDoc

    End Function
    Public Function GetFilePath() As String
        Return mReportFileInfo.FullName
    End Function
    Public Function Download() As Boolean

        'TODO Only download reports that haven't been downloaded
        If FileExists(mReportURL) Then
            If My.Computer.FileSystem.FileExists(downloadFilePath & mFileName) Then
                File.Delete(downloadFilePath & mFileName)
            End If
            'My.Computer.Network.DownloadFile(mReportURL, folderLocation & mFileName, "", "", False, 100, True)
            client.DownloadFile(mReportURL, downloadFilePath & mFileName)

            mReportFileInfo = My.Computer.FileSystem.GetFileInfo(downloadFilePath & mFileName)
            AddReport()
            Return True
        End If

        Return False

    End Function
    Protected Function FileExists(ByVal fileURL As String) As Boolean
        Dim request As WebRequest = WebRequest.Create(fileURL)
        Dim response As WebResponse

        Try
            response = request.GetResponse()
            response.Close()
            request = Nothing
        Catch ex As Exception
            request = Nothing
            Return False
        End Try
        Return True
    End Function
    Protected MustOverride Sub AddReport()

    Public MustInherit Class MTDReport
        Inherits Report

        Protected Sub New(reportDate As Date)
            MyBase.New(reportDate)
        End Sub

        Protected Shared Function GetDates() As List(Of Date)

            Dim dates As List(Of Date) = New List(Of Date)
            Dim months As List(Of Integer) = New List(Of Integer)
            Dim day As Date = dateLastDownloaded

            While day < Now.Date
                If Not months.Contains(day.Month) Then
                    months.Add(day.Month)
                    dates.Add(day)
                End If
                day = day.AddDays(1)
            End While

            Return dates

        End Function
        Protected Overrides Function FormatURLDate(reportDate As Date) As String
            Return reportDate.ToString("MMyyyy")
            'Return Right("0" & DatePart("m", reportDate.AddDays(-1)), 2) & DatePart("yyyy", reportDate.AddDays(-1))
        End Function
        Protected Overrides Function FormatFileDate(reportDate As Date) As String
            Return reportDate.ToString("yyyy-MM")
            'Return Right("0" & DatePart("m", reportDate.AddDays(-1)), 2) & DatePart("yyyy", reportDate.AddDays(-1))
        End Function
        Public Overrides Function GetPagesWithSection(searchString As String) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutMTDSection(searchString, dateLastDownloaded.Day) Then
                    Console.WriteLine("Found pages with """ & searchString & """ in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with """ & searchString & """ in " & mReportFileInfo.Name)
                End If
            End If

            reportJSOobject = Nothing
            Return reportPDDoc

        End Function
        Public Overrides Function GetPagesWithSections(searchStrings As Array) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutMTDSections(searchStrings, dateLastDownloaded.Day) Then
                    Console.WriteLine("Found pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                End If
            End If

            reportJSOobject = Nothing
            Return reportPDDoc

        End Function
        Public Overrides Function GetPagesWithExactSection(searchString As String) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutMTDSectionWithExactText(searchString, dateLastDownloaded.Day) Then
                    Console.WriteLine("Found pages with """ & searchString & """ in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with """ & searchString & """ in " & mReportFileInfo.Name)
                End If
            End If

            reportJSOobject = Nothing
            Return reportPDDoc

        End Function
        Public Overrides Function GetPagesWithExactSections(searchStrings As Array) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutMTDSectionsWithExactText(searchStrings, dateLastDownloaded.Day) Then
                    Console.WriteLine("Found pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                End If

                reportJSOobject = Nothing
            End If

            Return reportPDDoc

        End Function

    End Class
    Public MustInherit Class MonthlyReport
        Inherits Report

        Protected Sub New(reportDate As Date)
            MyBase.New(reportDate)
        End Sub

        Protected Shared Function GetDates() As List(Of Date)

            Dim dates As List(Of Date) = New List(Of Date)
            Dim months As List(Of Integer) = New List(Of Integer)
            Dim day As Date = dateLastDownloaded

            While day < Now.Date
                If Not months.Contains(day.Month) Then
                    months.Add(day.Month)
                    dates.Add(day)
                End If
                day = day.AddDays(1)
            End While

            Return dates

        End Function
        Protected Overrides Function FormatURLDate(reportDate As Date) As String
            Return reportDate.ToString("MMyyyy")
            'Return Right("0" & DatePart("m", reportDate.AddDays(-1)), 2) & DatePart("yyyy", reportDate.AddDays(-1))
        End Function
        Protected Overrides Function FormatFileDate(reportDate As Date) As String
            Return reportDate.ToString("yyyy-MM")
            'Return Right("0" & DatePart("m", reportDate.AddDays(-1)), 2) & DatePart("yyyy", reportDate.AddDays(-1))
        End Function
        Public Overrides Function GetPagesWithSection(searchString As String) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutSection(searchString) Then
                    Console.WriteLine("Found pages with """ & searchString & """ in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with """ & searchString & """ in " & mReportFileInfo.Name)
                End If
                reportJSOobject = Nothing
            End If

            Return reportPDDoc

        End Function
        Public Overrides Function GetPagesWithSections(searchStrings As Array) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutSections(searchStrings) Then
                    Console.WriteLine("Found pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                End If

                reportJSOobject = Nothing
            End If

            Return reportPDDoc

        End Function
        Public Overrides Function GetPagesWithExactSection(searchString As String) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutSectionWithExactText(searchString) Then
                    Console.WriteLine("Found pages with """ & searchString & """ in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with """ & searchString & """ in " & mReportFileInfo.Name)
                End If

                reportJSOobject = Nothing
            End If

            Return reportPDDoc

        End Function
        Public Overrides Function GetPagesWithExactSections(searchStrings As Array) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutSectionsWithExactText(searchStrings) Then
                    Console.WriteLine("Found pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                End If

                reportJSOobject = Nothing
            End If

            Return reportPDDoc

        End Function

    End Class
    Public MustInherit Class DailyReport
        Inherits Report

        Protected Sub New(reportDate As Date)
            MyBase.New(reportDate)
        End Sub

        Protected Shared Function GetDates() As List(Of Date)

            Dim dates As List(Of Date) = New List(Of Date)
            Dim day As Date = dateLastDownloaded

            While day < Now.Date
                dates.Add(day)
                day = day.AddDays(1)
            End While

            Return dates

        End Function
        Protected Overrides Function FormatURLDate(reportDate As Date) As String
            Return reportDate.ToString("MMddyyyy")
            'Return Right("0" & DatePart("m", reportDate.AddDays(-1)), 2) & Right("0" & DatePart("d", reportDate.AddDays(-1)), 2) & DatePart("yyyy", reportDate.AddDays(-1))
        End Function
        Protected Overrides Function FormatFileDate(reportDate As Date) As String
            Return reportDate.ToString("yyyy-MM-dd")
            'Return Right("0" & DatePart("m", reportDate.AddDays(-1)), 2) & Right("0" & DatePart("d", reportDate.AddDays(-1)), 2) & DatePart("yyyy", reportDate.AddDays(-1))
        End Function
        Public Overrides Function GetPagesWithSection(searchString As String) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutSection(searchString) Then
                    Console.WriteLine("Found pages with """ & searchString & """ in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with """ & searchString & """ in " & mReportFileInfo.Name)
                End If
            End If

            reportJSOobject = Nothing
            Return reportPDDoc

        End Function
        Public Overrides Function GetPagesWithSections(searchStrings As Array) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutSections(searchStrings) Then
                    Console.WriteLine("Found pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                End If
            End If

            reportJSOobject = Nothing
            Return reportPDDoc

        End Function
        Public Overrides Function GetPagesWithExactSection(searchString As String) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutSectionWithExactText(searchString) Then
                    Console.WriteLine("Found pages with """ & searchString & """ in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with """ & searchString & """ in " & mReportFileInfo.Name)
                End If
            End If

            reportJSOobject = Nothing
            Return reportPDDoc

        End Function
        Public Overrides Function GetPagesWithExactSections(searchStrings As Array) As CAcroPDDoc

            Dim reportPDDoc As CAcroPDDoc
            Dim reportJSOobject As Object

            reportPDDoc = CreateObject("AcroExch.PDDoc")

            If reportPDDoc.Open(mReportFileInfo.FullName) Then
                reportJSOobject = reportPDDoc.GetJSObject

                If reportJSOobject.RemovePagesWithoutSectionsWithExactText(searchStrings) Then
                    Console.WriteLine("Found pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                Else
                    Console.WriteLine("Could not find any pages with ")
                    For Each searchString As String In searchStrings
                        Console.Write("""" & searchString & """")
                        If searchString <> searchStrings(searchStrings.GetUpperBound(0)) Then
                            Console.Write(", ")
                        End If
                    Next
                    Console.Write(" in " & mReportFileInfo.Name)
                End If

                reportJSOobject = Nothing
            End If

            Return reportPDDoc

        End Function

    End Class

    Public Class Report1
        Inherits MTDReport

        Protected Overrides Property mURLBase As String = "https://example.com/Type1"
        Protected Overrides Property mURLName As String = "RP_1.pdf"
        Protected Overrides Property mReportName As String = "Report1.pdf"

        Public Shared mReports As Dictionary(Of Date, Report) = New Dictionary(Of Date, Report)

        Sub New(reportDate As Date)
            MyBase.New(reportDate)
            MyBase.mReportURL = mURLBase & FormatURLDate(reportDate) & "_" & mURLName
            MyBase.mFileName = FormatFileDate(reportDate) & "_" & mReportName
            MyBase.Download()
        End Sub

        Public Shared Sub DownloadAllReports()
            Dim Report As Report1

            For Each reportDate As Date In GetDates()
                Report = New Report1(reportDate)
            Next
        End Sub
        Public Shared Function GetAllReports() As Dictionary(Of Date, Report)
            Return mReports
        End Function
        Protected Overrides Sub AddReport()
            If Not mReports.ContainsKey(mReportDate) Then
                mReports.Add(mReportDate, Me)
            End If
        End Sub

    End Class
    Public Class Report2
        Inherits MTDReport

        Protected Overrides Property mURLBase As String = "https://example.com/Type1"
        Protected Overrides Property mURLName As String = "RP_2.pdf"
        Protected Overrides Property mReportName As String = "Report2.pdf"

        Protected Shared mReports As Dictionary(Of Date, Report) = New Dictionary(Of Date, Report)

        Sub New(reportDate As Date)
            MyBase.New(reportDate)
            MyBase.mReportURL = mURLBase & FormatURLDate(reportDate) & "_" & mURLName
            MyBase.mFileName = FormatFileDate(reportDate) & "_" & mReportName
            MyBase.Download()
        End Sub

        Public Shared Sub DownloadAllReports()
            Dim Report As Report2

            For Each reportDate As Date In GetDates()
                Report = New Report2(reportDate)
            Next
        End Sub
        Public Shared Function GetAllReports() As Dictionary(Of Date, Report)
            Return mReports
        End Function
        Protected Overrides Sub AddReport()
            If Not mReports.ContainsKey(mReportDate) Then
                mReports.Add(mReportDate, Me)
            End If
        End Sub

    End Class
    Public Class Report3
        Inherits MTDReport

        Protected Overrides Property mURLBase As String = "https://example.com/Type2"
        Protected Overrides Property mURLName As String = "RP_3.pdf"
        Protected Overrides Property mReportName As String = "Report3.pdf"

        Public Shared mReports As Dictionary(Of Date, Report) = New Dictionary(Of Date, Report)

        Sub New(reportDate As Date)
            MyBase.New(reportDate)
            MyBase.mReportURL = mURLBase & FormatURLDate(reportDate) & "_" & mURLName
            MyBase.mFileName = FormatFileDate(reportDate) & "_" & mReportName
            MyBase.Download()
        End Sub

        Public Shared Sub DownloadAllReports()
            Dim Report As Report3

            For Each reportDate As Date In GetDates()
                Report = New Report3(reportDate)
            Next
        End Sub
        Public Shared Function GetAllReports() As Dictionary(Of Date, Report)
            Return mReports
        End Function
        Protected Overrides Sub AddReport()
            If Not mReports.ContainsKey(mReportDate) Then
                mReports.Add(mReportDate, Me)
            End If
        End Sub

    End Class
    Public Class Report4
        Inherits MTDReport

        Protected Overrides Property mURLBase As String = "https://example.com/Type2"
        Protected Overrides Property mURLName As String = "RP_4.pdf"
        Protected Overrides Property mReportName As String = "Report4.pdf"

        Public Shared mReports As Dictionary(Of Date, Report) = New Dictionary(Of Date, Report)

        Sub New(reportDate As Date)
            MyBase.New(reportDate)
            MyBase.mReportURL = mURLBase & FormatURLDate(reportDate) & "_" & mURLName
            MyBase.mFileName = FormatFileDate(reportDate) & "_" & mReportName
            MyBase.Download()
        End Sub

        Public Shared Sub DownloadAllReports()
            Dim Report As Report4

            For Each reportDate As Date In GetDates()
                Report = New Report4(reportDate)
            Next
        End Sub
        Public Shared Function GetAllReports() As Dictionary(Of Date, Report)
            Return mReports
        End Function
        Protected Overrides Sub AddReport()
            If Not mReports.ContainsKey(mReportDate) Then
                mReports.Add(mReportDate, Me)
            End If
        End Sub


    End Class
    Public Class Report5
        Inherits MTDReport

        Protected Overrides Property mURLBase As String = "https://example.com/Type2"
        Protected Overrides Property mURLName As String = "RP_5.pdf"
        Protected Overrides Property mReportName As String = "Report5.pdf"

        Public Shared mReports As Dictionary(Of Date, Report) = New Dictionary(Of Date, Report)

        Sub New(reportDate As Date)
            MyBase.New(reportDate)
            MyBase.mReportURL = mURLBase & FormatURLDate(reportDate) & "_" & mURLName
            MyBase.mFileName = FormatFileDate(reportDate) & "_" & mReportName
            MyBase.Download()
        End Sub

        Public Shared Sub DownloadAllReports()
            Dim Report As Report5

            For Each reportDate As Date In GetDates()
                Report = New Report5(reportDate)
            Next
        End Sub
        Public Shared Function GetAllReports() As Dictionary(Of Date, Report)
            Return mReports
        End Function
        Protected Overrides Sub AddReport()
            If Not mReports.ContainsKey(mReportDate) Then
                mReports.Add(mReportDate, Me)
            End If
        End Sub

    End Class

    Public Class Report6
        Inherits MonthlyReport

        Protected Overrides Property mURLBase As String = "https://example.com/Type3"
        Protected Overrides Property mURLName As String = "RP_6.pdf"
        Protected Overrides Property mReportName As String = "Report6.pdf"

        Public Shared mReports As Dictionary(Of Date, Report) = New Dictionary(Of Date, Report)

        Sub New(reportDate As Date)
            MyBase.New(reportDate)
            MyBase.mReportURL = mURLBase & FormatURLDate(reportDate) & "_" & mURLName
            MyBase.mFileName = FormatFileDate(reportDate) & "_" & mReportName
            MyBase.Download()
        End Sub

        Public Shared Sub DownloadAllReports()
            Dim Report As Report6

            For Each reportDate As Date In GetDates()
                Report = New Report6(reportDate)
            Next
        End Sub
        Public Shared Function GetAllReports() As Dictionary(Of Date, Report)
            Return mReports
        End Function
        Protected Overrides Sub AddReport()
            If Not mReports.ContainsKey(mReportDate) Then
                mReports.Add(mReportDate, Me)
            End If
        End Sub

    End Class
    Public Class Report7
        Inherits MonthlyReport

        Protected Overrides Property mURLBase As String = "https://example.com/Type3"
        Protected Overrides Property mURLName As String = "RP_7.pdf"
        Protected Overrides Property mReportName As String = "Report7.pdf"

        Public Shared mReports As Dictionary(Of Date, Report) = New Dictionary(Of Date, Report)

        Sub New(reportDate As Date)
            MyBase.New(reportDate)
            MyBase.mReportURL = mURLBase & FormatURLDate(reportDate) & "_" & mURLName
            MyBase.mFileName = FormatFileDate(reportDate) & "_" & mReportName
            MyBase.Download()
        End Sub

        Public Shared Sub DownloadAllReports()
            Dim Report As Report7

            For Each reportDate As Date In GetDates()
                Report = New Report7(reportDate)
            Next
        End Sub
        Public Shared Function GetAllReports() As Dictionary(Of Date, Report)
            Return mReports
        End Function
        Protected Overrides Sub AddReport()
            If Not mReports.ContainsKey(mReportDate) Then
                mReports.Add(mReportDate, Me)
            End If
        End Sub

    End Class

    Public Class Report8
        Inherits DailyReport

        Protected Overrides Property mURLBase As String = "https://example.com/Type3"
        Protected Overrides Property mURLName As String = "RP_8.pdf"
        Protected Overrides Property mReportName As String = "Report8.pdf"

        Public Shared mReports As Dictionary(Of Date, Report) = New Dictionary(Of Date, Report)

        Sub New(reportDate As Date)
            MyBase.New(reportDate)
            MyBase.mReportURL = mURLBase & FormatURLDate(reportDate) & "_" & mURLName
            MyBase.mFileName = FormatFileDate(reportDate) & "_" & mReportName
            MyBase.Download()
        End Sub

        Public Shared Sub DownloadAllReports()
            Dim Report As Report8

            For Each reportDate As Date In GetDates()
                Report = New Report8(reportDate)
            Next
        End Sub
        Public Shared Function GetAllReports() As Dictionary(Of Date, Report)
            Return mReports
        End Function
        Protected Overrides Sub AddReport()
            If Not mReports.ContainsKey(mReportDate) Then
                mReports.Add(mReportDate, Me)
            End If
        End Sub

    End Class
    Public Class Report9
        Inherits DailyReport

        Protected Overrides Property mURLBase As String = "https://example.com/Type3"
        Protected Overrides Property mURLName As String = "RP_9.pdf"
        Protected Overrides Property mReportName As String = "Repor9.pdf"

        Public Shared mReports As Dictionary(Of Date, Report) = New Dictionary(Of Date, Report)

        Sub New(reportDate As Date)
            MyBase.New(reportDate)
            MyBase.mReportURL = mURLBase & FormatURLDate(reportDate) & "_" & mURLName
            MyBase.mFileName = FormatFileDate(reportDate) & "_" & mReportName
            MyBase.Download()
        End Sub

        Public Shared Sub DownloadAllReports()
            Dim Report As Report9

            For Each reportDate As Date In GetDates()
                Report = New Report9(reportDate)
            Next
        End Sub
        Public Shared Function GetAllReports() As Dictionary(Of Date, Report)
            Return mReports
        End Function
        Protected Overrides Sub AddReport()
            If Not mReports.ContainsKey(mReportDate) Then
                mReports.Add(mReportDate, Me)
            End If
        End Sub

    End Class
    Public Class Report10
        Inherits DailyReport

        Protected Overrides Property mURLBase As String = "https://example.com/Type3"
        Protected Overrides Property mURLName As String = "RP_10.pdf"
        Protected Overrides Property mReportName As String = "Report10.pdf"

        Public Shared mReports As Dictionary(Of Date, Report) = New Dictionary(Of Date, Report)

        Sub New(reportDate As Date)
            MyBase.New(reportDate)
            MyBase.mReportURL = mURLBase & FormatURLDate(reportDate) & "_" & mURLName
            MyBase.mFileName = FormatFileDate(reportDate) & "_" & mReportName
            MyBase.Download()
        End Sub

        Public Shared Sub DownloadAllReports()
            Dim Report As Report10

            For Each reportDate As Date In GetDates()
                Report = New Report10(reportDate)
            Next
        End Sub
        Public Shared Function GetAllReports() As Dictionary(Of Date, Report)
            Return mReports
        End Function
        Protected Overrides Sub AddReport()
            If Not mReports.ContainsKey(mReportDate) Then
                mReports.Add(mReportDate, Me)
            End If
        End Sub

    End Class
    Public Class Report11
        Inherits DailyReport

        Protected Overrides Property mURLBase As String = "https://example.com/Type2"
        Protected Overrides Property mURLName As String = "RP_11.pdf"
        Protected Overrides Property mReportName As String = "Report11.pdf"

        Public Shared mReports As Dictionary(Of Date, Report) = New Dictionary(Of Date, Report)

        Sub New(reportDate As Date)
            MyBase.New(reportDate)
            MyBase.mReportURL = mURLBase & FormatURLDate(reportDate) & "_" & mURLName
            MyBase.mFileName = FormatFileDate(reportDate) & "_" & mReportName
            MyBase.Download()
        End Sub

        Public Shared Sub DownloadAllReports()
            Dim Report As Report11

            For Each reportDate As Date In GetDates()
                Report = New Report11(reportDate)
            Next
        End Sub
        Public Shared Function GetAllReports() As Dictionary(Of Date, Report)
            Return mReports
        End Function
        Protected Overrides Sub AddReport()
            If Not mReports.ContainsKey(mReportDate) Then
                mReports.Add(mReportDate, Me)
            End If
        End Sub

    End Class
End Class
