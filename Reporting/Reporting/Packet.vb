﻿Imports Acrobat

Public Class Packet

    Private mFileName As String
    Private mPacketPDDoc As CAcroPDDoc

    Public Sub New(packetName As String)
        mPacketPDDoc = CreateObject("AcroExch.PDDoc")
        mPacketPDDoc.Create()

        mFileName = packetName & ".pdf"
    End Sub

    Public Sub InsertAllPages(dialerReports As Dictionary(Of Date, Report))

        For Each report As Report In dialerReports.Values

            Dim reportPDDoc = report.GetAllPages
            Dim reportJSObject As Object

            reportJSObject = reportPDDoc.GetJSObject()

            'reportJSObject.Print(False, 0, reportPDDoc.GetNumPages - 1, True, False, False, False, False)
            mPacketPDDoc.InsertPages(mPacketPDDoc.GetNumPages - 1, reportPDDoc, 0, reportPDDoc.GetNumPages, 0)

            reportPDDoc.Close()
        Next

    End Sub
    Public Sub InsertPagesWithSection(dialerReports As Dictionary(Of Date, Report), searchString As String)

        For Each report As Report In dialerReports.Values

            Dim reportPDDoc = report.GetPagesWithSection(searchString)
            Dim reportJSObject As Object

            reportJSObject = reportPDDoc.GetJSObject()

            'reportJSObject.Print(False, 0, reportPDDoc.GetNumPages - 1, True, False, False, False, False)
            mPacketPDDoc.InsertPages(mPacketPDDoc.GetNumPages - 1, reportPDDoc, 0, reportPDDoc.GetNumPages, 0)

            reportPDDoc.Close()
        Next

    End Sub
    Public Sub InsertPagesWithSections(dialerReports As Dictionary(Of Date, Report), searchStrings As Array)

        For Each report As Report In dialerReports.Values

            Dim reportPDDoc = report.GetPagesWithSections(searchStrings)
            Dim reportJSObject As Object

            reportJSObject = reportPDDoc.GetJSObject()

            'reportJSObject.Print(False, 0, reportPDDoc.GetNumPages - 1, True, False, False, False, False)
            mPacketPDDoc.InsertPages(mPacketPDDoc.GetNumPages - 1, reportPDDoc, 0, reportPDDoc.GetNumPages, 0)

            reportPDDoc.Close()
        Next

    End Sub
    Public Sub InsertPagesWithExactSection(dialerReports As Dictionary(Of Date, Report), searchString As String)

        For Each report As Report In dialerReports.Values

            Dim reportPDDoc = report.GetPagesWithExactSection(searchString)
            Dim reportJSObject As Object

            reportJSObject = reportPDDoc.GetJSObject()

            'reportJSObject.Print(False, 0, reportPDDoc.GetNumPages - 1, True, False, False, False, False)
            mPacketPDDoc.InsertPages(mPacketPDDoc.GetNumPages - 1, reportPDDoc, 0, reportPDDoc.GetNumPages, 0)

            reportPDDoc.Close()
        Next

    End Sub
    Public Sub InsertPagesWithExactSections(dialerReports As Dictionary(Of Date, Report), searchStrings As Array)

        For Each report As Report In dialerReports.Values

            Dim reportPDDoc = report.GetPagesWithExactSections(searchStrings)
            Dim reportJSObject As Object

            reportJSObject = reportPDDoc.GetJSObject()

            'reportJSObject.Print(False, 0, reportPDDoc.GetNumPages - 1, True, False, False, False, False)
            mPacketPDDoc.InsertPages(mPacketPDDoc.GetNumPages - 1, reportPDDoc, 0, reportPDDoc.GetNumPages, 0)

            reportPDDoc.Close()
        Next

    End Sub

    Public Function GetPDDoc() As CAcroPDDoc
        Return mPacketPDDoc
    End Function

    Public Sub Close()
        mPacketPDDoc.Close()
    End Sub

    Public Sub Save()
        mPacketPDDoc.Save(PDSaveFlags.PDSaveFull, downloadFilePath & "Output\" & mFileName)
        Console.WriteLine(mFileName & " saved to " & downloadFilePath & "Output")
    End Sub

    Public Sub Save(filePath As String)
        mPacketPDDoc.Save(PDSaveFlags.PDSaveFull, filePath & mFileName)
        Console.WriteLine(mFileName & " saved to " & filePath)
    End Sub
End Class
