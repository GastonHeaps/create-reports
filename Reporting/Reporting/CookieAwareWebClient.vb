Imports System.Net

Public Class CookieAwareWebClient
    Inherits WebClient

    Private lastPage As String
    Private cookies As CookieContainer

    Sub New(cookies As CookieContainer)
        Me.cookies = cookies
    End Sub

    Protected Overrides Function GetWebRequest(ByVal address As System.Uri) As System.Net.WebRequest
        Dim R = MyBase.GetWebRequest(address)
        If TypeOf R Is HttpWebRequest Then
            With DirectCast(R, HttpWebRequest)
                .CookieContainer = cookies
                If Not lastPage Is Nothing Then
                    .Referer = lastPage
                End If
            End With
        End If
        lastPage = address.ToString()
        Return R
    End Function

End Class
