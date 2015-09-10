Imports System.Web
Imports System.IO
Imports System.Collections.Specialized
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Net

Module SiteMinder

    ' use DLLImport to import the WinINet SetCookie function,
    ' this is used only to create persistent cookies and is not strictly necessary
    <DllImport("wininet.dll", CharSet:=CharSet.Unicode)>
    Public Function InternetSetCookie(
        ByVal lpszUrl As String,
        ByVal lpszCookieName As String,
        ByVal lpszCookieData As String
        ) As Boolean
    End Function

    Private Const PROTECTED_DOMAIN_URI As String = ""
    Private Const PROTECTED_URL As String = ""
    Private username As String
    Private password As String

    Private Const COOKIE_TIMEOUT_MINUTES = 60
    Private Const FORM_CONTENT_TYPE As String = "application/x-www-form-urlencoded"
    Private Const POST_METHOD As String = "POST"
    Private Const SET_COOKIE_HEADER As String = "Set-Cookie"

    Public Function GetAuthenticatedWebClient()

        ' local variables
        Dim cookies As CookieContainer

        Dim request As HttpWebRequest
        Dim response As HttpWebResponse
        Dim responseString As String
        Dim tags As NameValueCollection
        Dim url As String

        username = Environment.UserName
        Console.WriteLine("Logging in to Intranet as: " & username)
        'TODO Uncomment password input
        Console.Write("Password: ")
        password = Console.ReadLine()
        Stop

        ' Step 1: Open connection to Protected URI.  This will be intercepted and redirected
        ' to SiteMinder login screen.

        url = PROTECTED_URL

        Debug.WriteLine("Step 1: Requesting page @" & url)
        request = WebRequest.Create(url)
        request.AllowAutoRedirect = False

        ' Using blocks would normally be used for HttpWebResponse objects but
        ' are omitted here for clarity
        response = request.GetResponse
        ShowResponse(response)


        ' Step 2: Get the redirection location

        ' make sure we have a valid response
        If response.StatusCode <> HttpStatusCode.Found Then
            Throw New ApplicationException
        End If

        url = response.Headers("Location")


        ' Step 3: Open a connection to the redirect and load the login form, 
        ' from this screen we will capture the required form fields.

        Debug.WriteLine("Step 3: Requesting page @" & url)
        request = WebRequest.Create(url)
        request.AllowAutoRedirect = False

        response = request.GetResponse

        responseString = Response2String(response)
        ShowResponse(responseString)
        tags = GetTags(responseString)


        ' Step 4: Create the form data to post back to the server

        ' Two ContentTypes are valid for sending POST data.  The default is
        ' application/x-www-form-urlencoded and form data is formatted similar
        ' to a typical querystring...
        ' 
        ' Forms submitted with this content type must be encoded as follows:
        ' 
        ' Control names and values are escaped. Space characters are replaced by `+', 
        ' and then reserved characters are escaped as described in [RFC1738], section 2.2: 
        ' Non-alphanumeric characters are replaced by `%HH', a percent sign and 
        ' two hexadecimal digits representing the ASCII code of the character. 
        ' Line breaks are represented as "CR LF" pairs (i.e., `%0D%0A'). 
        ' 
        ' The control names/values are listed in the order they appear in the document. 
        ' The name is separated from the value by `=' and name/value pairs are separated 
        ' from each other by `&'. 
        ' 
        ' The alternative is multipart/form-data...
        ' 
        ' for more information, see "http://www.w3.org/TR/html401/interact/forms.html#h-17.13.4.1"

        Dim postData As String

        postData = ""
        For Each inputName As String In tags.Keys
            If inputName.Length > 0 Then
                If inputName.Substring(0, 2).ToLower = "sm" Then
                    postData &= inputName & "=" & HttpUtility.UrlEncode(tags(inputName)) & "&"
                End If
            End If
        Next
        postData += "postpreservationdata=&"
        postData += "USER=" + HttpUtility.UrlEncode(username) & "&"
        postData += "PASSWORD=" + HttpUtility.UrlEncode(password)


        ' Step 5: Submit the data back to SiteMinder

        Debug.WriteLine("Step 5: Requesting page @" & url)
        request = WebRequest.Create(url)

        cookies = New CookieContainer
        request.CookieContainer = cookies
        request.ContentType = FORM_CONTENT_TYPE
        request.ContentLength = postData.Length
        request.Method = POST_METHOD
        request.AllowAutoRedirect = False   ' Important: we need to handle the redirect ourselves

        ' actually send the request
        Dim sw As StreamWriter = New StreamWriter(request.GetRequestStream())
        sw.Write(postData)
        sw.Flush()
        sw.Close()


        ' Step 6: Important to get the cookies here (they will include the SMSESSION cookie)

        Debug.WriteLine("Step 6: Response from @" & url)
        response = request.GetResponse
        responseString = Response2String(response)
        ShowResponse(responseString)

        ' get the cookies we need...
        ' either by parsing the HTTP headers...
        'cookies = ParseHeadersForCookies(response)

        ' or directly from the Cookies property of the response object
        For Each c As Cookie In response.Cookies
            Debug.WriteLine("cookie=" & c.Name)
        Next
        cookies.Add(response.Cookies)

        ' get the page we should hit next (should be Cognos)
        Dim newUrl As String = response.Headers("Location")


        ' Step 7: Persist the cookie

        Debug.WriteLine("Step 7: Persisting the SiteMinder info in a cookie")
        SetPersistentCookies(cookies)


        ' Step 8: Access protected resource here

        'Debug.WriteLine("Step 8: Access protected resource now")
        ' pause here
        'Stop

        ' Step 9: Delete the cookie from cache when done

        'Debug.WriteLine("Step 9: Deleting the SiteMinder info cookie")
        'DeletePersistentCookies(cookies)

        Return New CookieAwareWebClient(cookies)

    End Function

    Public Sub SetPersistentCookies(ByVal cookies As CookieContainer)

        ' local variables
        Dim cookie As Cookie
        Dim cookieString As String = ""
        Dim expireDate As DateTime
        Dim rc As Boolean

        ' initialize variables
        expireDate = DateTime.Now.ToUniversalTime.AddMinutes(COOKIE_TIMEOUT_MINUTES)

        For Each cookie In cookies.GetCookies(New Uri(PROTECTED_DOMAIN_URI))
            cookieString = cookie.Name & "=" & cookie.Value & "; expires = " & expireDate.ToString("ddd, dd-MMM-yyyy HH:mm:ss 'GMT'") & ";"
            rc = InternetSetCookie(PROTECTED_DOMAIN_URI, Nothing, cookieString)
        Next

    End Sub

    Public Sub DeletePersistentCookies(ByVal cookies As CookieContainer)

        Dim cookie As Cookie
        Dim cookieString As String = ""
        Dim rc As Boolean

        For Each cookie In cookies.GetCookies(New Uri(PROTECTED_DOMAIN_URI))
            ' the cookie will be deleted from the persistent store if it has 
            ' no expiration date or it has an expiration date in the past

            ' pick one...
            'cookieString = cookie.Name & "=" & cookie.Value & "; expires = Sat, 01-Jan-2000 00:00:00 GMT;"
            cookieString = cookie.Name & "=" & cookie.Value & ";"

            rc = InternetSetCookie(PROTECTED_DOMAIN_URI, Nothing, cookieString)
        Next

    End Sub

    ''' <summary>
    ''' Parses the HTML response and returns a collection of INPUT tags
    ''' </summary>
    ''' <param name="responseString"></param>
    ''' <returns></returns>
    ''' <remarks>This routine uses a WebBrowser control and HtmlDocument to
    ''' facilitate the parsing of the HTML.  However, any simple string parsing
    ''' algorithm would suffice.</remarks>
    Public Function GetTags(ByVal responseString As String) As NameValueCollection

        ' local variables
        Dim doc As HtmlDocument
        Dim hCol As HtmlElementCollection
        Dim webBrowser As WebBrowser
        Dim tags As NameValueCollection

        ' initialize variables
        tags = New NameValueCollection
        webBrowser = New WebBrowser
        webBrowser.Visible = True

        ' load the response HTML into a document object
        webBrowser.Navigate("about:blank")
        webBrowser.Document.Write(responseString)
        doc = webBrowser.Document
        hCol = doc.GetElementsByTagName("input")

        ' get the 'input' tags
        For Each item As HtmlElement In hCol

            Dim i As Integer
            Dim t As String = item.OuterHtml
            Dim typ As String
            Dim val As String

            ' ignore inputs with type=button
            i = t.IndexOf("type=")
            If i <> -1 Then
                typ = t.Substring(i + 5)
                i = typ.IndexOf(" "c)
                typ = typ.Substring(0, i)

                If typ.ToLower = "button" Then Continue For
            End If

            i = t.IndexOf("value=")
            If i = -1 Then
                val = String.Empty
            Else
                val = t.Substring(i + 6)
                i = val.IndexOf(" "c)
                val = val.Substring(0, i)
            End If

            Debug.WriteLine("adding tag=" & item.Name & ", value=" & val)
            tags.Add(item.Name, val)

        Next

        Return tags

    End Function

    ''' <summary>
    ''' Checks an HttpWebResponse for cookies to be set and adds them to the internal 
    ''' CookieContainer.
    ''' </summary>
    ''' <param name="res">The response to be checked for cookies</param>
    ''' <returns></returns>
    ''' <remarks>This function is necessary only when the developer specifically wants
    ''' to inspect the headers directly.  The cookie container can be obtained directly 
    ''' from the response using the <see cref="HttpWebResponse.Cookies"/> property.</remarks>
    Public Function ParseHeadersForCookies(ByVal res As HttpWebResponse) As CookieContainer

        Dim cc As CookieContainer = New CookieContainer

        For Each header As String In res.Headers

            If header.Equals(SET_COOKIE_HEADER) Then
                Console.WriteLine("header={0}", header)

                Dim setCookies() As String = res.Headers.GetValues(header)

                For i As Integer = 0 To setCookies.Length - 1
                    Dim setCookie As String = setCookies(i)

                    ' handle the extra comma in the expires attribute
                    Dim j As Integer = setCookie.IndexOf("expires")
                    If j <> -1 Then
                        i += 1
                        setCookie &= setCookies(i)
                    End If

                    ' parse the cookie and update its status in the collection
                    Dim equalsPos As Integer = setCookie.IndexOf("="c)
                    If equalsPos <> -1 Then
                        Dim name As String = setCookie.Substring(0, equalsPos)
                        Dim value As String = setCookie.Substring(equalsPos + 1, setCookie.IndexOf(";"c) - equalsPos - 1)

                        ' need to get the cookie's domain
                        Dim domain As String = ""
                        Dim domainPos As Integer = setCookie.IndexOf("domain")

                        If domainPos <> -1 Then
                            Dim endPos As Integer = setCookie.IndexOf(";"c, domainPos)
                            Dim eqPos As Integer = setCookie.IndexOf("="c, domainPos)

                            If endPos <> -1 Then
                                domain = setCookie.Substring(eqPos + 1, endPos - eqPos - 1)
                            Else
                                domain = setCookie.Substring(eqPos + 1)
                            End If

                            ' add/Update the cookie
                            cc.Add(New Cookie(name, value, "/", domain))
                        Else
                            ' add/Update the cookie
                            cc.Add(New Cookie(name, value, "/", ".metlife.com"))
                        End If

                    End If

                Next
            End If
        Next

        Return cc

    End Function

    Public Function Response2String(ByVal response As HttpWebResponse) As String

        Dim dataStream As Stream
        Dim reader As StreamReader
        Dim responseString As String

        ' Get the stream containing content returned by the server.
        dataStream = response.GetResponseStream()
        ' Open the stream using a StreamReader for easy access.
        reader = New StreamReader(dataStream)
        ' Read the content.
        responseString = reader.ReadToEnd()

        Return responseString

    End Function

    Public Sub ShowResponse(ByVal response As HttpWebResponse)

        Dim dataStream As Stream
        Dim reader As StreamReader
        Dim responseFromServer As String

        Debug.WriteLine("--------------")
        Debug.WriteLine(response.StatusDescription)
        ' Get the stream containing content returned by the server.
        dataStream = response.GetResponseStream()
        ' Open the stream using a StreamReader for easy access.
        reader = New StreamReader(dataStream)
        ' Read the content.
        responseFromServer = reader.ReadToEnd()
        ' Display the content.
        Debug.WriteLine(responseFromServer)

    End Sub

    Public Sub ShowResponse(ByVal responseString As String)

        Debug.WriteLine("--------------")
        ' Display the content.
        Debug.WriteLine(responseString)

    End Sub

End Module