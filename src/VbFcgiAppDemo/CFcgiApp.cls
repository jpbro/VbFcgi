VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFcgiApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Copyright (c) 2017 Jason Peter Brown <jason@bitspaces.com>
'
' MIT License
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Implements VbFcgiLib.IFcgiApp

' Every VbFcgiApp.dll must inclde a CFcgiApp class as a minimum.
' The CFcgiApp class must include the "Implements VbFcgiLib.IFcgiApp" directive in the General section.
' The CFcgiApp class must include all members of the IFcgiApp interface. They are:
'        IFcgiApp_BuildResponse and IFcgiApp_DownstreamObject
' After instatiation, your CFcgiApp class will be passed a VbFcgiLib.CFcgiDownstream object to IFcgiApp_DownstreamObject
'     The Downstream object should be stored and its methods should be called for sending responses and raising errors.
'     The Downstream object class methods are:
'        WriteBytes, Error, and Finished

' This is a very small sample class for building responses to FCGI requests.
' It simply spits back an HTML page with the following:
' Unicode/UTF-8 test, generated date/time, list of received parameters, approx. time taken to build response
' In reality you would parse out the QUERY_STRING and HTTP_COOKIES parameters where available
' And respond accordingly, but this is beyond the scope of this demo.

Private mo_FcgiParams As VbFcgiLib.CFcgiParams
Private mo_FcgiStdin As VbFcgiLib.CFcgiStdIn
Private mo_HttpParams As VbFcgiLib.CHttpQueryParams

Private WithEvents mo_Vbml As VbFcgiLib.CBuilderTemplate
Attribute mo_Vbml.VB_VarHelpID = -1

Private mo_FormImages As VbFcgiApp.frmImages    ' Legacy app form

Private m_Title As String
Private m_HttpHost As String

Private Sub Class_Initialize()
   Randomize
End Sub

Private Function HasPathlessHttpQueryParam() As Boolean
   Dim ii As Long
   
   If mo_HttpParams.Exists("download") Then
      HasPathlessHttpQueryParam = True
   
   ElseIf mo_HttpParams.Exists("imagebyindex") Then
      HasPathlessHttpQueryParam = True
   
   ElseIf mo_HttpParams.Exists("showparams") Then
      HasPathlessHttpQueryParam = True
   
   ElseIf mo_HttpParams.Exists("getbytes") Then
      HasPathlessHttpQueryParam = True
   
   Else
      For ii = 0 To mo_HttpParams.KeyCount - 1
         If LCase$(Left$(mo_HttpParams.KeyByIndex(ii), Len("json_"))) = "json_" Then
            HasPathlessHttpQueryParam = True
            
            Exit Function
         End If
      Next ii
   End If
End Function

Private Sub IFcgiApp_ProcessRequest(po_Request As VbFcgiLib.CFcgiRequest, po_Response As VbFcgiLib.CFcgiResponse)
         Dim l_StartedBuildAt As Double
         Dim ii As Long
         Dim lo_Header As vbRichClient5.cStringBuilder
         Dim lo_IBuilder As VbFcgiLib.IBuilder
         Dim lo_Html As VbFcgiLib.CBuilderHtml
         Dim lo_File As VbFcgiLib.CBuilderFile
         Dim lo_Json As VbFcgiLib.CBuilderJson
         Dim l_SubTitle As String
         Dim l_SubTitleExample As String
         Dim l_SubTitleExampleUrl As String
         Dim l_VisitCount As Long
         Dim l_TagIndex As Long
         Dim l_TagIndex2 As Long
         Dim l_Filename As String
         Dim l_Path As String
         Dim lo_Cnn As vbRichClient5.cConnection
         Dim lo_Cmd As vbRichClient5.cCommand
         Dim lo_Rs As vbRichClient5.cRecordset
         Dim lo_CairoSurface As vbRichClient5.cCairoSurface

10       On Error GoTo ErrorHandler

20       Debug.Assert False   ' For test/simulation mode

         ' Make sure that FCGI parameters are complete built and the Downstream FCGI object has been set
         ' otherwise raise fcgierr_NotReadyForResponse
         ' Just a sanity check - this should never happen
30       If po_Request.Fcgi.Params.State <> paramstate_Built Then Err.Raise fcgierr_NotReadyForResponse, , "FCGI Parameters incomplete."

40       l_StartedBuildAt = libRc5Factory.C.HPTimer

50       Set mo_FcgiParams = po_Request.Fcgi.Params
60       Set mo_HttpParams = po_Request.Http.QueryParameters
70       Set mo_FcgiStdin = po_Request.Fcgi.Stdin

80       m_HttpHost = po_Request.Fcgi.Params.ValueByEnum(stdparam_HttpHost)

         ' Check if we should use a default vbml file
90       l_Path = po_Request.Fcgi.Params.ValueByEnum(stdparam_PathInfo)

100      apiOutputDebugString "Request path: " & l_Path

110      If stringIsEmptyOrWhitespaceOnly(libFso.GetFileExtension(l_Path)) Then
120         If Not HasPathlessHttpQueryParam Then
130            l_Path = l_Path & "\index.vbml"
140         End If
150      End If

160      apiOutputDebugString "Normalized request path: " & l_Path

         ' *** START DEMONSTRATION OF HTTP QUERY PARAMETER HANDLING
170      If po_Request.Http.QueryParameters.Exists("title") Then
180         m_Title = po_Request.Http.QueryParameters("title")
190      Else
200         m_Title = ""
210      End If
         ' *** END DEMONSTRATION OF HTTP QUERY PARAMETER HANDLING

         Dim l_ContentType As String
         
         l_ContentType = po_Request.Fcgi.Params.ValueByEnum(stdparam_ContentType)
         If Not stringIsEmptyOrWhitespaceOnly(l_ContentType) Then
            If LCase$(Split(l_ContentType, ";")(0)) = "application/json" Then
               ' JSON body content received
               Set lo_Json = po_Response.Builders.Builder(builder_Json)
               
               lo_Json.Initialize po_Request.Fcgi.Stdin.Content
               Set lo_IBuilder = lo_Json
               
               lo_Json.Finish
               
               GoTo Cleanup
            End If
         End If

         ' Check for VBML template request
220      Select Case LCase$(stringRemoveWhitespace(libFso.GetFileExtension(l_Path)))
         Case "vbml"
            ' Request is for a parsed VBML file
230         apiOutputDebugString "Getting Template Builder helper."

240         Set mo_Vbml = po_Response.Builders.Builder(builder_Template)
250         Set lo_IBuilder = mo_Vbml

260         l_Path = pathBin & "vbml" & l_Path

270         apiOutputDebugString "Parsing file: " & l_Path

280         mo_Vbml.ParseFile l_Path

290         apiOutputDebugString "Parsed file: " & l_Path

300         mo_Vbml.Finish contentencoding_UTF8

310         apiOutputDebugString "Finished parsing VBML template."

320      Case Else
            ' Check for file download request
330         If po_Request.Http.QueryParameters.Exists("download") Then
               ' Requester wants to download a file

340            l_Filename = po_Request.Http.QueryParameters("download")

350            If InStr(1, l_Filename, "\") > 0 Or InStr(1, l_Filename, "/") Then
360               Err.Raise 70, , "Permission denied for parent folders and sub-folders."
370            End If

380            Set lo_File = po_Response.Builders.Builder(builder_File)

390            lo_File.OpenFile pathBin & "downloads\" & l_Filename, po_Request.Fcgi.Params.ValueByEnum(stdparam_HttpIfNoneMatch)

               ' Check for image request from VB6 form
400         ElseIf po_Request.Http.QueryParameters.Exists("imagebyindex") Then
               ' Get image by VB6 form list view index
               ' We must make sure NOT to show the form!

               Dim l_ImageIndex As Long
               Dim l_Etag As String
               Dim la_Jpg() As Byte

410            apiOutputDebugString "Requester wants an ImageList by index from frmImages."

               ' Check if we have the file cached already
420            If po_Request.Fcgi.Params.ExistsByEnum(stdparam_HttpIfNoneMatch) Then
430               l_Etag = po_Request.Fcgi.Params.ValueByEnum(stdparam_HttpIfNoneMatch)
440               l_Path = libFso.GetTmpPath & l_Etag & ".jpg"
450               If Not libFso.FileExists(l_Path) Then
460                  l_Path = ""
470               End If
480            End If

490            If l_Path = "" Then
                  ' Image file is not already cached, so build & cache it.

500               l_ImageIndex = po_Request.Http.QueryParameters("imagebyindex")

510               apiOutputDebugString "Requested image index: " & l_ImageIndex

                  ' Create an instance of our form and load it
                  ' But do NOT show it!
520               Set mo_FormImages = New VbFcgiApp.frmImages
530               Load mo_FormImages

                  ' Interact with the VB6 form to select the appropriate image index
                  ' from the ListBox (which will in turn fire the List1_Click even
                  ' Where the appropriate image is loaded into the PictureBox
                  ' This simulates a "legacy" VB6 program.
540               If (l_ImageIndex > 0) And (l_ImageIndex < mo_FormImages.List1.ListCount) Then
550                  mo_FormImages.List1.ListIndex = l_ImageIndex
560               End If

                  ' Convert the image in the VB6 form PictureBox into a JPG
570               Set lo_CairoSurface = imageStdPictureToCairoSurface(mo_FormImages.Picture1.Picture)

580               lo_CairoSurface.WriteContentToJpgByteArray la_Jpg

                  ' Get the SHA256 hash of the image
590               l_Etag = libRc5Factory.C.Crypt.SHA256(la_Jpg, True)

600               l_Path = libFso.GetTmpPath & l_Etag & ".jpg"

610               libFso.WriteByteContent l_Path, la_Jpg
620               Erase la_Jpg
630            End If

               ' Use the file response builder helper to send the JPG downstream
640            Set lo_File = po_Response.Builders.Builder(builder_File)

650            lo_File.OpenFile l_Path, po_Request.Fcgi.Params.ValueByEnum(stdparam_HttpIfNoneMatch)

               ' Check for JSON request
660         ElseIf po_Request.Http.QueryParameters.Exists("json_gettime") Then
670            Set lo_Json = po_Response.Builders.Builder(builder_Json)
680            Set lo_IBuilder = lo_Json

               ' Add some key & value pairs
690            If Int(Rnd * 10) = 1 Then
                  ' Simulate random error
700               lo_Json.IJsonObject.AddJsonObjectByKeyValuePairs "status", "error", "error_message", "Simulated Error!"
710            Else
                  ' Return OK status and current server-local datetime
720               lo_Json.IJsonObject.AddJsonObjectByKeyValuePairs "status", "ok", "time", Now
730            End If

740            lo_Json.Finish

750         ElseIf mo_HttpParams.Exists("json_getdata") Then
               ' Simulate getting data from a database

               ' Create an in-memory database
760            Set lo_Cnn = libRc5Factory.C.Connection(, DBCreateInMemory)
770            lo_Cnn.Execute "CREATE TABLE mytable (code TEXT, value1 INTEGER, value2 INTEGER, value3 REAL)"

               ' Build random table data
780            Set lo_Cmd = lo_Cnn.CreateCommand("INSERT INTO mytable (code, value1, value2, value3) VALUES (?,?,?,?)")
790            For ii = 0 To 25
800               With lo_Cmd
810                  .SetAllParamsNull

820                  .SetText 1, Chr$(65 + ii)
830                  .SetInt32 2, Int(Rnd * 100)
840                  .SetInt32 3, Int(Rnd * 100)
850                  .SetDouble 4, Rnd

860                  .Execute
870               End With
880            Next ii

               ' Get data into recordset
890            Set lo_Rs = lo_Cnn.OpenRecordset("SELECT * FROM mytable")

               ' Build JSON response
900            Set lo_Json = po_Response.Builders.Builder(builder_Json)

910            lo_Json.Initialize lo_Rs.ToJSONUTF8
920            Set lo_IBuilder = lo_Json

930            lo_Json.Finish

            ' Check for binary byte request
            ElseIf mo_HttpParams.Exists("getbytes") Then
               Dim la_Content() As Byte
               
               Set lo_Rs = libRc5Factory.C.Recordset
               
               la_Content = lo_Rs.Content
               
               po_Response.WriteBytes stringVbToUtf8("Content-Length: " & arraySize(la_Content) & vbNewLine)
               po_Response.WriteBytes stringVbToUtf8("Content-Type: application/octet-stream" & vbNewLine)
               po_Response.WriteBytes stringVbToUtf8("Content-Type: application/octet-stream" & vbNewLine)
               po_Response.WriteBytes stringVbToUtf8(vbNewLine)
               po_Response.WriteBytes la_Content
               po_Response.Finished
               
940         ElseIf mo_HttpParams.Exists("showparams") Then
               ' Initialize the HTML builder/helper
950            Set lo_Html = po_Response.Builders.Builder(builder_Html)

960            lo_Html.AppendDocType htmldoctype_Html5

970            With lo_Html
980               .OpenTags "html"
990               l_TagIndex = .OpenTags("head")

1000              .AppendWithTag "Request Parameters", "title"

1010              .CloseOpenedTagsToIndex l_TagIndex  ' Close <head> tag

                  ' Build BODY
1020              .OpenTags "body"

1030              .AppendWithTag "FCGI Params", "h1"

1040              l_TagIndex = .OpenTags("table")

1050              For ii = 0 To po_Request.Fcgi.Params.Count - 1
1060                 l_TagIndex2 = .OpenTags("tr")

1070                 .AppendWithTag po_Request.Fcgi.Params.KeyByIndex(ii), "td"
1080                 .AppendWithTag po_Request.Fcgi.Params.ValueByIndex(ii), "td"

1090                 .CloseOpenedTagsToIndex l_TagIndex2
1100              Next ii

1110              .CloseOpenedTagsToIndex l_TagIndex

1120              .OpenTags "table"

1130              For ii = 0 To po_Request.Http.QueryParameters.KeyCount - 1
1140                 l_TagIndex = .OpenTags("tr")

1150                 .AppendWithTag po_Request.Http.QueryParameters.KeyByIndex(ii), "td"
1160                 .AppendWithTag po_Request.Http.QueryParameters.ValuesByIndex(ii).ValueByIndex(0), "td"

1170                 .CloseOpenedTagsToIndex l_TagIndex
1180              Next ii

1190              .CloseAllOpenedTags

1200              .Finish
1210           End With

1220        Else
               ' Build the HTML portion of the HTTP response

               ' Initialize the HTML builder/helper
1230           Set lo_Html = po_Response.Builders.Builder(builder_Html)

1240           With lo_Html
1250              .AppendDocType htmldoctype_Html5

1260              .OpenTags "html"

1270              l_TagIndex = .OpenTags("head")

1280              If stringIsEmptyOrWhitespaceOnly(m_Title) Then
1290                 m_Title = "vbFcgi Demo App"
1300                 l_SubTitle = "Pass a ""title"" query to change the title of this page."
1310                 l_SubTitleExampleUrl = "http://" & m_HttpHost & po_Request.Fcgi.Params.ValueByEnum(stdparam_ScriptName) & "?title=Greetings from planet earth!"
1320              End If

1330              .AppendWithTag m_Title, "title"

1340              .CloseOpenedTagsToIndex l_TagIndex  ' Close <head> tag

                  ' Build BODY
1350              .OpenTags "body"

1360              .AppendWithTag m_Title, "h1"

1370              If Not stringIsEmptyOrWhitespaceOnly(l_SubTitle) Then
1380                 .AppendWithTag l_SubTitle, "h2"

1390                 l_TagIndex = .OpenTags("p")
1400                 .Append "Example: "
1410                 .OpenHyperlinkTag l_SubTitleExampleUrl
1420                 .Append .EncodeHtmlEntities(l_SubTitleExampleUrl)
1430                 .CloseOpenedTagsToIndex l_TagIndex   ' Close up to p tag
1440              End If

1450              l_TagIndex = .OpenTags("p")
1460              .OpenHyperlinkTag "http://" & m_HttpHost & po_Request.Fcgi.Params.ValueByEnum(stdparam_ScriptName) & "/template_demo.vbml"
1470              .Append "Click to visit a dynamically altered template page."
1480              .CloseOpenedTagsToIndex l_TagIndex   ' Close up to the p tag

1490              l_TagIndex = .OpenTags("p")
1500              .OpenHyperlinkTag "http://" & m_HttpHost & po_Request.Fcgi.Params.ValueByEnum(stdparam_ScriptName) & "?getjson=getjson"
1510              .Append "Click to get JSON data."
1520              .CloseOpenedTagsToIndex l_TagIndex   ' Close up to the p tag

1530              .Append "<hr>"

1540              l_TagIndex = .OpenTags("p", "b")
1550              .Append "<a href='https://www.github.com/jpbro/VbFcgi'>Learn more about VbFcgi on GitHub.</a>"
1560              .CloseOpenedTagsToIndex l_TagIndex  ' Close B and P tags

1570              .Append "<hr>"

                  ' *** START DEMONSTRATION OF COOKIES
1580              l_TagIndex = .OpenTagWithAttributes("p", , , "color: orange; font-weight: bold;")
1590              If po_Request.Http.Cookies.Exists("visits") Then
1600                 On Error Resume Next
1610                 l_VisitCount = po_Request.Http.Cookies.CookieByKey("visits").Value
1620                 On Error GoTo ErrorHandler

1630                 If l_VisitCount = 0 Then
                        ' Bad cookie value!
1640                    .Append "Hey! Have you been mucking about with your cookies?"
1650                 Else
                        ' Display number of visits
1660                    .Append4 "You have previously visited this page ", l_VisitCount, " time", IIf(l_VisitCount <> 1, "s.", ".")
1670                 End If
1680              Else
                     ' First visit
1690                 .Append "This is your first visit, pleased to meet you!"
1700              End If
1710              .CloseOpenedTagsToIndex l_TagIndex  ' Close P tag

                  ' Increment "Visits" cookie
1720              po_Request.Http.Cookies.AddOrReplaceCookie "visits", l_VisitCount + 1
                  ' *** END DEMONSTRATION OF COOKIES

                  ' Add download links to test CBuilderFile class
                  Dim l_DownloadUrl As String

1730              l_DownloadUrl = "http://" & m_HttpHost & po_Request.Fcgi.Params.ValueByEnum(stdparam_ScriptName) & "?download="

1740              l_TagIndex = .OpenTags("p")
1750              .OpenHyperlinkTag l_DownloadUrl & "test.pdf", "_blank"
1760              .Append "Download a PDF"
1770              .CloseOpenedTagsToIndex l_TagIndex   ' Close P tag

1780              l_TagIndex = .OpenTags("p")
1790              .OpenHyperlinkTag l_DownloadUrl & "test.png", "_blank"
1800              .Append "Download a PNG"
1810              .CloseOpenedTagsToIndex l_TagIndex   ' Close P tag

                  ' End of download test

1820              .Append4 "<p>", "The current date & time on the server is: ", Now, "</p>"
1830              .AppendWithTag "VbFcgi is " & ChrW$(&HAAA&) & ChrW$(&HE01&) & ChrW$(&H671&) & ChrW$(&H188&) & ChrW$(&H47B&) & ChrW$(&H257&) & ChrW$(&HFEC9&) & " capable via UTF-8!", "p"

                  ' Build FCGI Parameters table
1840              .AppendWithTag "FCGI Parameters received from downstream Webserver:", "h2"
1850              .OpenTags "table"
1860              For ii = 1 To mo_FcgiParams.Count - 1
1870                 .OpenTags "tr"
1880                 .AppendWithTag mo_FcgiParams.KeyByIndex(ii), "td"
1890                 .AppendWithTag mo_FcgiParams.ValueByIndex(ii), "td"
1900                 .CloseLastOpenedTag  ' Close tr tag
1910              Next ii
1920              .CloseLastOpenedTag  ' Close table tag

                  ' Build STDIN info
1930              .AppendWithTag "FCGI STDIN Content Length: " & mo_FcgiStdin.ContentLength, "p"

                  ' Build response time info
1940              .AppendWithTag "Response Build Time: " & Format$(libRc5Factory.C.HPTimer - l_StartedBuildAt, "0.00000") & " seconds", "p"

                  ' Close all open tags in reverse order
1950              .CloseAllOpenedTags

                  ' Add cookies to HTTP header
1960              Set lo_IBuilder = lo_Html
1970              lo_IBuilder.HttpHeader.Append po_Request.Http.Cookies.CookieHeaders

                  ' Finish building HTML, encode as UTF-8 and write it back downstream
1980              .Finish contentencoding_UTF8
1990           End With
2000        End If

2010     End Select

Cleanup:
         On Error GoTo ErrorHandler
         
         ' Check if we remembered to call Finish() on the IBuilder object (if set)
2020     If Not lo_IBuilder Is Nothing Then
2030        If Not lo_IBuilder.IsFinished Then
               ' We forgot to call Finish()!
               ' Call Finish with the recommended default
               ' content encoding of the IBuilder helper class.
2040           lo_IBuilder.Finish
2050        End If

2060        If envRunningInIde Then
               ' For debugging purposes, print the content that will be sent downstream
2070           Debug.Print stringUtf8ToVb(lo_IBuilder.Content)
2080        End If
2090     End If

2100     Cleanup

2110     Exit Sub

ErrorHandler:
         Dim l_ErrNum As Long
         Dim l_ErrDesc As String

2120     l_ErrNum = Err.Number
2130     l_ErrDesc = Err.Description & ", Line #" & Erl

2140     apiOutputDebugString "Error in " & App.EXEName & " #" & l_ErrNum & " " & l_ErrDesc

2150     Cleanup

         ' Send error information downstream
2160     po_Response.Error l_ErrNum, l_ErrDesc
End Sub

Private Sub Cleanup()
   ' Reset all module level variables to prevent leakage to subsequent requests

   On Error Resume Next

   m_Title = ""
   m_HttpHost = ""

   Set mo_FcgiParams = Nothing
   Set mo_FcgiStdin = Nothing
   Set mo_HttpParams = Nothing
   Set mo_Vbml = Nothing

   If Not mo_FormImages Is Nothing Then
      Unload mo_FormImages
      Set mo_FormImages = Nothing
   End If

   libRc5Factory.C.CleanupRichClientDll

   Err.Clear
End Sub

Private Sub mo_Vbml_FoundTag(ByVal p_Tag As String, po_Replacement As VbFcgiLib.CWebStringTemplate, ByRef p_DoNotReplace As Boolean)
   Dim ii As Long
   Dim l_OpenTagIndex As Long
   Dim lo_DirList As vbRichClient5.cDirList
   Dim l_DownloadUrl As String

   Select Case UCase$(p_Tag)
   Case "SCRIPTNAME"
      po_Replacement = mo_FcgiParams.ValueByEnum(stdparam_ScriptName)

   Case "TITLE"
      ' Generate a dynamic title
      po_Replacement = IIf(stringIsEmptyOrWhitespaceOnly(m_Title), "Back end coding for the web using VB6!", m_Title)

   Case "FILE_LIST"
      Set lo_DirList = libFso.GetDirList(pathBin & "downloads\", dlSortByNameLogically)

      With po_Replacement
         If lo_DirList.FilesCount = 0 Then
            po_Replacement = "No files available for download."
         Else
            l_DownloadUrl = "http://" & m_HttpHost & mo_FcgiParams.ValueByEnum(stdparam_ScriptName) & "?download="

            .SkipEncodeEntities = True

            For ii = 0 To lo_DirList.FilesCount - 1
               po_Replacement.OpenTags "p"
               po_Replacement.OpenHyperlinkTag l_DownloadUrl & lo_DirList.FileName(ii)
               po_Replacement.Append po_Replacement.EncodeHtmlEntities(lo_DirList.FileName(ii))
               po_Replacement.CloseAllOpenedTags
            Next ii
         End If

      End With

   Case "IMAGE_LIST"
      ' Get image list from legacy form
      Dim lo_Images As VbFcgiApp.frmImages

      ' Load but don't show the form!
      Set lo_Images = New VbFcgiApp.frmImages
      Load lo_Images

      po_Replacement.SkipEncodeEntities = True  ' We're returning HTML, so don't encode HTML entities

      ' Build the image list
      po_Replacement.OpenTagWithAttributes "select", , "imagelist"
      po_Replacement.Append "<option value='-1'>Select an image...</option>"

      For ii = 0 To lo_Images.List1.ListCount - 1
         po_Replacement.Append "<option value='" & ii & "'>"
         po_Replacement.Append po_Replacement.EncodeHtmlEntities(lo_Images.List1.List(ii))
         po_Replacement.Append "</option>"
      Next ii

      po_Replacement.CloseAllOpenedTags

      Unload lo_Images  ' Unload form

   Case "TABLE_001"
      ' Generate a random table of data, and tell the VBML template parser not to encode HTML entities on the assumption we've done it ourselves

      With po_Replacement
         .SkipEncodeEntities = True ' Turn off automatic entity encoding since we are building HTML

         .OpenTags "table"
         For ii = 0 To 5
            l_OpenTagIndex = .OpenTags("tr")

            .AppendWithTag Chr$(65 + ii), "td"
            .AppendWithTagAndAttributes Format$(Rnd * 100, "0.000"), "td", "right"
            .AppendWithTagAndAttributes Format$(Rnd * 100, "0.000"), "td", "right"
            .AppendWithTagAndAttributes Format$(Rnd * 100, "0.000"), "td", "right"

            .CloseOpenedTagsToIndex l_OpenTagIndex  ' Close the <tr> tag

            .Append vbNewLine
         Next ii

         .CloseAllOpenedTags
      End With

   Case Else
      ' Unknown tag, replace with a warning.
      apiOutputDebugString "Encountered unknown tag: " & p_Tag

      po_Replacement.SkipEncodeEntities = True ' Turn off automatic entity encoding since we are building HTML
      po_Replacement = "<b style='color: red'>UNKNOWN TAG WARNING:</b> <i>&laquo;" & p_Tag & "&raquo;</i>"

   End Select
End Sub

Private Function GenerateSessionId() As String
   libCrypt.GetHashedPassword ("this is a password")
End Function

