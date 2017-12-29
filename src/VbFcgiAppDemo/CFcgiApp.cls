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

Private WithEvents mo_Vbml As VbFcgiLib.CBuilderTemplate
Attribute mo_Vbml.VB_VarHelpID = -1

Private Sub IFcgiApp_ProcessRequest(po_Request As VbFcgiLib.CFcgiRequest, po_Response As VbFcgiLib.CFcgiResponse)
   Dim l_StartedBuildAt As Double
   Dim ii As Long
   
   Dim lo_Header As vbRichClient5.cStringBuilder
   Dim lo_IBuilder As VbFcgiLib.IBuilder
   Dim lo_Html As VbFcgiLib.CBuilderHtml
   Dim lo_File As VbFcgiLib.CBuilderFile
   Dim lo_Json As VbFcgiLib.CBuilderJson
   Dim l_Title As String
   Dim l_SubTitle As String
   Dim l_SubTitleExample As String
   Dim l_SubTitleExampleUrl As String
   Dim l_VisitCount As Long
   Dim l_HttpHost As String
   Dim l_TagIndex As Long
   Dim l_Filename As String
   Dim l_Path As String
   
   On Error GoTo ErrorHandler

   ' Make sure that FCGI parameters are complete built and the Downstream FCGI object has been set
   ' otherwise raise fcgierr_NotReadyForResponse
   ' Just a sanity check - this should never happen
   If po_Request.Fcgi.Params.State <> paramstate_Built Then Err.Raise fcgierr_NotReadyForResponse, , "FCGI Parameters incomplete."
   
   l_StartedBuildAt = libRc5Factory.C.HPTimer

   Set mo_FcgiParams = po_Request.Fcgi.Params
   Set mo_FcgiStdin = po_Request.Fcgi.Stdin
   
   l_HttpHost = po_Request.Fcgi.Params.ValueByEnum(stdparam_HttpHost)
   
   ' Check for VBML template request
   l_Path = po_Request.Fcgi.Params.ValueByEnum(stdparam_PathInfo)
   Select Case LCase$(stringRemoveWhitespace(libFso.GetFileExtension(l_Path)))
   Case "vbml"
      ' Request is for a parsed VBML file
      Set mo_Vbml = po_Response.Builders.Builder(builder_Template)
      mo_Vbml.ParseFile pathBin & "vbml" & po_Request.Fcgi.Params.ValueByEnum(stdparam_PathInfo)
      mo_Vbml.Finish contentencoding_UTF8
      
   Case Else
      ' Check for file download request
      If po_Request.Http.QueryParameters.Exists("download") Then
         ' Requester wants to download a file
         
         l_Filename = po_Request.Http.QueryParameters("download")
         
         If InStr(1, l_Filename, "\") > 0 Or InStr(1, l_Filename, "/") Then
            Err.Raise 70, , "Permission denied for parent folders and sub-folders."
         End If
         
         Set lo_File = po_Response.Builders.Builder(builder_File)
         Set lo_IBuilder = lo_File
         
         lo_File.OpenFile pathBin & "downloads\" & l_Filename, po_Request.Fcgi.Params.ValueByEnum(stdparam_HttpIfNoneMatch)
          
         lo_IBuilder.Finish
          
      ' Check for JSON request
      ElseIf po_Request.Http.QueryParameters.Exists("getjson") Then
         Set lo_Json = po_Response.Builders.Builder(builder_Json)
         
         lo_Json.Initialize Nothing ' Initialize to empty collection
         ' Add some key & value pairs
         lo_Json.IJsonObject.AddJsonObjectByKeyValuePairs "status", 200, "message", "OK"
         lo_Json.Finish contentencoding_RecommendedDefault
         
      Else
         ' Build the HTML portion of the HTTP response
         
         ' Initialize the HTML builder/helper
         Set lo_Html = po_Response.Builders.Builder(builder_Html)
         
         With lo_Html
            .AppendDocType htmldoctype_Html5
            
            .OpenTags "html"
            
            l_TagIndex = .OpenTags("head")
                  
            ' *** START DEMONSTRATION OF QUERY PARAMETER HANDLING
            If po_Request.Http.QueryParameters.Exists("title") Then
               l_Title = lo_Html.EncodeHtmlEntities(po_Request.Http.QueryParameters("title"))
            End If
            ' *** END DEMONSTRATION OF QUERY PARAMETER HANDLING
            If stringIsEmptyOrWhitespaceOnly(l_Title) Then
               l_Title = "vbFcgi Demo App"
               l_SubTitle = "Pass a ""title"" query to change the title of this page."
               l_SubTitleExampleUrl = "http://" & l_HttpHost & po_Request.Fcgi.Params.ValueByEnum(stdparam_ScriptName) & "?title=Greetings from planet earth!"
            End If
            
            .AppendWithTag l_Title, "title"
            
            .CloseOpenedTagsToIndex l_TagIndex  ' Close <head> tag
            
            ' Build BODY
            .OpenTags "body"
            
            .AppendWithTag l_Title, "h1"
         
            If Not stringIsEmptyOrWhitespaceOnly(l_SubTitle) Then
               .AppendWithTag l_SubTitle, "h2"
                                          
               l_TagIndex = .OpenTags("p")
               .Append "Example: "
               .OpenHyperlinkTag l_SubTitleExampleUrl
               .Append .EncodeHtmlEntities(l_SubTitleExampleUrl)
               .CloseOpenedTagsToIndex l_TagIndex ' Close up to H3 tag
            End If
         
            l_TagIndex = .OpenTags("p", "b")
            .Append "<a href='https://www.github.com/jpbro/VbFcgi'>Learn more about VbFcgi on GitHub.</a>"
            .CloseOpenedTagsToIndex l_TagIndex  ' Close B and P tags
         
            ' *** START DEMONSTRATION OF COOKIES
            l_TagIndex = .OpenTagWithAttributes("p", , , "color: orange; font-weight: bold;")
            If po_Request.Http.Cookies.Exists("visits") Then
               On Error Resume Next
               l_VisitCount = po_Request.Http.Cookies.CookieByKey("visits").Value
               On Error GoTo ErrorHandler
               
               If l_VisitCount = 0 Then
                  ' Bad cookie value!
                  .Append "Hey! Have you been mucking about with your cookies?"
               Else
                  ' Display number of visits
                  .Append4 "You have previously visited this page ", l_VisitCount, " time", IIf(l_VisitCount <> 1, "s.", ".")
               End If
            Else
               ' First visit
               .Append "This is your first visit, pleased to meet you!"
            End If
            .CloseOpenedTagsToIndex l_TagIndex  ' Close P tag
                  
            ' Increment "Visits" cookie
            po_Request.Http.Cookies.AddOrReplaceCookie "visits", l_VisitCount + 1
            ' *** END DEMONSTRATION OF COOKIES
         
            ' Add download links to test CBuilderFile class
            Dim l_DownloadUrl As String
            
            l_DownloadUrl = "http://" & l_HttpHost & po_Request.Fcgi.Params.ValueByEnum(stdparam_ScriptName) & "?download="
            
            l_TagIndex = .OpenTags("p")
            .OpenHyperlinkTag l_DownloadUrl & "test.pdf", "_blank"
            .Append "Download a PDF"
            .CloseOpenedTagsToIndex l_TagIndex ' Close P tag
            
            l_TagIndex = .OpenTags("p")
            .OpenHyperlinkTag l_DownloadUrl & "test.png", "_blank"
            .Append "Download a PNG"
            .CloseOpenedTagsToIndex l_TagIndex ' Close P tag
                     
            ' End of download test
            
            .Append4 "<p>", "The current date & time on the server is: ", Now, "</p>"
            .AppendWithTag "VbFcgi is " & ChrW$(&HAAA&) & ChrW$(&HE01&) & ChrW$(&H671&) & ChrW$(&H188&) & ChrW$(&H47B&) & ChrW$(&H257&) & ChrW$(&HFEC9&) & " capable via UTF-8!", "p"
          
            ' Build FCGI Parameters table
            .AppendWithTag "FCGI Parameters received from downstream Webserver:", "h2"
            .OpenTags "table"
            For ii = 1 To mo_FcgiParams.Count - 1
               .OpenTags "tr"
               .AppendWithTag mo_FcgiParams.KeyByIndex(ii), "td"
               .AppendWithTag mo_FcgiParams.ValueByIndex(ii), "td"
               .CloseLastOpenedTag  ' Close tr tag
            Next ii
            .CloseLastOpenedTag  ' Close table tag
            
            ' Build STDIN info
            .AppendWithTag "FCGI STDIN Content Length: " & mo_FcgiStdin.ContentLength, "p"
            
            ' Build response time info
            .AppendWithTag "Response Build Time: " & Format$(libRc5Factory.C.HPTimer - l_StartedBuildAt, "0.00000") & " seconds", "p"
            
            ' Close all open tags in reverse order
            .CloseAllOpenedTags
                  
            ' Add cookies to HTTP header
            Set lo_IBuilder = lo_Html
            lo_IBuilder.HttpHeader.Append po_Request.Http.Cookies.CookieHeaders
            
            ' Finish building HTML, encode as UTF-8 and write it back downstream
            .Finish contentencoding_UTF8
         End With
      End If
   End Select
   
   Exit Sub

ErrorHandler:
   Dim l_ErrNum As Long
   Dim l_ErrDesc As String
   
   l_ErrNum = Err.Number
   l_ErrDesc = Err.Description
   
   ' Send error information downstream
   po_Response.Error l_ErrNum, l_ErrDesc
End Sub

Private Sub mo_Vbml_FoundTag(ByVal p_Tag As String, po_Replacement As VbFcgiLib.CWebStringTemplate, ByRef p_DoNotReplace As Boolean)
   Dim ii As Long
   Dim l_OpenTagIndex As Long
            
   Select Case UCase$(p_Tag)
   Case "TITLE"
      ' Generate a dynamic title
      po_Replacement = "It is " & Format$(Now, "MMMM D, YYYY at H:NN ampm")
   
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


