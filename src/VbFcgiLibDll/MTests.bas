Attribute VB_Name = "MTests"
Option Explicit

Public Sub TestBuilderHtml()
   Dim x As New CBuilderHtml
   Dim l_TagIndex As Long
   
   x.AppendDocType
   x.OpenTags "html"
   l_TagIndex = x.OpenTags("head")
   x.AppendWithTag "Test page", "title"
   x.CloseOpenedTagsToIndex l_TagIndex
   x.Append vbNewLine
   
   x.OpenTags "body"
   l_TagIndex = x.OpenTags("table", "tr")
   x.AppendWithTag "This is a test & stuff.", "td"
   x.CloseLastOpenedTag
   x.OpenTags "tr"
   x.AppendWithTag "This is a test2.", "td"
   x.CloseOpenedTagsToIndex l_TagIndex
   
   x.OpenHyperlinkTag "http://www.statslog.com"
   x.CloseAllOpenedTags ' Optional, calling Finished will also take care of this.
   
   x.Finished contentencoding_UTF16_LE
   
   Debug.Print x.Content
End Sub
