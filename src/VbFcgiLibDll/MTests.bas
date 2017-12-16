Attribute VB_Name = "MTests"
Option Explicit

Public Sub TestHtml()
   Dim y As New CBuilders
   Dim x As CBuilderHtml
   Dim l_TagIndex As Long
   
   Set x = y.Builder(builder_Html)
   
   x.AppendDocType htmldoctype_Html5
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
   
   x.Finish contentencoding_USASCII
   
   Debug.Print y.Builder.Length
   
   Debug.Print stringIso88591ToVb(x.BuilderInterface.HttpHeader.Content(True))
End Sub

Public Sub TestCollection()
   Dim x As vbRichClient5.cCollection
   
   Set x = libRc5Factory.C.Collection(False)
   
   x.Add "AD"
   
   Debug.Print x.KeyByIndex(0) = ""
End Sub
   
