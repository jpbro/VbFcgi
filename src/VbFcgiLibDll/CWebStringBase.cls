VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWebStringBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private mo_StringBuilder As vbRichClient5.cStringBuilder
Private mo_Tags As vbRichClient5.cArrayList

Public Sub AppendDocType(Optional ByVal p_DocType As VbFcgiLib.e_HtmlDocType = htmldoctype_Html5, Optional ByVal p_CustomDocType As String)
   ' Append a standard DOCTYPE to the top of an HTML page
   ' Currently only HTML5 and custom DOCTYPE strings are supported.
   
   If mo_StringBuilder.Length > 0 Then Err.Raise 5, , "DOCTYPE must be the first thing you append to your HTML page."
   
   If p_DocType <> htmldoctype_Custom Then If Not stringIsEmptyOrWhitespaceOnly(p_CustomDocType) Then Err.Raise 5, , "Custom DOCTYPE must be empty for non-custom DOCTYPEs."
   
   Select Case p_DocType
   Case htmldoctype_Custom
      If stringIsEmptyOrWhitespaceOnly(p_CustomDocType) Then
         Err.Raise 5, , "Custom DOCTYPE required."
      End If
      
      mo_StringBuilder.AppendNL p_CustomDocType
   
   Case htmldoctype_Html5
      mo_StringBuilder.AppendNL "<!DOCTYPE html>"
   
   Case Else
      Err.Raise 5, , "Unknown DocType enum: " & p_DocType
   End Select
End Sub

Public Sub Append(ByVal p_Text As String)
   ' Append any string to the content
   
   mo_StringBuilder.Append p_Text
End Sub

Public Sub Append2(ByVal p_Text1 As String, ByVal p_Text2 As String)
   ' Append any 2 strings (one after the other) to the content
   
   With mo_StringBuilder
      .AppendNL p_Text1, p_Text2
   End With
End Sub

Public Sub Append3(ByVal p_Text1 As String, ByVal p_Text2 As String, ByVal p_Text3 As String)
   ' Append any 3 strings (one after the other) to the content
   
   With mo_StringBuilder
      .AppendNL p_Text1, p_Text2
      .Append p_Text3
   End With
End Sub

Public Sub Append4(ByVal p_Text1 As String, ByVal p_Text2 As String, ByVal p_Text3 As String, ByVal p_Text4 As String)
   ' Append any 4 strings (one after the other) to the content
   
   With mo_StringBuilder
      .AppendNL p_Text1, p_Text2
      .AppendNL p_Text3, p_Text4
   End With
End Sub

Public Sub AppendWithNewline(ByVal p_Text As String, Optional ByVal p_NewLine As String = vbNewLine)
   mo_StringBuilder.AppendNL p_Text, p_NewLine
End Sub

Public Sub AppendWithTag(ByVal p_Text As String, ByVal p_Tag As String, Optional ByVal p_EncodeEntities As Boolean = True)
   ' In one pass this method will:
   
   ' Append an HTML tag (don't pass the "<>" characters),
   ' Then Append any arbitrary text (optionally encoding HTML entities (default))
   ' Then append the closing tag.
   
   With mo_StringBuilder
      OpenTag p_Tag, False
      
      If p_EncodeEntities Then
         .Append htmlEscape(p_Text)
      Else
         .Append p_Text
      End If
         
      CloseTag p_Tag
   End With
End Sub

Public Sub AppendWithTagAndAttributes(ByVal p_Text As String, ByVal p_Tag As String, Optional ByVal p_TagClass As String, Optional ByVal p_TagId As String, Optional ByVal p_TagStyle As String, Optional ByVal p_EncodeEntities As Boolean = True)
   ' In one pass this method will:
   
   ' Append an HTML tag (don't pass the "<>" characters),
   ' Optionally appending a class='passedclass' attribute to the tag
   ' Optionally appending an id='passedid' attribute to the tag
   ' Optionally appending a style='passedstyles' attribute to the tag
   ' Then Append any arbitrary text (optionally encoding HTML entities (default))
   ' Then append the closing tag.
   
   OpenTagWithAttributes p_Tag, p_TagClass, p_TagId, p_TagStyle
   
   If p_EncodeEntities Then
      mo_StringBuilder.Append htmlEscape(p_Text)
   Else
      mo_StringBuilder.Append p_Text
   End If
   
   Me.CloseLastOpenedTag
End Sub

Public Sub AppendImageTag(ByVal p_Source As String, Optional ByVal p_AlternateText As String, Optional ByVal p_WidthPx As Long = -1, Optional ByVal p_HeightPx As Long = -1, Optional ByVal p_Class As String, Optional ByVal p_Id As String, Optional ByVal p_Style As String)
   With mo_StringBuilder
      .Append "<"
      .Append "img"
      
      If Not stringIsEmptyOrWhitespaceOnly(p_Source) Then
         .Append " src='"
         .Append htmlEscape(p_Source, "%20")
         .Append "' "
      End If
      
      If Not stringIsEmptyOrWhitespaceOnly(p_AlternateText) Then
         .Append " alt='"
         .Append htmlEscape(p_AlternateText)
         .Append "' "
      End If
      
      If p_WidthPx >= 0 Then
         .Append " width='"
         .Append CStr(p_WidthPx)
         .Append "' "
      End If
      
      If p_HeightPx >= 0 Then
         .Append " height='"
         .Append CStr(p_HeightPx)
         .Append "' "
      End If
            
      If Not stringIsEmptyOrWhitespaceOnly(p_Class) Then
         .Append " class='"
         .Append htmlEscape(p_Class)
         .Append "' "
      End If
      
      If Not stringIsEmptyOrWhitespaceOnly(p_Id) Then
         .Append " id='"
         .Append htmlEscape(p_Id)
         .Append "' "
      End If
      
      If Not stringIsEmptyOrWhitespaceOnly(p_Style) Then
         .Append " style='"
         .Append htmlEscape(p_Style)
         .Append "' "
      End If
      
      .Append ">"
   End With
End Sub

Public Function OpenTags(ParamArray pa_Tags() As Variant) As Long
   ' Pass as many HTML tags as you would like and they will be pushed to a stack for
   ' closing in reverse order.
   ' The passed tags can be strings, and/or arrays of strings.
   ' Do NOT include the < > characters, just the tag e.g. "p", "table", "tr", "td", etc...
   
   ' Later on you can call one of the Close* methods to close open tags.
   
   ' This function returns the index of the first added tag from the paramarray
   ' You can pass this to CloseOpenedTagsToIndex method to close every tag up to and including the passed index.
   
   ' NOTE: The Index returned by this function becomes invalid as soon as an Close* method is called!
   '       DO NOT use the returned Index for any purpose after calling any Close* method!
   
   Dim ii As Long
   Dim jj As Long
   
   If UBound(pa_Tags) = -1 Then Err.Raise 5, , "At least one tag is required."

   OpenTags = mo_Tags.Count   ' Return start index of this batch of tags so we can close back to this point if we want.
   
   For ii = LBound(pa_Tags) To UBound(pa_Tags)
      Select Case VarType(pa_Tags)
      Case vbString And vbArray
         ' Array of strings
         For jj = LBound(pa_Tags(ii)) To UBound(pa_Tags(ii))
            OpenTag pa_Tags(ii)(jj), True
         Next jj
         
      Case Else
         OpenTag pa_Tags(ii), True
         
      End Select
   Next ii
End Function

Public Function OpenHyperlinkTag(ByVal p_Href As String, Optional ByVal p_Target As String, Optional ByVal p_Class As String, Optional ByVal p_Id As String, Optional ByVal p_Style As String) As Long
   OpenHyperlinkTag = mo_Tags.Count
   
   mo_Tags.Push "a"
   
   With mo_StringBuilder
      .Append "<"
      .Append "a"
      
      If Not stringIsEmptyOrWhitespaceOnly(p_Href) Then
         .Append " href='"
         .Append htmlEscape(p_Href, "%20")
         .Append "' "
      End If
      
      If Not stringIsEmptyOrWhitespaceOnly(p_Target) Then
         .Append " target='"
         .Append htmlEscape(p_Target)
         .Append "' "
      End If
      
      If Not stringIsEmptyOrWhitespaceOnly(p_Class) Then
         .Append " class='"
         .Append htmlEscape(p_Class)
         .Append "' "
      End If
            
      If Not stringIsEmptyOrWhitespaceOnly(p_Id) Then
         .Append " id='"
         .Append htmlEscape(p_Id)
         .Append "' "
      End If
      
      If Not stringIsEmptyOrWhitespaceOnly(p_Style) Then
         .Append " style='"
         .Append htmlEscape(p_Style)
         .Append "' "
      End If
      
      .Append ">"
   End With
End Function

Public Function OpenTagWithAttributes(ByVal p_Tag As String, Optional ByVal p_TagClass As String, Optional ByVal p_TagId As String, Optional ByVal p_TagStyle As String) As Long
   OpenTagWithAttributes = mo_Tags.Count
   
   mo_Tags.Push p_Tag
   
   With mo_StringBuilder
      .Append "<"
      .Append p_Tag
      
      If Not stringIsEmptyOrWhitespaceOnly(p_TagClass) Then
         .Append " class='"
         .Append htmlEscape(p_TagClass)
         .Append "' "
      End If
            
      If Not stringIsEmptyOrWhitespaceOnly(p_TagId) Then
         .Append " id='"
         .Append htmlEscape(p_TagId)
         .Append "' "
      End If
      
      If Not stringIsEmptyOrWhitespaceOnly(p_TagStyle) Then
         .Append " style='"
         .Append htmlEscape(p_TagStyle)
         .Append "' "
      End If
      
      .Append ">"
   End With
End Function

Public Sub CloseOpenedTagsToIndex(ByVal p_Index As Long)
   Dim ii As Long
   Dim l_Tag As String
   
   ' This method will append all the closing tags remaining in the tag stack up to the passed Index in reverse order.
   
   For ii = mo_Tags.Count - 1 To p_Index Step -1
      CloseTag mo_Tags.Pop
   Next ii
End Sub

Public Sub CloseLastOpenedTag()
   Dim l_Tag As String
   
   ' This will append the last closing tags in the tag stack.
   
   If mo_Tags.Count = 0 Then Err.Raise 5, , "There are no open tags."
   
   CloseTag mo_Tags.Pop
End Sub

Public Sub CloseAllOpenedTags()
   Dim ii As Long
   Dim l_Tag As String
   
   ' This will append all the closing tags remaining in the tag stack in reverse order.
   
   For ii = mo_Tags.Count - 1 To 0 Step -1
      CloseTag mo_Tags.Pop
   Next ii
End Sub

Private Sub CloseTag(ByVal p_Tag As String)
   With mo_StringBuilder
      .Append "</"
      .Append p_Tag
      .Append ">"
   End With
End Sub

Private Sub OpenTag(ByVal p_Tag As String, ByVal p_PushToStack As Boolean)
   If p_PushToStack Then mo_Tags.Push p_Tag

   With mo_StringBuilder
      .Append "<"
      .Append p_Tag
      .Append ">"
   End With
End Sub

Public Function EncodeHtmlEntities(ByVal p_Text As String) As String
   EncodeHtmlEntities = htmlEscape(p_Text)
End Function

Public Property Let Content(ByVal p_Content As String)
   Set mo_StringBuilder = libRc5Factory.C.StringBuilder
   Set mo_Tags = libRc5Factory.C.ArrayList(vbString)
   
   Me.Append p_Content
End Property

Public Property Get Content() As String
Attribute Content.VB_UserMemId = 0
   Content = mo_StringBuilder.ToString
End Property

Public Function Length() As Long
   Length = mo_StringBuilder.Length
End Function

Public Sub Clear()
   mo_StringBuilder.Clear
End Sub

Private Sub Class_Initialize()
   Set mo_StringBuilder = libRc5Factory.C.StringBuilder
   Set mo_Tags = libRc5Factory.C.ArrayList(vbString)
End Sub

Friend Function ToUtf8() As Byte()
   ToUtf8 = mo_StringBuilder.ToUtf8
End Function

Friend Function ToString() As String
   ToString = mo_StringBuilder.ToString
End Function
