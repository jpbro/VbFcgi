VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBuilderHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum e_HtmlDocType
   htmldoctype_Html5
End Enum

Private mo_StringBuilder As vbRichClient5.cStringBuilder
Private mo_Tags As vbRichClient5.cArrayList
Private m_Finished As Boolean
Private m_Length As Long
Private m_ContentEncoding As VbFcgiLib.e_ContentEncoding
Private ma_Content() As Byte

Private Sub Class_Initialize()
   Me.Reset
End Sub

Public Sub Reset()
   m_Finished = False
   m_ContentEncoding = contentencoding_Unknown
   
   Erase ma_Content
   
   Set mo_StringBuilder = libRc5Factory.C.StringBuilder
   Set mo_Tags = libRc5Factory.C.ArrayList(vbString)
End Sub

Public Sub Finished(Optional ByVal p_ContentEncoding As VbFcgiLib.e_ContentEncoding = contentencoding_UTF8)
   ' Call this method when you are finished building your HTML
   ' Once called, the Content, Length, and ContentEncoding properties will become valid.
      
   Select Case p_ContentEncoding
   Case contentencoding_UTF8
      Me.CloseAllOpenedTags
      
      ma_Content = mo_StringBuilder.ToUTF8
   
   Case contentencoding_UTF16_LE
      Me.CloseAllOpenedTags
      
      ma_Content = mo_StringBuilder.ToString
   
   Case Else
      Err.Raise 5, , "Unknown content encoding enum: " & p_ContentEncoding
   End Select
   
   m_Finished = True
   m_ContentEncoding = p_ContentEncoding
   m_Length = UBound(ma_Content) - LBound(ma_Content) + 1
   
   Set mo_StringBuilder = Nothing
   Set mo_Tags = Nothing
End Sub

Public Property Get Content() As Byte()
   ' You must call the Finish method before you call this property.
   ' Call this property to get your finished content as a Me.ContentEncoding format Byte Array
   
   If Not m_Finished Then Err.Raise 5, , "You must call the Finished method before requesting content."
   
   Content = ma_Content
End Property

Public Property Get Length() As Long
   ' You must call the Finish method before you call this property.
   
   ' This property returns the length of the Content byte array.
   
   If Not m_Finished Then Err.Raise 5, , "You must call the Finished method before requesting the content length."
   
   Length = m_Length
End Property

Public Property Get ContentEncoding() As VbFcgiLib.e_ContentEncoding
   ' You must call the Finish method before you call this property.
   
   ' This property returns the content encoding enum of the Content byte array
   ' e.g. UTF8 or UTF16LE.
   
   If Not m_Finished Then Err.Raise 5, , "You must call the Finished method before requesting content encoding."
   
   ContentEncoding = m_ContentEncoding
End Property

Public Sub AppendDocType(Optional ByVal p_DocType As VbFcgiLib.e_HtmlDocType)
   ' Append a standard DOCTYPE to the top of an HTML page
   ' Currently only HTML5 is supported.
   
   If m_Finished Then Err.Raise 5, , "The content buffer is closed."
   
   If mo_StringBuilder.Length > 0 Then Err.Raise 5, , "DOCTYPE must be the first thing you append to your HTML page."
   
   Select Case p_DocType
   Case htmldoctype_Html5
      mo_StringBuilder.AppendNL "<!DOCTYPE html>"
   Case Else
      Err.Raise 5, , "Unknown DocType enum: " & p_DocType
   End Select
End Sub

Public Sub Append(ByVal p_Text As String)
   ' Append any string to the content
   
   If m_Finished Then Err.Raise 5, , "The content buffer is closed."
   
   mo_StringBuilder.Append p_Text
End Sub

Public Sub Append2(ByVal p_Text1 As String, ByVal p_Text2 As String)
   ' Append any 2 strings (one after the other) to the content
   
   If m_Finished Then Err.Raise 5, , "The content buffer is closed."
   
   With mo_StringBuilder
      .Append p_Text1
      .Append p_Text2
   End With
End Sub

Public Sub Append3(ByVal p_Text1 As String, ByVal p_Text2 As String, ByVal p_Text3 As String)
   ' Append any 3 strings (one after the other) to the content
   
   If m_Finished Then Err.Raise 5, , "The content buffer is closed."
   
   With mo_StringBuilder
      .Append p_Text1
      .Append p_Text2
      .Append p_Text3
   End With
End Sub

Public Sub Append4(ByVal p_Text1 As String, ByVal p_Text2 As String, ByVal p_Text3 As String, ByVal p_Text4 As String)
   ' Append any 4 strings (one after the other) to the content
   
   If m_Finished Then Err.Raise 5, , "The content buffer is closed."
   
   With mo_StringBuilder
      .Append p_Text1
      .Append p_Text2
      .Append p_Text3
      .Append p_Text4
   End With
End Sub

Public Sub AppendWithTag(ByVal p_Text As String, ByVal p_Tag As String, Optional ByVal p_EncodeEntities As Boolean = True)
   ' In one pass this method will:
   
   ' Append an HTML tag (don't pass the "<>" characters),
   ' Then Append any arbitrary text (optionally encoding HTML entities (default))
   ' Then append the closing tag.
   
   If m_Finished Then Err.Raise 5, , "The content buffer is closed."
   
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

Public Sub AppendWithTagAndAttributes(ByVal p_Text As String, ByVal p_Tag As String, Optional ByVal p_TagName As String, Optional ByVal p_TagId As String, Optional ByVal p_TagStyle As String, Optional ByVal p_EncodeEntities As Boolean = True)
   ' In one pass this method will:
   
   ' Append an HTML tag (don't pass the "<>" characters),
   ' Optionally appending a name='passedname' attribute to the tag
   ' Optionally appending an id='passedid' attribute to the tag
   ' Optionally appending a style='passedstyles' attribute to the tag
   ' Then Append any arbitrary text (optionally encoding HTML entities (default))
   ' Then append the closing tag.
   
   If m_Finished Then Err.Raise 5, , "The content buffer is closed."
   
   With mo_StringBuilder
      .Append "<"
      .Append p_Tag
      
      If Not stringIsEmptyOrWhitespaceOnly(p_TagName) Then
         .Append " name='"
         .Append htmlEscape(p_TagName)
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
      
      If p_EncodeEntities Then
         .Append htmlEscape(p_Text)
      Else
         .Append p_Text
      End If
      
      CloseTag p_Tag
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
   
   If m_Finished Then Err.Raise 5, , "The content buffer is closed."
   
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

Public Sub CloseOpenedTagsToIndex(ByVal p_Index As Long)
   Dim ii As Long
   Dim l_Tag As String
   
   ' This method will append all the closing tags remaining in the tag stack up to the passed Index in reverse order.
   
   If m_Finished Then Err.Raise 5, , "The content buffer is closed."
   
   For ii = mo_Tags.Count - 1 To p_Index Step -1
      CloseTag mo_Tags.Pop
   Next ii
End Sub

Public Sub CloseLastOpenedTag()
   Dim l_Tag As String
   
   ' This will append the last closing tags in the tag stack.
   
   If m_Finished Then Err.Raise 5, , "The content buffer is closed."
   
   If mo_Tags.Count = 0 Then Err.Raise 5, , "There are no open tags."
   
   CloseTag mo_Tags.Pop
End Sub

Public Sub CloseAllOpenedTags()
   Dim ii As Long
   Dim l_Tag As String
   
   ' This will append all the closing tags remaining in the tag stack in reverse order.
   
   If m_Finished Then Err.Raise 5, , "The content buffer is closed."
   
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
   If p_PushToStack Then mo_Tags.Push CStr(p_Tag)

   With mo_StringBuilder
      .Append "<"
      .Append p_Tag
      .Append ">"
   End With
End Sub

