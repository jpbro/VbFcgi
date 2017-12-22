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

Implements VbFcgiLib.IBuilder

Public Event Finished()

Public Enum e_HtmlDocType
   htmldoctype_Html5
   
   htmldoctype_Custom = &H7FFFFFFF
End Enum

Private WithEvents mo_HttpHeader As VbFcgiLib.CHttpHeader
Attribute mo_HttpHeader.VB_VarHelpID = -1

Private mo_WebString As VbFcgiLib.CWebStringBase

Private m_Finished As Boolean
Private m_Length As Long
Private m_ContentEncoding As VbFcgiLib.e_ContentEncoding
Private ma_Content() As Byte

Public Property Get IBuilderInterface() As VbFcgiLib.IBuilder
   Set IBuilderInterface = Me
End Property

Private Sub Class_Initialize()
   IBuilder_Reset
End Sub

Public Property Get ContentEncoding() As VbFcgiLib.e_ContentEncoding
   ' You must call the Finish method before you call this property.
   
   ' This property returns the content encoding enum of the Content byte array
   ' e.g. UTF8 or UTF16LE.
   
   If Not m_Finished Then Err.Raise 5, , "You must call the Finished method before requesting content encoding."
   
   ContentEncoding = m_ContentEncoding
End Property

Public Sub AppendDocType(Optional ByVal p_DocType As VbFcgiLib.e_HtmlDocType = htmldoctype_Html5, Optional ByVal p_CustomDocType As String)
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   mo_WebString.AppendDocType p_DocType, p_CustomDocType
End Sub

Public Sub Append(ByVal p_Text As String)
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   mo_WebString.Append p_Text
End Sub

Public Sub Append2(ByVal p_Text1 As String, ByVal p_Text2 As String)
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   mo_WebString.Append2 p_Text1, p_Text2
End Sub

Public Sub Append3(ByVal p_Text1 As String, ByVal p_Text2 As String, ByVal p_Text3 As String)
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   mo_WebString.Append3 p_Text1, p_Text2, p_Text3
End Sub

Public Sub Append4(ByVal p_Text1 As String, ByVal p_Text2 As String, ByVal p_Text3 As String, ByVal p_Text4 As String)
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   mo_WebString.Append4 p_Text1, p_Text2, p_Text3, p_Text4
End Sub

Public Sub AppendWithTag(ByVal p_Text As String, ByVal p_Tag As String, Optional ByVal p_EncodeEntities As Boolean = True)
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   mo_WebString.AppendWithTag p_Text, p_Tag, p_EncodeEntities
End Sub

Public Sub AppendWithTagAndAttributes(ByVal p_Text As String, ByVal p_Tag As String, Optional ByVal p_TagClass As String, Optional ByVal p_TagId As String, Optional ByVal p_TagStyle As String, Optional ByVal p_EncodeEntities As Boolean = True)
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   mo_WebString.AppendWithTagAndAttributes p_Text, p_Tag, p_TagClass, p_TagId, p_TagStyle, p_EncodeEntities
End Sub

Public Sub AppendImageTag(ByVal p_Source As String, Optional ByVal p_AlternateText As String, Optional ByVal p_WidthPx As Long = -1, Optional ByVal p_HeightPx As Long = -1, Optional ByVal p_Class As String, Optional ByVal p_Id As String, Optional ByVal p_Style As String)
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   mo_WebString.AppendImageTag p_Source, p_AlternateText, p_WidthPx, p_HeightPx, p_Class, p_Id, p_Style
End Sub

Public Sub CloseAllOpenedTags()
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   mo_WebString.CloseAllOpenedTags
End Sub

Public Sub CloseLastOpenedTag()
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   mo_WebString.CloseLastOpenedTag
End Sub

Public Sub CloseOpenedTagsToIndex(ByVal p_Index As Long)
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   mo_WebString.CloseOpenedTagsToIndex p_Index
End Sub

Public Function EncodeHtmlEntities(ByVal p_Text As String) As String
   EncodeHtmlEntities = mo_WebString.EncodeHtmlEntities(p_Text)
End Function

Public Function OpenHyperlinkTag(ByVal p_Href As String, Optional ByVal p_Target As String, Optional ByVal p_Class As String, Optional ByVal p_Id As String, Optional ByVal p_Style As String) As Long
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   OpenHyperlinkTag = mo_WebString.OpenHyperlinkTag(p_Href, p_Target, p_Class, p_Id, p_Style)
End Function

Public Function OpenTags(ParamArray pa_Tags() As Variant) As Long
   ' This is an ugly solution for passing ParamArray parameter values up to another method that takes a ParamArry parameter.
   ' Since we can't just pass pa_Tags as a single parameter, we have to check the Ubound of pa_tags and pass each element individually.
   ' See converstaion here: http://www.vbforums.com/showthread.php?857003-Default-Class-Property-As-Obejct
   
   ' I don't like it for many reasons, but the worst might be that we are
   ' stuck at a range of X to Y parameters unless we modify the source code.
   ' I've set the range at 0 to 9 as of December 21, 2017 (which should be sufficient, but let's all laugh at the comment at some unspecified future time).
   
   ' A different approach is demonstrated here: http://www.vbforums.com/showthread.php?844667-VB6-VB6-tlb-Passing-a-ParamArray-without-Copying
   ' But I'm unsure of the licensing status of that code, plus the maintenance burden looks quite scary if there should be any bugs.
   
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   Select Case UBound(pa_Tags)
   Case -1
      ' No parameters
      OpenTags = mo_WebString.OpenTags
   
   Case 0
      OpenTags = mo_WebString.OpenTags(pa_Tags(0))
   
   Case 1
      OpenTags = mo_WebString.OpenTags(pa_Tags(0), pa_Tags(1))
   
   Case 2
      OpenTags = mo_WebString.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2))
   
   Case 3
      OpenTags = mo_WebString.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3))
   
   Case 4
      OpenTags = mo_WebString.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4))
   
   Case 5
      OpenTags = mo_WebString.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4), pa_Tags(5))
   
   Case 6
      OpenTags = mo_WebString.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4), pa_Tags(5), pa_Tags(6))
   
   Case 7
      OpenTags = mo_WebString.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4), pa_Tags(5), pa_Tags(6), pa_Tags(7))
   
   Case 8
      OpenTags = mo_WebString.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4), pa_Tags(5), pa_Tags(6), pa_Tags(7), pa_Tags(8))
   
   Case 9
      OpenTags = mo_WebString.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4), pa_Tags(5), pa_Tags(6), pa_Tags(7), pa_Tags(8), pa_Tags(9))
   
   Case Else
      ' Max tags reached - we can always extend this limit as required by adding more Cases above.
      Err.Raise 5, , "Maximum Tag count reached."
   End Select
End Function

Public Function OpenTagWithAttributes(ByVal p_Tag As String, Optional ByVal p_TagClass As String, Optional ByVal p_TagId As String, Optional ByVal p_TagStyle As String) As Long
   If m_Finished Then Err.Raise 5, , "You can not call this method after calling the Finish method."
   
   OpenTagWithAttributes = mo_WebString.OpenTagWithAttributes(p_Tag, p_TagClass, p_TagId, p_TagStyle)
End Function

Private Property Get IBuilder_Content() As Byte()
   ' You must call the Finish method before you call this property.
   ' Call this property to get your finished content as a Me.ContentEncoding format Byte Array
   
   If Not m_Finished Then Err.Raise 5, , "You must call the Finished method before requesting content."
   
   IBuilder_Content = ma_Content
End Property

Private Property Get IBuilder_HttpHeader() As VbFcgiLib.CHttpHeader
   Set IBuilder_HttpHeader = mo_HttpHeader
End Property

Private Sub IBuilder_Finish(Optional ByVal p_ContentEncoding As VbFcgiLib.e_ContentEncoding = contentencoding_RecommendedDefault)
   ' Call this method when you are finished building your HTML
   ' Once called, the Content, Length, and ContentEncoding properties will become valid.
      
   If m_Finished Then Err.Raise 5, , "Already finished."
      
   If p_ContentEncoding = contentencoding_RecommendedDefault Then p_ContentEncoding = contentencoding_UTF8
      
   Select Case p_ContentEncoding
   Case contentencoding_UTF8
      Me.CloseAllOpenedTags
      
      ma_Content = mo_WebString.ToUtf8
   
   Case contentencoding_UTF16_LE
      Me.CloseAllOpenedTags
      
      ma_Content = mo_WebString.ToString
   
   Case contentencoding_ISO8859_1
      Me.CloseAllOpenedTags
      
      ma_Content = stringVbToIso88591(mo_WebString.ToString)
      
   Case contentencoding_USASCII
      Me.CloseAllOpenedTags
      
      ma_Content = stringVbToUsAscii(mo_WebString.ToString)
      
   Case Else
      Err.Raise 5, , "Unknown content encoding enum: " & p_ContentEncoding
   End Select
   
   m_Finished = True
   m_ContentEncoding = p_ContentEncoding
   m_Length = arraySize(ma_Content)
   
   Set mo_WebString = New VbFcgiLib.CWebStringBase

   RaiseEvent Finished
End Sub

Private Property Get IBuilder_IsFinished() As Boolean
   IBuilder_IsFinished = m_Finished
End Property

Private Property Get IBuilder_Length() As Long
   ' You must call the Finish method before you call this property.
   
   ' This property returns the length of the Content byte array.
   
   If Not m_Finished Then Err.Raise 5, , "You must call the Finished method before requesting the content length."
   
   IBuilder_Length = m_Length
End Property

Private Sub IBuilder_Reset()
   m_Finished = False
   m_ContentEncoding = contentencoding_Unknown
   
   Erase ma_Content
   
   Set mo_HttpHeader = New VbFcgiLib.CHttpHeader
   Set mo_WebString = New VbFcgiLib.CWebStringBase
End Sub

Public Sub Finish(ByVal p_ContentEncoding As VbFcgiLib.e_ContentEncoding)
   IBuilder_Finish p_ContentEncoding
End Sub

Public Function Content() As Byte()
   Content = IBuilder_Content
End Function

Public Function Length() As Long
   Length = IBuilder_Length
End Function

Public Function IsFinished() As Boolean
   IsFinished = IBuilder_IsFinished
End Function

Private Sub mo_HttpHeader_BeforeBuildHttpHeader()
   ' Make sure the HTTP header includes the "Content-Type: text/html"
   ' and "Content-Length" headers.
   
   Dim l_Charset As String
   
   ' Add Content-Type header if it doesn't already exist
   If Not mo_HttpHeader.Exists("Content-Type") Then
      With mo_HttpHeader.Add("Content-Type", "text/html")
         With .KeyValuePairs()
            Select Case Me.ContentEncoding
            Case contentencoding_UTF8
               l_Charset = "utf-8"
               
            Case contentencoding_UTF16_LE
               l_Charset = "utf-16le"
               
            Case contentencoding_USASCII
               l_Charset = "us-ascii"
            
            Case contentencoding_ISO8859_1
               l_Charset = "iso-8859-1"
               
            Case Else
               Debug.Assert False  ' Unhandled content encoding
               
            End Select
            
            If Not stringIsEmptyOrWhitespaceOnly(l_Charset) Then
               .Add "charset", l_Charset
            End If
            
         End With
      End With
   End If
   
   ' Add or Replace Content-Length header if HTML build is finished
   If Me.IsFinished Then
      mo_HttpHeader.AddOrReplace "Content-Length", Me.Length
   End If
End Sub
