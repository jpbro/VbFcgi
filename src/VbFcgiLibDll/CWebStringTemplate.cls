VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWebStringTemplate"
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

' Wraps the CWebStringBase class and adds extra properties specific to CBuilderTemplate

Private mo_String As VbFcgiLib.CWebStringBase

Public SkipCache As Boolean   ' Do not cache this content for use in other matching tags without raising an event.
Public SkipEncodeEntities As Boolean   ' Do not encode HTML entities.

Private Sub Class_Initialize()
   Set mo_String = New VbFcgiLib.CWebStringBase
End Sub

' The following methods wrap the corresponding methods in the CWebString class.

Public Sub Append(ByVal p_Text As String)
   mo_String.Append p_Text
End Sub

Public Sub Append2(ByVal p_Text1 As String, ByVal p_Text2 As String)
   mo_String.Append2 p_Text1, p_Text2
End Sub

Public Sub Append3(ByVal p_Text1 As String, ByVal p_Text2 As String, ByVal p_Text3 As String)
   mo_String.Append3 p_Text1, p_Text2, p_Text3
End Sub

Public Sub Append4(ByVal p_Text1 As String, ByVal p_Text2 As String, ByVal p_Text3 As String, ByVal p_Text4 As String)
   mo_String.Append4 p_Text1, p_Text2, p_Text3, p_Text4
End Sub

Public Sub AppendDocType(Optional ByVal p_DocType As e_HtmlDocType = 0&, Optional ByVal p_CustomDocType As String)
   mo_String.AppendDocType p_DocType, p_CustomDocType
End Sub

Public Sub AppendImageTag(ByVal p_Source As String, Optional ByVal p_AlternateText As String, Optional ByVal p_WidthPx As Long = -1&, Optional ByVal p_HeightPx As Long = -1&, Optional ByVal p_Class As String, Optional ByVal p_Id As String, Optional ByVal p_Style As String)
   mo_String.AppendImageTag p_Source, p_AlternateText, p_WidthPx, p_HeightPx, p_Class, p_Id, p_Style
End Sub

Public Sub AppendWithTag(ByVal p_Text As String, ByVal p_Tag As String, Optional ByVal p_EncodeEntities As Boolean = True)
   mo_String.AppendWithTag p_Text, p_Tag, p_EncodeEntities
End Sub

Public Sub AppendWithTagAndAttributes(ByVal p_Text As String, ByVal p_Tag As String, Optional ByVal p_TagClass As String, Optional ByVal p_TagId As String, Optional ByVal p_TagStyle As String, Optional ByVal p_EncodeEntities As Boolean = True)
   mo_String.AppendWithTagAndAttributes p_Text, p_Tag, p_TagClass, p_TagId, p_TagStyle, p_EncodeEntities
End Sub

Public Sub Clear()
   mo_String.Clear
End Sub

Public Sub CloseAllOpenedTags()
   mo_String.CloseAllOpenedTags
End Sub

Public Sub CloseLastOpenedTag()
   mo_String.CloseLastOpenedTag
End Sub

Public Sub CloseOpenedTagsToIndex(ByVal p_Index As Long)
   mo_String.CloseOpenedTagsToIndex p_Index
End Sub

Public Property Let Content(ByVal p_Content As String)
   mo_String.Content = p_Content
End Property

Public Property Get Content() As String
Attribute Content.VB_UserMemId = 0
   Content = mo_String.Content
End Property

Public Function EncodeHtmlEntities(ByVal p_Text As String) As String
   EncodeHtmlEntities = mo_String.EncodeHtmlEntities(p_Text)
End Function

Public Function Length() As Long
   Length = mo_String.Length
End Function

Public Function OpenHyperlinkTag(ByVal p_Href As String, Optional ByVal p_Target As String, Optional ByVal p_Class As String, Optional ByVal p_Id As String, Optional ByVal p_Style As String) As Long
   OpenHyperlinkTag = mo_String.OpenHyperlinkTag(p_Href, p_Target, p_Class, p_Id, p_Style)
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
   
   Select Case UBound(pa_Tags)
   Case -1
      ' No parameters
      OpenTags = mo_String.OpenTags
   
   Case 0
      OpenTags = mo_String.OpenTags(pa_Tags(0))
   
   Case 1
      OpenTags = mo_String.OpenTags(pa_Tags(0), pa_Tags(1))
   
   Case 2
      OpenTags = mo_String.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2))
   
   Case 3
      OpenTags = mo_String.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3))
   
   Case 4
      OpenTags = mo_String.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4))
   
   Case 5
      OpenTags = mo_String.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4), pa_Tags(5))
   
   Case 6
      OpenTags = mo_String.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4), pa_Tags(5), pa_Tags(6))
   
   Case 7
      OpenTags = mo_String.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4), pa_Tags(5), pa_Tags(6), pa_Tags(7))
   
   Case 8
      OpenTags = mo_String.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4), pa_Tags(5), pa_Tags(6), pa_Tags(7), pa_Tags(8))
   
   Case 9
      OpenTags = mo_String.OpenTags(pa_Tags(0), pa_Tags(1), pa_Tags(2), pa_Tags(3), pa_Tags(4), pa_Tags(5), pa_Tags(6), pa_Tags(7), pa_Tags(8), pa_Tags(9))
   
   Case Else
      ' Max tags reached - we can always extend this limit as required by adding more Cases above.
      Err.Raise 5, , "Maximum Tag count reached."
   End Select
End Function

Public Function OpenTagWithAttributes(ByVal p_Tag As String, Optional ByVal p_TagClass As String, Optional ByVal p_TagId As String, Optional ByVal p_TagStyle As String) As Long
   OpenTagWithAttributes = mo_String.OpenTagWithAttributes(p_Tag, p_TagClass, p_TagId, p_TagStyle)
End Function
