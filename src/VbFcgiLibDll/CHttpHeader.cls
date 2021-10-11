VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHttpHeader"
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

Public Event BeforeBuildHttpHeader()   ' Give host a chance to add missing headers

Private mo_Headers As vbRichClient5.cCollection

Public Property Get ItemByName(ByVal p_HeaderName As String) As VbFcgiLib.CHttpHeaderField
   Set ItemByName = mo_Headers.Item(p_HeaderName)
End Property

Public Property Get ItemByIndex(ByVal p_ZeroBasedIndex As Long) As VbFcgiLib.CHttpHeaderField
   Set ItemByIndex = mo_Headers.ItemByIndex(p_ZeroBasedIndex)
End Property

Public Function Exists(ByVal p_HeaderName As String) As Boolean
   Exists = mo_Headers.Exists(p_HeaderName)
End Function

Public Sub Remove(ByVal p_HeaderName As String)
   mo_Headers.Remove p_HeaderName
End Sub

Public Function AddOrReplace(ByVal p_HeaderName As String, Optional ByVal p_PrimaryValue As String = "") As VbFcgiLib.CHttpHeaderField
   If Me.Exists(p_HeaderName) Then Me.Remove p_HeaderName
   
   Set AddOrReplace = Me.Add(p_HeaderName, p_PrimaryValue)
End Function

Public Function Add(ByVal p_HeaderName As String, Optional ByVal p_PrimaryValue As String) As VbFcgiLib.CHttpHeaderField
   ' All HTTP headers fields accept a primary value (for example: Content-Type: text/html")
   ' Pass the primary value via the p_PrimaryValue parameter.
   ' Some HTTP header fields accept multiple subsequent ";" separated values
   ' Those should be added via the InitializeByKeyValuePairs method
   
   If InStr(1, p_PrimaryValue, ";") > 0 Then Err.Raise 5, , "Only pass the primary value to this method and use InitializeByKeyValuePairs to add additional key/value pairs to the field."
   
   Set Add = New VbFcgiLib.CHttpHeaderField
   Add.HttpHeaderFieldName = p_HeaderName
   
   mo_Headers.Add Add, p_HeaderName

   If Not stringIsEmptyOrWhitespaceOnly(p_PrimaryValue) Then
      With Add.InitializeByKeyValuePairs()
         .Add "", p_PrimaryValue
      End With
   End If
End Function

Public Sub Append(ByVal p_HttpHeaderFields As String)
   Dim ii As Long
   Dim la_Lines() As String
   Dim la_Header() As String
   
   la_Lines = Split(p_HttpHeaderFields, vbNewLine)
   
   For ii = LBound(la_Lines) To UBound(la_Lines)
      If stringIsEmptyOrWhitespaceOnly(la_Lines(ii)) Then
         Debug.Assert False
      
      Else
         
         la_Header = Split(la_Lines(ii), ":", 2)
         
         If arraySize(la_Header) = 2 Then
            With Me.Add(stringTrimWhitespace(la_Header(0)))
               .InitializeByCustomString stringTrimWhitespace(la_Header(1), stringtrimwhitespace_Left)
            End With
         
         Else
            Err.Raise 5, , "Bad HTTP header line: " & la_Lines(ii)
         End If
      
      End If
   Next ii
End Sub

Public Function Count() As Long
   Count = mo_Headers.Count
End Function

Public Function Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4
Attribute Enumerator.VB_MemberFlags = "40"
   Set Enumerator = mo_Headers.Enumerator
End Function

Public Function Content(Optional ByVal p_Encoding As VbFcgiLib.e_ContentEncoding = contentencoding_RecommendedDefault, Optional ByVal p_IncludeTrailingDoubleNewlines As Boolean = True) As Byte()
   Dim lo_HttpHeader As vbRichClient5.cStringBuilder
   Dim lo_HttpHeaderField As VbFcgiLib.CHttpHeaderField
   Dim l_HttpHeaderField As String

   If p_Encoding = contentencoding_RecommendedDefault Then p_Encoding = contentencoding_USASCII

   ' Raise event to allow caller to add/replace headers
   RaiseEvent BeforeBuildHttpHeader
   
   ' Begin building HTTP headers
   Set lo_HttpHeader = libRc5Factory.C.StringBuilder
   
   For Each lo_HttpHeaderField In mo_Headers
      l_HttpHeaderField = stringChomp(lo_HttpHeaderField.HttpHeader)
      
      lo_HttpHeader.AppendNL l_HttpHeaderField
   Next lo_HttpHeaderField
   
   If p_IncludeTrailingDoubleNewlines Then
      lo_HttpHeader.Append vbNewLine
   End If
   
   Select Case p_Encoding
   Case contentencoding_UTF16_LE
      ' Return as VB6 string
      Content = lo_HttpHeader.ToString
   
   Case contentencoding_UTF8
      ' Return in UTF-8 encoding
      Content = lo_HttpHeader.ToUtf8
      
   Case contentencoding_ISO8859_1
      ' Return in ISO-8859-1 encoding
      Content = stringVbToIso88591(lo_HttpHeader.ToString)
   
   Case contentencoding_USASCII
      ' Return in US-ASCII encoding
      Content = stringVbToUsAscii(lo_HttpHeader.ToString)
   
   Case Else
      ' Unhandled encoding type
      Err.Raise vbObjectError, , "Unknown content encoding: " & p_Encoding
   
   End Select
   
End Function

Private Sub Class_Initialize()
   Set mo_Headers = libRc5Factory.C.Collection(False, TextCompare, False)
End Sub
