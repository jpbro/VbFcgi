Attribute VB_Name = "MStrings"
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

Public Enum e_StringTrimWhitespace
   stringtrimwhitespace_None = 0
   stringtrimwhitespace_Left = 1
   stringtrimwhitespace_Right = 2
   stringtrimwhitespace_Both = 3
End Enum

Public Function stringMidByEndIndex(ByVal p_String As String, ByVal p_StartIndex As Long, ByVal p_EndIndex As Long) As String
   stringMidByEndIndex = Mid$(p_String, p_StartIndex, p_EndIndex - p_StartIndex)
End Function

Public Function stringRemoveWhitespace(ByVal pText As String) As String
   Dim i As Long
   Dim l_Scan As Boolean
   Dim l_Len As Long
   Dim l_Start As Long
   Dim l_Insert As Long
   Dim l_ChunkLen As Long

   l_Len = Len(pText)
   If l_Len = 0 Then Exit Function

   If Asc(Left$(pText, 1)) = 32 Then
      pText = Trim$(pText)
      l_Len = Len(pText)
      If l_Len = 0 Then Exit Function
   ElseIf Asc(Right$(pText, 1)) = 32 Then
      pText = Trim$(pText)
      l_Len = Len(pText)
      If l_Len = 0 Then Exit Function
   End If

   If InStr(1, pText, " ") Then
      pText = Replace$(pText, " ", "")
   End If

   If InStr(1, pText, vbTab) Then
      pText = Replace$(pText, vbTab, "")
   End If

   If InStr(1, pText, vbCr) Then
      pText = Replace$(pText, vbCr, "")
   End If

   If InStr(1, pText, vbFormFeed) Then
      pText = Replace$(pText, vbFormFeed, "")
   End If

   If InStr(1, pText, vbLf) Then
      pText = Replace$(pText, vbLf, "")
   End If

   If InStr(1, pText, vbVerticalTab) Then
      pText = Replace$(pText, vbVerticalTab, "")
   End If

   If InStr(1, pText, Chr$(160)) Then
      pText = Replace$(pText, Chr$(160), "")
   End If

   If InStr(1, pText, vbNullChar) Then
      pText = Replace$(pText, vbNullChar, "")
   End If

   stringRemoveWhitespace = pText
End Function

Public Function stringFlattenWhitespace(ByVal p_Text As String, Optional ByVal p_Trim As e_StringTrimWhitespace = stringtrimwhitespace_Both) As String
   ' Flatten sequential runs of 2 or more "whitespace" characters into a single space character.
   ' Optionally trimming all spaces from left and/or right ends of the string
   
   If InStr(1, p_Text, vbTab) > 0 Then
      p_Text = Replace$(p_Text, vbTab, " ")
   End If

   If InStr(1, p_Text, vbCr) > 0 Then
      p_Text = Replace$(p_Text, vbCr, " ")
   End If

   If InStr(1, p_Text, vbLf) > 0 Then
      p_Text = Replace$(p_Text, vbLf, " ")
   End If

   If InStr(1, p_Text, vbFormFeed) Then
      p_Text = Replace$(p_Text, vbFormFeed, " ")
   End If

   If InStr(1, p_Text, vbVerticalTab) Then
      p_Text = Replace$(p_Text, vbVerticalTab, " ")
   End If

   If InStr(1, p_Text, Chr$(160)) Then
      p_Text = Replace$(p_Text, Chr$(160), " ")
   End If

   If InStr(1, p_Text, vbNullChar) Then
      p_Text = Replace$(p_Text, vbNullChar, " ")
   End If

   If p_Trim = True Then p_Trim = stringtrimwhitespace_Both

   If p_Trim <> stringtrimwhitespace_None Then
      If p_Trim = stringtrimwhitespace_Both Then
         p_Text = Trim$(p_Text)
      ElseIf p_Trim = stringtrimwhitespace_Left Then
         p_Text = LTrim$(p_Text)
      ElseIf p_Trim = stringtrimwhitespace_Right Then
         p_Text = RTrim$(p_Text)
      End If
   End If

   Do While InStr(1, p_Text, "  ") > 0
      ' Flatten 2 spaces into a single space
      ' We loop to catch all longer runs of spaces
      p_Text = Replace$(p_Text, "  ", " ")
   Loop

   ' Return the result
   stringFlattenWhitespace = p_Text
End Function

Public Function stringIsEmptyOrWhitespaceOnly(p_Text As String) As Boolean
   Dim ii As Long
   Dim ll As Long

   ll = Len(p_Text)
   If ll <> 0 Then
      For ii = 1 To ll
         Select Case Asc(Mid$(p_Text, ii, 1))
         Case 9, 10, 11, 12, 13, 32, 160
         Case Else
            Exit Function
         End Select
      Next ii
   End If

   stringIsEmptyOrWhitespaceOnly = True
End Function

Public Function stringTrimWhitespace(ByVal pText As String, Optional ByVal p_TrimFrom As e_StringTrimWhitespace = stringtrimwhitespace_Left Or stringtrimwhitespace_Right) As String
   Dim ii As Long
   Dim jj As Long
   Dim ll As Long

   ll = Len(pText)

   If ll > 0 Then
      If p_TrimFrom And stringtrimwhitespace_Left Then
         For ii = 1 To ll
            Select Case AscW(Mid$(pText, ii, 1))
            Case 9, 10, 11, 12, 13, 32, 160
            Case Else
               Exit For
            End Select
         Next ii
      End If
      
      If ii < ll Then
         If p_TrimFrom And stringtrimwhitespace_Right Then
            For jj = ll To 1 Step -1
               Select Case AscW(Mid$(pText, jj, 1))
               Case 9, 10, 11, 12, 13, 32, 160
               Case Else
                  Exit For
               End Select
            Next jj
         End If
      End If
   
      If jj >= 1 Then
         If (ii <> 1) Or (jj <> ll) Then
            If ii = 0 Then ii = 1
            
            stringTrimWhitespace = Mid$(pText, ii, jj - ii + 1)
         Else
            stringTrimWhitespace = pText
         End If
      
      ElseIf ii > 1 Then
         stringTrimWhitespace = Mid$(pText, ii)
      
      Else
         stringTrimWhitespace = pText
      End If
   End If
End Function

Public Function stringChomp(ByVal p_String As String, Optional ByVal p_ChompChars As String = vbNewLine) As String
   ' Removes all p_ChompChars (if any) from the right side of the passed p_String
   Dim l_Chars As String
   Dim l_ChompLen As Long
   Dim l_StringLen As Long
   Dim l_ChopAt As Long
   Dim ii As Long
   
   l_StringLen = Len(p_String)
   l_ChompLen = Len(p_ChompChars)
   
   If l_StringLen < l_ChompLen Then Exit Function
   
   For ii = l_StringLen - l_ChompLen + 1 To 1 Step -l_ChompLen
      If Mid$(p_String, ii, l_ChompLen) = p_ChompChars Then
         l_ChopAt = ii
      Else
         Exit For
      End If
   Next ii
   
   If l_ChopAt > 0 Then
      stringChomp = Left$(p_String, l_ChopAt - 1)
   Else
      stringChomp = p_String
   End If
End Function

Public Function stringVbToMultiByteCodePage(ByVal p_String As String, ByVal p_ConvertToCodePage As Long) As Byte()
   Dim l_BufferLen As Long
   Dim la_Buffer() As Byte
   
   l_BufferLen = apiWideCharToMultiByte(p_ConvertToCodePage, 0, StrPtr(p_String), Len(p_String), 0, 0, 0, ByVal 0&)

   If l_BufferLen > 0 Then
      ReDim la_Buffer(l_BufferLen - 1)
      
      apiWideCharToMultiByte p_ConvertToCodePage, 0, StrPtr(p_String), Len(p_String), VarPtr(la_Buffer(0)), l_BufferLen, 0, ByVal 0&
   End If
   
   stringVbToMultiByteCodePage = la_Buffer
End Function

Public Function stringMultiByteCodePageToVb(pa_Bytes() As Byte, ByVal p_ConvertToCodePage As Long) As String
   Dim l_BufferLen As Long
   Dim l_Buffer As String
   
   l_BufferLen = apiMultiByteToWideChar(p_ConvertToCodePage, 0, VarPtr(pa_Bytes(0)), arraySize(pa_Bytes), 0, 0)

   If l_BufferLen > 0 Then
      l_Buffer = String$(l_BufferLen, 0)
      
      apiMultiByteToWideChar p_ConvertToCodePage, 0, VarPtr(pa_Bytes(0)), arraySize(pa_Bytes), StrPtr(l_Buffer), l_BufferLen
   End If
   
   stringMultiByteCodePageToVb = l_Buffer
End Function

Public Function stringVbToUsAscii(ByVal p_String As String) As Byte()
   stringVbToUsAscii = stringVbToMultiByteCodePage(p_String, 20127)
End Function

Public Function stringVbToIso88591(ByVal p_String As String) As Byte()
   stringVbToIso88591 = stringVbToMultiByteCodePage(p_String, 28591)
End Function

Public Function stringUsAsciiToVb(pa_Bytes() As Byte) As String
   stringUsAsciiToVb = stringMultiByteCodePageToVb(pa_Bytes, 20127)
End Function

Public Function stringIso88591ToVb(pa_Bytes() As Byte) As String
   stringIso88591ToVb = stringMultiByteCodePageToVb(pa_Bytes, 28591)
End Function

Public Function stringUtf8ToVb(pa_Bytes() As Byte) As String
   stringUtf8ToVb = stringMultiByteCodePageToVb(pa_Bytes, 65001)
End Function

Public Function stringVbToUtf8(p_String As String) As Byte()
   stringVbToUtf8 = stringVbToMultiByteCodePage(p_String, 65001)
End Function

