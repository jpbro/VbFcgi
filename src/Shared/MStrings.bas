Attribute VB_Name = "MStrings"
Option Explicit

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

Public Function stringIsEmptyOrWhitespaceOnly(p_Text As String) As Boolean
   Dim i As Long
   Dim l As Long

   l = Len(p_Text)
   If l <> 0 Then
      For i = 1 To l
         Select Case Asc(Mid$(p_Text, i, 1))
         Case 9, 10, 11, 12, 13, 32, 160
         Case Else
            Exit Function
         End Select
      Next i
   End If

   stringIsEmptyOrWhitespaceOnly = True
End Function

Public Function stringTrimWhitespace(ByVal pText As String) As String
   Dim i As Long
   Dim l As Long
   Dim j As Long

   l = Len(pText)

   If l > 0 Then
      For i = 1 To l
         Select Case AscW(Mid$(pText, i, 1))
         Case 9, 10, 11, 12, 13, 32, 160
         Case Else
            Exit For
         End Select
      Next i
   
      If i < l Then
         For j = l To 1 Step -1
            Select Case AscW(Mid$(pText, j, 1))
            Case 9, 10, 11, 12, 13, 32, 160
            Case Else
               Exit For
            End Select
         Next j
      End If
   
      If j >= 1 Then
         'If j > 0 Then
         If (i <> 1) Or (j <> l) Then
            stringTrimWhitespace = Mid$(pText, i, j - i + 1)
         Else
            stringTrimWhitespace = pText
         End If
         'End If
      ElseIf i > 1 Then
         stringTrimWhitespace = Mid$(pText, i)
      Else
         stringTrimWhitespace = pText
      End If
   End If
End Function

