VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHttpHeaderFieldKeyValuePairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mo_Pairs As vbRichClient5.cCollection
Private m_Dirty As Boolean

Public Function Exists(ByVal p_Key As String) As Boolean
   If mo_Pairs.Count = 0 Then Exit Function
   
   If stringIsEmptyOrWhitespaceOnly(p_Key) Then
      Exists = (mo_Pairs.KeyByIndex(0) = "")
   Else
      Exists = mo_Pairs.Exists(p_Key)
   End If
End Function

Public Sub Add(ByVal p_Key As String, ByVal p_Value As String)
   If stringIsEmptyOrWhitespaceOnly(p_Key) Then
      ' Empty key must be first item in collection
      If mo_Pairs.Count Then Err.Raise 5, , "Empty key must be first key/pair."
   End If

   If InStr(1, p_Key, vbCr) > 0 Then Err.Raise 5, , "Key can not contain new line character."
   If InStr(1, p_Key, vbLf) > 0 Then Err.Raise 5, , "Key can not contain new line character."
   If InStr(1, p_Value, vbCr) > 0 Then Err.Raise 5, , "Value can not contain new line character."
   If InStr(1, p_Value, vbLf) > 0 Then Err.Raise 5, , "Value can not contain new line character."

   m_Dirty = True
   
   If stringIsEmptyOrWhitespaceOnly(p_Key) Then
      mo_Pairs.Add p_Value
   Else
      mo_Pairs.Add p_Value, p_Key
   End If
End Sub

Friend Function FieldValue() As String
   Static s_FieldValue As String
   
   Dim ii As Long
   Dim jj As Long
   Dim l_Key As String
   Dim l_Value As String
   Dim l_QuoteValue As Boolean
   Dim l_Char As String
   
   If stringIsEmptyOrWhitespaceOnly(s_FieldValue) Or m_Dirty Then
      s_FieldValue = ""
      
      For ii = 0 To mo_Pairs.Count - 1
         l_Key = mo_Pairs.KeyByIndex(ii)
         
         If l_Key = "" Then
            If ii > 0 Then
               Err.Raise vbObjectError, , "Empty key must be the first parameter."
            End If
         Else
            l_Key = l_Key & "="
         End If
         
         If ii > 0 Then
            l_Key = "; " & l_Key
         End If
         
         l_Value = mo_Pairs.ItemByIndex(ii)
         
         If Left$(l_Value, 1) <> """" Then
            If l_Key <> "" Then  ' First parameters don't seem to require quoting?
                                 ' For example Content-Type: text/html would appear to need quoting
                                 ' as per RFC2616 since it contains a "/" character (which is on the special characters list)
                                 ' but it is never done in the wild. Dates contain ":" characters, but they too are never quoted
                                 ' when they are the first token of a header field.
                                 
               ' See if we need to quote this value
               For jj = 1 To Len(l_Value)
                  
                  l_Char = Mid$(l_Value, jj, 1)
                  
                  Select Case l_Char
                  Case "(", ")", "<", ">", "@", _
                       ",", ";", ":", "\", "<", _
                       ">", "/", "[", "]", "?", "=", _
                       "{", "}", vbTab, " "
                     
                     ' RFC2616 says these characters must be placed in quoted strings
                     l_QuoteValue = True
                  
                  Case Else
                     Select Case AscW(l_Char)
                     Case Is < 0, &H0 To &H8, &HA To &H1A, &H22, &H7F To &H7FFF
                        ' Illegal character
                        Err.Raise vbObjectError, , "Illegal character '" & Mid$(l_Value, jj, 1) & "' found in parameter for key: " & l_Key
                     End Select
                  End Select
               Next jj
            End If
         End If
         
         If l_QuoteValue Then
            l_QuoteValue = False
            l_Value = """" & Replace$(l_Value, """", "\""") & """"
         End If
         
         s_FieldValue = s_FieldValue & l_Key & l_Value
         
      Next ii
   End If
   
   FieldValue = s_FieldValue
End Function

Private Sub Class_Initialize()
   Set mo_Pairs = libRc5Factory.C.Collection(False)
End Sub
