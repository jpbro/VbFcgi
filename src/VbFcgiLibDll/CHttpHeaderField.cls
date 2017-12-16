VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHttpHeaderField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' An HTTP header "item" is a single HTTP header line
' For example, "Content-Type: text/html; charset=utf-8"

Private m_Initialized As Boolean

Private m_HttpHeaderFieldName As String

Private mo_Pairs As VbFcgiLib.CHttpHeaderFieldKeyValuePairs
Private m_HttpHeaderCustomString As String

Friend Property Let HttpHeaderFieldName(ByVal p_Name As String)
   If stringIsEmptyOrWhitespaceOnly(p_Name) Then Err.Raise 5, , "HTTP Header name required."
   
   If InStr(1, p_Name, vbLf) > 0 Then Err.Raise 5, , "Newline characters are not allowed in header field names."
   If InStr(1, p_Name, vbCr) > 0 Then Err.Raise 5, , "Newline characters are not allowed in header field names."
   
   m_HttpHeaderFieldName = stringRemoveWhitespace(p_Name)
End Property

Friend Property Get HttpHeader() As String
   If Not m_Initialized Then Err.Raise 5, , "CHttpHeaderField class must be initialized first!"
   
   If Not stringIsEmptyOrWhitespaceOnly(m_HttpHeaderCustomString) Then
      ' Use custom HTTP header string passed to InitializeByString
      HttpHeader = m_HttpHeaderCustomString
   
   Else
      ' Not using custom HTTP header string
      ' so build the string from the parameters collection
      
      If mo_Pairs Is Nothing Then Err.Raise 5, , "Header parameters have not been initialized!"
      
      HttpHeader = m_HttpHeaderFieldName & ": "
      HttpHeader = HttpHeader & stringChomp(mo_Pairs.FieldValue) & vbNewLine
   End If
End Property

Public Sub InitializeByCustomString(ByVal p_CustomHttpHeaderStringWithoutName As String)
   ' Use this for a completely customized HTTP header
   If m_Initialized Then Err.Raise 5, , "Already initialized."
   
   If InStr(1, p_CustomHttpHeaderStringWithoutName, vbLf) > 0 Then Err.Raise 5, , "Newline characters are not allowed in header fields."
   If InStr(1, p_CustomHttpHeaderStringWithoutName, vbCr) > 0 Then Err.Raise 5, , "Newline characters are not allowed in header fields."
   
   Set mo_Pairs = Nothing
   
   ' Make sure the HTTP header name wasn't passed
   If LCase$(Left$(p_CustomHttpHeaderStringWithoutName, Len(m_HttpHeaderFieldName & ":")) <> LCase$(m_HttpHeaderFieldName) & ":") Then
      p_CustomHttpHeaderStringWithoutName = m_HttpHeaderFieldName & ": " & LTrim$(p_CustomHttpHeaderStringWithoutName)
   End If
   
   m_HttpHeaderCustomString = p_CustomHttpHeaderStringWithoutName
   
   m_Initialized = True
End Sub

Public Function InitializeByKeyValuePairs(Optional po_PairsCollectionOrNothing As VbFcgiLib.CHttpHeaderFieldKeyValuePairs) As VbFcgiLib.CHttpHeaderFieldKeyValuePairs
   If m_Initialized Then Err.Raise 5, , "Already initialized."
   
   m_HttpHeaderCustomString = ""
   
   If po_PairsCollectionOrNothing Is Nothing Then
      Set mo_Pairs = New VbFcgiLib.CHttpHeaderFieldKeyValuePairs
   Else
      Set mo_Pairs = po_PairsCollectionOrNothing
   End If
   
   Set InitializeByKeyValuePairs = mo_Pairs

   m_Initialized = True
End Function

