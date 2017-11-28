VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHttpQueryParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mo_QueryParams As vbRichClient5.cCollection

Public Property Get Exists(ByVal p_Key As String) As Boolean
   Exists = mo_QueryParams.Exists(p_Key)
End Property

Public Sub RemoveByKey(ByVal p_Key As String)
   If Me.Exists(p_Key) Then
      mo_QueryParams.Remove p_Key
   End If
End Sub

Public Sub AppendValue(ByVal p_Key As String, ByVal p_Value As String)
   Dim lo_Values As CHttpQueryParamValues
   Dim l_Exists As Boolean
   
   l_Exists = mo_QueryParams.Exists(p_Key)
   If l_Exists Then
      Set lo_Values = mo_QueryParams.Item(p_Key)
   Else
      Set lo_Values = New CHttpQueryParamValues
   End If
   
   lo_Values.Add p_Value

   If Not l_Exists Then
      mo_QueryParams.Add lo_Values, p_Key
   End If
End Sub

Public Property Get DefaultValue(p_KeyOrZeroBasedIndex As Variant) As String
Attribute DefaultValue.VB_UserMemId = 0
Attribute DefaultValue.VB_MemberFlags = "40"
   ' Shortcut for ValuesByKey("key").ValueByIndex(0)
   '  or ValuesByIndex(#).ValueByIndex(0)
   
   If VarType(p_Key) = vbString Then
      DefaultValue = Me.ValuesByKey(p_Key).ValueByIndex(0)
   Else
      DefaultValue = Me.ValuesByIndex(p_Key).ValueByIndex(0)
   End If
End Property

Public Property Get ValuesByKey(ByVal p_Key As String) As CHttpQueryParamValues
   Dim ii As Long
   
   If Trim$(p_Key) = "" Then
      ' Get default, requires loop search since cCollection raises error on empty key
      For ii = 0 To mo_QueryParams.Count - 1
         If IsEmpty(mo_QueryParams.ItemByIndex(ii)) Then
            Set ValuesByKey = mo_QueryParams.ItemByIndex(ii)
            
            Exit For
         End If
      Next ii
      
   Else
      Set ValuesByKey = mo_QueryParams.Item(p_Key)
      
   End If
   
   If ValuesByKey Is Nothing Then
      Err.Raise vbObjectError, , "Values could not be found for key: " & p_Key
   End If
End Property

Public Property Get KeyByIndex(ByVal p_ZeroBasedIndex As Long) As String
   KeyByIndex = mo_QueryParams.KeyByIndex(p_ZeroBasedIndex)
End Property

Public Property Get KeyCount() As Long
   KeyCount = mo_QueryParams.Count
End Property

Public Property Get ValueCountByKey(ByVal p_Key As String) As Long
   ValueCountByKey = ValuesByKey(p_Key).Count
End Property

Public Property Get ValuesByIndex(ByVal p_ZeroBasedIndex As Long) As CHttpQueryParamValues
   Set ValuesByIndex = mo_QueryParams.ItemByIndex(p_ZeroBasedIndex)
End Property

Public Sub ParseQueryParams(po_FcgiParams As CFcgiParams, po_FcgiStdin As CFcgiStdIn)
   Dim l_QueryString As String
   Dim la_QueryPairs() As String
   Dim la_QueryPair() As String
   Dim lo_Values As CHttpQueryParamValues
   
   Dim ii As Long
   Dim jj As Long
   
   ' ********** Start of procedure
   
   mo_QueryParams.RemoveAll
   
   If po_FcgiParams.ExistsByEnum(stdparam_QueryString) Then
      l_QueryString = po_FcgiParams.ValueByEnum(stdparam_QueryString)
   End If
   
   Select Case UCase$(po_FcgiParams.ValueByEnum(stdparam_RequestMethod))
   Case "POST"
      ' May have parameters in body
      Select Case LCase$(Left$(po_FcgiParams.ValueByEnum(stdparam_ContentType), Len("application/x-www-form-urlencoded")))
      Case "application/x-www-form-urlencoded"
      
         ' Add body to querystring
         apiOutputDebugString "application/x-www-form-urlencoded content - will parse query parameters"
         apiOutputDebugString "POST body: " & StrConv(po_FcgiStdin.Content, vbUnicode)
         
         l_QueryString = l_QueryString & StrConv(po_FcgiStdin.Content, vbUnicode)
      End Select
   End Select
   
   If stringIsEmptyOrWhitespaceOnly(l_QueryString) Then Exit Sub
   
   If InStr(1, l_QueryString, ";") Then
      ' Accept ; or & as pair separators
      l_QueryString = Replace(l_QueryString, ";", "&")
   End If
   
   la_QueryPairs = Split(l_QueryString, "&")
   For ii = 0 To UBound(la_QueryPairs)
      
      Set lo_Values = Nothing
      la_QueryPair = Split(la_QueryPairs(ii), "=")
      
      Select Case UBound(la_QueryPair)
      Case 0
         ' Default argument, not a pair
         la_QueryPair(0) = libCrypt.URLDecode(la_QueryPair(0))
         
         For jj = 0 To mo_QueryParams.Count - 1
            If IsEmpty(mo_QueryParams.KeyByIndex(jj)) Then
               Set lo_Values = mo_QueryParams.ItemByIndex(jj)
               
               mo_QueryParams.RemoveByIndex jj
               
               Exit For
            End If
         Next jj
         
         If lo_Values Is Nothing Then
            Set lo_Values = New CHttpQueryParamValues
         End If
         
         lo_Values.Add la_QueryPair(0)
         
         mo_QueryParams.Add lo_Values
      
      Case 1
         ' Key/Value pair
         ' There's no definitive spec on what to do with duplicate keys.
         la_QueryPair(0) = Trim$(libCrypt.URLDecode(la_QueryPair(0)))
         la_QueryPair(1) = libCrypt.URLDecode(la_QueryPair(1))
         
         For jj = 0 To mo_QueryParams.Count - 1
            If mo_QueryParams.KeyByIndex(jj) = la_QueryPair(0) Then
               ' Key exists, remove it
               Set lo_Values = mo_QueryParams.ItemByIndex(jj)
               
               mo_QueryParams.RemoveByIndex jj
               
               Exit For
            End If
         Next jj
         
         If lo_Values Is Nothing Then
            Set lo_Values = New CHttpQueryParamValues
         End If
         
         lo_Values.Add la_QueryPair(1)
         
         mo_QueryParams.Add lo_Values, la_QueryPair(0)
      
      Case Else
         ' Huh?
         Debug.Assert False
         apiOutputDebugString "WARNING: Bad query string parameter: " & la_QueryPairs(ii)
      End Select
   Next ii
End Sub

Private Sub Class_Initialize()
   Set mo_QueryParams = libRc5Factory.C.Collection(False, BinaryCompare)
End Sub

