VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHttpCookies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mo_Cookies As vbRichClient5.cCollection

Public Property Get Exists(ByVal p_Key As String) As Boolean
   Exists = mo_Cookies.Exists(p_Key)
End Property

Public Property Get CookieHeaders() As String
   ' Return a SetCookie line for each Cookie in the collection
   Dim lo_CookieHeaders As vbRichClient5.cStringBuilder
   Dim lo_Cookie As CHttpCookie
   Dim ii As Long
   
   Set lo_CookieHeaders = libRc5Factory.C.StringBuilder
   
   For ii = 0 To mo_Cookies.Count - 1
      Set lo_Cookie = Me.CookieByIndex(ii)
      
      With lo_CookieHeaders
         .Append "Set-Cookie: "
         .Append libCrypt.URLEncode(Me.KeyByIndex(ii))
         .Append "="
         .Append libCrypt.URLEncode(lo_Cookie.Value)
         
         If stringIsEmptyOrWhitespaceOnly(lo_Cookie.Domain) Then
            ' Empty domain
         Else
            .Append "; Domain="
            .Append lo_Cookie.Domain
         End If
         
         If stringIsEmptyOrWhitespaceOnly(lo_Cookie.Path) Then
            ' Cookie path is empty.
         Else
            .Append "; Path="
            .Append lo_Cookie.Path
         End If
         
         If lo_Cookie.ExpiresGmt > 0 Then
            .Append "; Expires="
            .Append dateUtcToCookieDate(lo_Cookie.ExpiresGmt)
         End If
         
         If lo_Cookie.Secure Then
            .Append "; Secure"
         End If
         
         If lo_Cookie.HttpOnly Then
            .Append "; HttpOnly"
         End If
         
         .AppendNL ""
      End With
   Next ii
   
   CookieHeaders = lo_CookieHeaders.ToString
End Property

Public Sub ParseCookies(po_FcgiParams As CFcgiParams, ByVal p_CookieDomain As String)
   Dim l_Cookies As String
   Dim la_Cookies() As String
   Dim la_Cookie() As String
   Dim lo_Cookie As VbFcgiLib.CHttpCookie
   
   Dim ii As Long
   Dim jj As Long
   
   ' ********** Start of procedure
   
   mo_Cookies.RemoveAll
   
   If Not po_FcgiParams.ExistsByEnum(stdparam_HttpCookie) Then
      apiOutputDebugString "No cookie header found in request."
      Exit Sub
   End If
   
   l_Cookies = Trim$(po_FcgiParams.ValueByEnum(stdparam_HttpCookie))
   
   If LenB(l_Cookies) = 0 Then
      apiOutputDebugString "Cookies parameter was found but is empty."
      Exit Sub
   End If
      
   apiOutputDebugString "Parsing Cookies: " & l_Cookies
      
   la_Cookies = Split(l_Cookies, ";")
   For ii = 0 To UBound(la_Cookies)
      la_Cookie = Split(la_Cookies(ii), "=")
      
      Select Case UBound(la_Cookie)
      Case 1
         ' Key/Value pair
         ' There's no definitive spec on what to do with duplicate keys.
         la_Cookie(0) = Trim$(libCrypt.URLDecode(la_Cookie(0)))
         
         If mo_Cookies.Exists(la_Cookie(0)) Then
            ' Duplicate!
            Debug.Assert False
            apiOutputDebugString "WARNING: Duplicate cookie key found! Key: " & la_Cookie(0)
            
         Else
            la_Cookie(1) = libCrypt.URLDecode(la_Cookie(1))
            
            Set lo_Cookie = New VbFcgiLib.CHttpCookie
            With lo_Cookie
               .Value = la_Cookie(1)
            End With
            
            lo_Cookie.Domain = p_CookieDomain
            lo_Cookie.Path = "/"
            lo_Cookie.Secure = True
            
            mo_Cookies.Add lo_Cookie, la_Cookie(0)
         End If
      
      Case Else
         ' Huh?
         Debug.Assert False
         apiOutputDebugString "WARNING: Bad cookie: " & la_Cookies(ii)
         
      End Select
   Next ii
End Sub

Public Sub AddOrReplaceCookie(ByVal p_Key As String, ByVal p_Value As String, Optional ByVal p_Domain As String, Optional ByVal p_Path As String, Optional ByVal p_ExpiryDateGmt As Date = 0, Optional ByVal p_Secure As Boolean)
   Dim lo_Cookie As CHttpCookie
   
   Set lo_Cookie = New VbFcgiLib.CHttpCookie
   With lo_Cookie
      .Domain = p_Domain
      .ExpiresGmt = p_ExpiryDateGmt
      .Path = p_Path
      .Secure = p_Secure
      .Value = p_Value
   End With
   
   If mo_Cookies.Exists(p_Key) Then
      mo_Cookies.Remove p_Key
   End If
   
   mo_Cookies.Add lo_Cookie, p_Key
End Sub

Public Property Get CookieByKey(ByVal p_Key As String) As CHttpCookie
   Set CookieByKey = mo_Cookies.Item(p_Key)
End Property

Public Property Get CookieCount() As Long
   CookieCount = mo_Cookies.Count
End Property

Public Property Get CookieByIndex(ByVal p_ZeroBasedIndex As Long) As CHttpCookie
   Set CookieByIndex = mo_Cookies.ItemByIndex(p_ZeroBasedIndex)
End Property

Public Property Get KeyByIndex(ByVal p_ZeroBasedIndex As Long) As String
   KeyByIndex = mo_Cookies.KeyByIndex(p_ZeroBasedIndex)
End Property

Private Sub Class_Initialize()
   Set mo_Cookies = libRc5Factory.C.Collection(False, BinaryCompare)
End Sub

