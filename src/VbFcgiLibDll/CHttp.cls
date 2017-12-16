VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHttp"
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

Private mo_Cookies As VbFcgiLib.CHttpCookies
Private mo_QueryParams As VbFcgiLib.CHttpQueryParams
Private mo_Upstream As VbFcgiLib.CFcgiUpstream

Friend Property Set FcgiUpstream(po_Upstream As VbFcgiLib.CFcgiUpstream)
   Set mo_Upstream = po_Upstream
End Property

Public Property Get Cookies(Optional pv_CookieDomain As Variant) As VbFcgiLib.CHttpCookies
   ' The optional pv_CookieDomain parameter must be a string if passed
   ' If MISSING/Not Passed, pv_CookieDomain will default to an empt string if no domain has ever been passed
   ' OR if a domain has been previously passed, it will default to the last passed cookie domain
   
   Static s_LastDomain As String
   
   Dim l_Domain As String
   Dim l_New As Boolean
   
   If mo_Upstream Is Nothing Then Err.Raise vbObjectError, , "FCGI upstream object must be set to access cookies."
   
   If mo_Cookies Is Nothing Then
      ' Initialize on first call
      Set mo_Cookies = New VbFcgiLib.CHttpCookies
      l_New = True   ' Indicate we need to build the cookies collection on the first pass
   End If
   
   If IsMissing(pv_CookieDomain) Then
      ' No cookie domain passed, use last passed domain
      l_Domain = s_LastDomain
   
   Else
      If VarType(pv_CookieDomain) <> vbString Then Err.Raise 5, , "Cookie domain must be a string."
      
      ' Normalize the passed domain
      ' Domains are case insensitive and can't include whitespace
      l_Domain = LCase$(stringRemoveWhitespace(pv_CookieDomain))
   End If
   
   If (l_Domain <> s_LastDomain) Or l_New Then
      ' Build cookies collection for newly created object or changed domain
      s_LastDomain = l_Domain
      
      mo_Cookies.ParseCookies mo_Upstream.Params, s_LastDomain
   End If
   
   Set Cookies = mo_Cookies
End Property

Public Property Get QueryParameters() As VbFcgiLib.CHttpQueryParams
   If mo_QueryParams Is Nothing Then
      ' Initialize on first call
      Set mo_QueryParams = New VbFcgiLib.CHttpQueryParams
      If Not mo_Upstream Is Nothing Then
         mo_QueryParams.ParseQueryParams mo_Upstream.Params, mo_Upstream.Stdin
      End If
   End If
   
   Set QueryParameters = mo_QueryParams
End Property
