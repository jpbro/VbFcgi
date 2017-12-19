VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFcgiParams"
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

' This class holds a collection of CGI parameters (and related methods)
' that were received from web server

Public Enum e_FcgiParametersState
   paramstate_Initialized  ' Nothing has happened yet
   paramstate_Building  ' We've received some parameter data, but not the closing record
   paramstate_Built  ' The parameters collection is complete
End Enum

Public Enum e_StandardParameters ' Enum of standard CGI parameters for easier retrieval
   stdparam_PathInfo
   stdparam_PathTranslated
   stdparam_QueryString
   stdparam_RequestMethod
   stdparam_ContentType
   stdparam_ContentLength
   stdparam_ScriptName
   stdparam_RequestUri
   stdparam_DocumentUri
   stdparam_DocumentRoot
   stdparam_ServerProtocol
   stdparam_GatewayInterface
   stdparam_ServerSoftware
   stdparam_RemoteAddr
   stdparam_RemotePort
   stdparam_ServerAddr
   stdparam_ServerPort
   stdparam_ServerName
   stdparam_RedirectStatus
   stdparam_HttpHost
   stdparam_HttpUserAgent
   stdparam_HttpAccept
   stdparam_HttpAcceptLanguage
   stdparam_HttpAcceptEncoding
   stdparam_HttpConnection
   stdparam_HttpCacheControl
   stdparam_HttpCookie
   stdparam_HttpIfNoneMatch
End Enum

Private mo_Params As vbRichClient5.cCollection  ' Parameters collection (key-value pairs)
Private m_State As e_FcgiParametersState ' The current state of the parameter records retrieval from the web server

Public Property Get State() As e_FcgiParametersState
   State = m_State
End Property

Public Property Let State(ByVal p_State As e_FcgiParametersState)
   If p_State = paramstate_Initialized Then Err.Raise 5, , "Can't re-initialize parameters object."
   If p_State = paramstate_Building And m_State = paramstate_Built Then Err.Raise 5, , "Can't re-open parameters collection."
   
   m_State = p_State
End Property

Private Sub Class_Initialize()
   Set mo_Params = libRc5Factory.C.Collection(False, TextCompare)
End Sub

Public Function Exists(ByVal p_Key As String) As Boolean
   Exists = mo_Params.Exists(p_Key)
End Function
   
Public Function ExistsByEnum(ByVal p_StandardParameterEnum As VbFcgiLib.e_StandardParameters) As Boolean
   ExistsByEnum = mo_Params.Exists(EnumParameterKey(p_StandardParameterEnum))
End Function
   
Public Function Count() As Long
   Count = mo_Params.Count
End Function

Public Function KeyByIndex(ByVal p_ZeroBasedIndex As Long) As String
   KeyByIndex = mo_Params.KeyByIndex(p_ZeroBasedIndex)
End Function

Public Property Get ValueByIndex(ByVal p_ZeroBasedIndex As Long) As String
   ValueByIndex = mo_Params.ItemByIndex(p_ZeroBasedIndex)
End Property

Public Property Let ValueByIndex(ByVal p_ZeroBasedIndex As Long, ByVal p_Value As String)
   mo_Params.ItemByIndex(p_ZeroBasedIndex) = p_Value
End Property

Public Property Get ValueByKey(ByVal p_Key As String) As String
   If mo_Params.Exists(p_Key) Then
      ValueByKey = mo_Params.Item(p_Key)
   Else
      ValueByKey = Empty
   End If
End Property

Public Property Let ValueByKey(ByVal p_Key As String, ByVal p_Value As String)
   If mo_Params.Exists(p_Key) Then
      mo_Params.Item(p_Key) = p_Value
   Else
      mo_Params.Add p_Value, p_Key
   End If
End Property

Public Property Get ValueByEnum(ByVal p_ParameterKeyEnum As e_StandardParameters) As String
   ValueByEnum = ValueByKey(EnumParameterKey(p_ParameterKeyEnum))
End Property

Public Property Let ValueByEnum(ByVal p_ParameterKeyEnum As e_StandardParameters, ByVal p_Value As String)
   ValueByKey(EnumParameterKey(p_ParameterKeyEnum)) = p_Value
End Property

Private Function EnumParameterKey(ByVal p_ParameterKeyEnum As e_StandardParameters)
   ' Convert a standard FCGI parameter enum into a key
   
   Select Case p_ParameterKeyEnum
   Case stdparam_PathInfo
      EnumParameterKey = "PATH_INFO"
      
   Case stdparam_PathTranslated
      EnumParameterKey = "PATH_TRANSLATED"
      
   Case stdparam_QueryString
      EnumParameterKey = "QUERY_STRING"
   
   Case stdparam_RequestMethod
      EnumParameterKey = "REQUEST_METHOD"
   
   Case stdparam_ContentType
      EnumParameterKey = "CONTENT_TYPE"
   
   Case stdparam_ContentLength
      EnumParameterKey = "CONTENT_LENGTH"
   
   Case stdparam_ScriptName
      EnumParameterKey = "SCRIPT_NAME"
   
   Case stdparam_RequestUri
      EnumParameterKey = "REQUEST_URI"
   
   Case stdparam_DocumentUri
      EnumParameterKey = "DOCUMENT_URI"
   
   Case stdparam_DocumentRoot
      EnumParameterKey = "DOCUMENT_ROOT"
   
   Case stdparam_ServerProtocol
      EnumParameterKey = "SERVER_PROTOCOL"
   
   Case stdparam_GatewayInterface
      EnumParameterKey = "GATEWAY_INTERFACE"
   
   Case stdparam_ServerSoftware
      EnumParameterKey = "SERVER_SOFTWARE"
   
   Case stdparam_RemoteAddr
      EnumParameterKey = "REMOTE_ADDR"
   
   Case stdparam_RemotePort
      EnumParameterKey = "REMOTE_PORT"
   
   Case stdparam_ServerAddr
      EnumParameterKey = "SERVER_ADDR"
   
   Case stdparam_ServerPort
      EnumParameterKey = "SERVER_PORT"
   
   Case stdparam_ServerName
      EnumParameterKey = "SERVER_NAME"
   
   Case stdparam_RedirectStatus
      EnumParameterKey = "REDIRECT_STATUS"
   
   Case stdparam_HttpHost
      EnumParameterKey = "HTTP_HOST"
   
   Case stdparam_HttpUserAgent
      EnumParameterKey = "HTTP_USER_AGENT"
   
   Case stdparam_HttpAccept
      EnumParameterKey = "HTTP_ACCEPT"
   
   Case stdparam_HttpAcceptLanguage
      EnumParameterKey = "HTTP_ACCEPT_LANGUAGE"
   
   Case stdparam_HttpAcceptEncoding
      EnumParameterKey = "HTTP_ACCEPT_ENCODING"
   
   Case stdparam_HttpConnection
      EnumParameterKey = "HTTP_CONNECTION"
   
   Case stdparam_HttpCacheControl
      EnumParameterKey = "HTTP_CACHE_CONTROL"
   
   Case stdparam_HttpCookie
      EnumParameterKey = "HTTP_COOKIE"
   
   Case stdparam_HttpIfNoneMatch
      EnumParameterKey = "HTTP_IF_NONE_MATCH"
   
   Case Else
      Debug.Assert False
      Err.Raise 5, , "Unknown Parameter Key Enum #: " & p_ParameterKeyEnum
   End Select
End Function

