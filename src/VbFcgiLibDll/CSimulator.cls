VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mo_Downstream As VbFcgiLib.CFcgiDownstream

Public Sub SimulateRequest(ByVal p_Url As String, po_FcgiApp As VbFcgiLib.IFcgiApp)
   Dim lo_Request As VbFcgiLib.CFcgiRequest
   Dim lo_Response As VbFcgiLib.CFcgiResponse
   Dim la_Url() As String
   Dim ii As Long
   Dim l_ChunkPos As Long
   
   Set lo_Request = New VbFcgiLib.CFcgiRequest
   Set lo_Response = New VbFcgiLib.CFcgiResponse
   
   ' Cleanup URL
   p_Url = Replace$(p_Url, " ", "%20")
   p_Url = Replace$(p_Url, "\", "/")
   
   l_ChunkPos = InStr(1, p_Url, "?")
   If l_ChunkPos > 1 Then
      mo_Downstream.Params.ValueByEnum(stdparam_QueryString) = Mid$(p_Url, l_ChunkPos + 1)
      
      ' Chop query
      p_Url = Left$(p_Url, l_ChunkPos - 1)
   End If
   
   l_ChunkPos = InStr(1, p_Url, "://")
   If l_ChunkPos < 1 Then
      mo_Downstream.Params.ValueByEnum(stdparam_RequestScheme) = "http"
   Else
      mo_Downstream.Params.ValueByEnum(stdparam_RequestScheme) = LCase$(Left$(p_Url, l_ChunkPos - 1))
      
      ' Chop protocol
      p_Url = Mid$(p_Url, l_ChunkPos + 3)
   End If
   
   ' Remove trailing "/" if exists
   If Right$(p_Url, 1) = "/" Then p_Url = Left$(p_Url, Len(p_Url) - 1)
   
   la_Url = Split(p_Url, "\")
   
   ' Check for port
   l_ChunkPos = InStr(1, la_Url(0), ":")
   If l_ChunkPos > 0 Then
      mo_Downstream.Params.ValueByEnum(stdparam_ServerPort) = Mid$(la_Url(0), l_ChunkPos + 1)
      
      ' Strip port
      la_Url(0) = Left$(la_Url(0), l_ChunkPos - 1)
   Else
      Select Case LCase$(mo_Downstream.Params.ValueByEnum(stdparam_RequestMethod))
      Case "https"
         mo_Downstream.Params.ValueByEnum(stdparam_ServerPort) = 443
      Case Else
         mo_Downstream.Params.ValueByEnum(stdparam_ServerPort) = 80
      End Select
   End If

   ' Remove user name if exists
   l_ChunkPos = InStr(1, la_Url(0), "@")
   If l_ChunkPos > 0 Then
      la_Url(0) = Mid$(la_Url(0), l_ChunkPos + 1)
   End If

   ' Init HTTP_HOST parameter
   mo_Downstream.Params.ValueByEnum(stdparam_HttpHost) = la_Url(0)
   
   ' Init DOCUMENT_URI, REQUEST_URI, SCRIPTNAME
   ' TODO: Make sure this works for all kinds of passed URLs...there undoubtedly issues with some
   ' URLS in the current for (for example,
   mo_Downstream.Params.ValueByEnum(stdparam_ScriptName) = "/" & la_Url(UBound(la_Url))
   mo_Downstream.Params.ValueByEnum(stdparam_DocumentUri) = "/"
   For ii = 1 To UBound(la_Url)
      mo_Downstream.Params.ValueByEnum(stdparam_DocumentUri) = mo_Downstream.Params.ValueByEnum(stdparam_DocumentUri) & la_Url(ii)
   Next ii
   mo_Downstream.Params.ValueByEnum(stdparam_RequestUri) = mo_Downstream.Params.ValueByEnum(stdparam_DocumentUri) & "?" & mo_Downstream.Params.ValueByEnum(stdparam_QueryString)
   
   mo_Downstream.Params.State = paramstate_Built
   
   ' Parse HTTP query params
   Set lo_Request.FcgiDownstream = mo_Downstream
   
   lo_Request.Http.QueryParameters.ParseQueryParams mo_Downstream.Params, Nothing
   
   mo_Downstream.State = fcgireqstate_Ready
   Set lo_Request.FcgiDownstream = mo_Downstream
   
   po_FcgiApp.ProcessRequest lo_Request, lo_Response
End Sub

Private Sub Class_Initialize()
   Set mo_Downstream = New VbFcgiLib.CFcgiDownstream
   
   mo_Downstream.Initialize 0, 0
   
   With mo_Downstream.Params
      .State = paramstate_Building
      
      ' Set up default FCGI parameter values
      
      .ValueByEnum(stdparam_ServerSoftware) = "nginx/1.13.7"
      .ValueByEnum(stdparam_RequestMethod) = "GET"
      .ValueByEnum(stdparam_GatewayInterface) = "CGI/1.1"
      .ValueByEnum(stdparam_ServerProtocol) = "HTTP/1.1"
      .ValueByEnum(stdparam_RemoteAddr) = "127.0.0.1"
      .ValueByEnum(stdparam_RemotePort) = Int(Rnd * 32767) + 32767   ' Random remote port, not used for anything
      .ValueByEnum(stdparam_DocumentRoot) = Replace$(pathBin, "\", "/") & "nginx/html"
      .ValueByEnum(stdparam_RedirectStatus) = 200
      .ValueByEnum(stdparam_HttpUserAgent) = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:58.0) Gecko/20100101 Firefox/58.0"
      .ValueByEnum(stdparam_HttpAccept) = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
      .ValueByEnum(stdparam_HttpAcceptEncoding) = "gzip, deflate"
      .ValueByEnum(stdparam_HttpAcceptLanguage) = "en-US,en;q=0.5"
      .ValueByEnum(stdparam_HttpConnection) = "keep-alive"
      .ValueByEnum(stdparam_ServerName) = "localhost"
      .ValueByEnum(stdparam_ServerPort) = "80"
   End With
End Sub

