VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFcgiServer"
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

' This class is the main FCGI listener/processor

' Internal Error codes
Public Enum e_FcgiError
   [_fcgierr_Base] = vbObjectError + &H1000&
   
   fcgierr_UnexpectedError = [_fcgierr_Base]
   fcgierr_UnsupportedRole
   fcgierr_Overloaded
   fcgierr_FcgiThreadInitializationError  ' Could not initialize the thread handler
   fcgierr_NotReadyForResponse
   fcgierr_ContentLengthMismatch  ' CGI Content-Length value doesn't match STDIN content length
End Enum

' Object Variables
Private WithEvents mo_TcpServer As vbRichClient5.cTCPServer ' TCP server for receiving FASTCGI data
Attribute mo_TcpServer.VB_VarHelpID = -1
Private WithEvents mo_ResponseTimer As vbRichClient5.cTimer
Attribute mo_ResponseTimer.VB_VarHelpID = -1
Private WithEvents mo_ActiveDownstream As VbFcgiLib.CFcgiResponse
Attribute mo_ActiveDownstream.VB_VarHelpID = -1
Private mo_ActiveRequests As vbRichClient5.cCollection   ' Collection of active requests

' Listener Variables
Private m_ListenPort As Long

Private mo_FcgiApps As vbRichClient5.cCollection

Public Function IsListening() As Boolean
   If mo_TcpServer Is Nothing Then Exit Function
   
   IsListening = (m_ListenPort <> 0)
End Function

Public Sub StopListening()
   m_ListenPort = 0
   Set mo_TcpServer = Nothing
End Sub

Public Sub StartListening(ByVal p_Host As String, ByVal p_Port As Long)
   ' This method must be called to start the listener
      
   ' Destroy old listener
   Me.StopListening
   
   If (p_Port < 1) Or (p_Port > 49151) Then Err.Raise 5, , "Port must be in 1-65535 range."
   
   apiOutputDebugString "Starting TCP Server on " & p_Host & ":" & p_Port
   
   Set mo_TcpServer = libRc5Factory.C.TCPServer
   If mo_TcpServer.Listen(p_Host, p_Port) = 0 Then
      Err.Raise vbObjectError, , "Could not start TCP listener on " & p_Host & ":" & p_Port
   End If
   
   apiOutputDebugString "Started TCP Server on " & p_Host & ":" & p_Port
   
   m_ListenPort = p_Port
End Sub

Private Sub Class_Initialize()
   
   On Error GoTo ErrorHandler
   
   apiOutputDebugString "Initializing " & TypeName(Me)
   
   Set mo_ActiveRequests = libRc5Factory.C.Collection
      
   ' Make sure you have a DLL called vbFcgiApp.dll
   ' With a class called CFcgiApp that implements the VbFcgiHost.dll IFcgiApp interface
   Set mo_FcgiApps = libRc5Factory.C.Collection
   
   ' Create a timer for building responses (disabled at startup)
   Set mo_ResponseTimer = libRc5Factory.C.Timer(15, False)
   
   apiOutputDebugString "Initialized " & TypeName(Me)
   
   Exit Sub

ErrorHandler:
   apiOutputDebugString "Initialization error #" & Err.Number & " - " & Err.Description
   
   Err.Raise vbObjectError, , "Initialization Error."
End Sub

Public Property Get Port() As Long
   Port = m_ListenPort
End Property

Private Function RequestKey(ByVal p_RequestId As Integer) As String
   ' Generate a key string for the passed request ID
   
   RequestKey = "REQ_XXXXXX"
   Mid$(RequestKey, Len(RequestKey) - 5) = Format$(p_RequestId, "000000;-00000")
End Function

Private Sub mo_ActiveDownstream_SendError(ByVal p_RequestId As Long, ByVal p_ErrorNumber As Long, ByVal p_ErrorDescription As String)
   Dim l_RequestKey As String
   Dim lo_Downstream As CFcgiDownstream
   
   ' There was an error building the response
   apiOutputDebugString "Error building response for ID #" & p_RequestId & ". Error Code " & p_ErrorNumber & ", Description: " & p_ErrorDescription
   
   l_RequestKey = RequestKey(p_RequestId)
   If mo_ActiveRequests.Exists(l_RequestKey) Then
      Set lo_Downstream = mo_ActiveRequests.Item(l_RequestKey)
   End If
   
   If lo_Downstream Is Nothing Then
      apiOutputDebugString "Request object not found for ID#" & p_RequestId
   Else
      lo_Downstream.SetError p_ErrorNumber, p_ErrorDescription
      
      SendStdErrByRequestNoErr lo_Downstream
   End If
End Sub

Private Sub mo_ActiveDownstream_SendFinished(ByVal p_RequestId As Long)
   Dim l_RequestKey As String
   Dim lo_Downstream As CFcgiDownstream
   
   On Error GoTo ErrorHandler
   
   apiOutputDebugString "Finished building response for ID#" & p_RequestId
   
   l_RequestKey = RequestKey(p_RequestId)
   If mo_ActiveRequests.Exists(l_RequestKey) Then
      Set lo_Downstream = mo_ActiveRequests.Item(l_RequestKey)
   End If
   
   If lo_Downstream Is Nothing Then
      apiOutputDebugString "Request object not found for ID#" & p_RequestId
   Else
      lo_Downstream.StdOut.State = fcgistdoutstate_Built
      lo_Downstream.State = fcgireqstate_Ready
   
      If fcgiFlushStdOut(mo_TcpServer, lo_Downstream.SocketNumber, lo_Downstream.Id, lo_Downstream.StdOut, True) Then
         ' Flushed - send FCGI_END_REQUEST
         fcgiSendEndRequest mo_TcpServer, lo_Downstream.SocketNumber, lo_Downstream.Id, 0, 0
      Else
         apiOutputDebugString "STDOUT could not be flushed in finished event for Request #" & lo_Downstream.Id
         
         lo_Downstream.State = fcgireqstate_Error
      End If
      
      RemoveRequestNoErr lo_Downstream
   End If
   
   Exit Sub

ErrorHandler:
   apiOutputDebugString "Error in Finished event: " & Err.Number & " " & Err.Description
End Sub

Private Sub mo_ActiveDownstream_SendBytes(ByVal p_RequestId As Long, pa_Bytes() As Byte)
   ' Append the content to STDOUT
   Dim l_RequestKey As String
   Dim lo_Downstream As VbFcgiLib.CFcgiDownstream
   
   On Error GoTo ErrorHandler
   
   apiOutputDebugString "Received response bytes for ID#" & p_RequestId
   
   l_RequestKey = RequestKey(p_RequestId)
   If mo_ActiveRequests.Exists(l_RequestKey) Then
      Set lo_Downstream = mo_ActiveRequests.Item(l_RequestKey)
   End If
   
   If lo_Downstream Is Nothing Then
      apiOutputDebugString "Request object not found for ID#" & p_RequestId
   Else
      lo_Downstream.StdOut.AppendContent pa_Bytes
   
      If Not fcgiFlushStdOut(mo_TcpServer, lo_Downstream.SocketNumber, lo_Downstream.Id, lo_Downstream.StdOut, False) Then
         apiOutputDebugString "STDOUT could not be flushed in WriteBytes event for Request #" & lo_Downstream.Id
         
         lo_Downstream.State = fcgireqstate_Error
      
         RemoveRequestNoErr lo_Downstream
      End If
   End If
   
   Exit Sub

ErrorHandler:
   apiOutputDebugString "Error in WriteBytes event: " & Err.Number & " " & Err.Description
End Sub

Private Sub mo_TcpServer_DataArrival(ByVal hSocket As Long, ByVal BytesTotal As Long, ByVal FirstBufferAfterOverflow As Boolean)
   Dim la_Payload() As Byte
   Dim lt_Record As FCGX_RECORD
   Dim lt_BeginRequestBody As FCGX_BEGIN_REQUEST_BODY
   Dim lt_EndRequestBody As FCGX_END_REQUEST_BODY
   Dim l_RequestKey As String
   Dim l_RecordPos As Long
   Dim lo_Downstream As VbFcgiLib.CFcgiDownstream
   Dim l_Param As String
   Dim l_Value As String
   Dim la_Param() As Byte
   Dim la_Value() As Byte
   Dim la_Content() As Byte
   Dim l_ParamLen As Long
   Dim l_ValueLen As Long
   Dim l_Offset As Long
   
   On Error GoTo ErrorHandler

   mo_ResponseTimer.Enabled = False

   libRc5Factory.C.Timing True

   apiOutputDebugString "Received Request on socket " & hSocket & ", port: " & Me.Port & ". Bytes: " & BytesTotal
   
   If BytesTotal < 8 Then
      ' Not an FCGI record
      Debug.Assert False

      apiOutputDebugString "Received too few bytes for FCGI request. Bytes: " & BytesTotal
      
      If BytesTotal > 0 Then
         mo_TcpServer.GetData hSocket, VarPtr(la_Payload(0)), BytesTotal
      End If

      GoTo Cleanup
   End If

   ' Get entire payload - may be multiple FCGI records
   ReDim la_Payload(BytesTotal - 1)

   mo_TcpServer.GetData hSocket, VarPtr(la_Payload(0)), BytesTotal

   ' Process FCGI records
   Do While l_RecordPos < BytesTotal
      Set lo_Downstream = Nothing

      ' Process FCGI record version
      lt_Record.Version = la_Payload(l_RecordPos)
      If lt_Record.Version <> 1 Then
         ' Unsupported Version
      End If

      ' Process Request ID
      apiCopyMemory lt_Record.RequestId, la_Payload(l_RecordPos + 2), 2
      lt_Record.RequestId = apiNtohs(lt_Record.RequestId)
      l_RequestKey = RequestKey(lt_Record.RequestId)  ' Key for our active requests collection

      ' Process Content and Padding Lengths
      apiCopyMemory lt_Record.ContentLength, la_Payload(l_RecordPos + 4), 2
      lt_Record.ContentLength = apiNtohs(lt_Record.ContentLength)

      If lt_Record.ContentLength > 0 Then
         ReDim la_Content(lt_Record.ContentLength - 1)
         apiCopyMemory la_Content(0), la_Payload(l_RecordPos + 8), lt_Record.ContentLength
      Else
         Erase la_Content
      End If

      lt_Record.PaddingLength = la_Payload(l_RecordPos + 6)

      ' Process Record type
      lt_Record.RecordType = la_Payload(l_RecordPos + 1)
      Select Case lt_Record.RecordType
      Case FCGI_BEGIN_REQUEST
         If mo_ActiveRequests.Exists(l_RequestKey) Then
            ' Stale request still exists in active queue - remove it
            mo_ActiveRequests.Remove l_RequestKey
         End If

         ' Create request object and add it to active requests collection
         Set lo_Downstream = New VbFcgiLib.CFcgiDownstream  ' GetFreeRequest
         lo_Downstream.Initialize lt_Record.RequestId, hSocket
         mo_ActiveRequests.Add lo_Downstream, l_RequestKey
         
         If Not mo_ResponseTimer.Enabled Then
            mo_ResponseTimer.Enabled = True
         End If
         
         ' Build the request body
         apiCopyMemory lt_BeginRequestBody, la_Content(0), UBound(la_Content) + 1
         lt_BeginRequestBody.Role = apiNtohs(lt_BeginRequestBody.Role)     ' Swap byte order
                  
         ' Check for supported/unsupported role
         Select Case lt_BeginRequestBody.Role
         Case FCGI_RESPONDER
            ' Supported Role, build the request state object
            
         Case Else
            ' Unsupported Role
            lo_Downstream.SetError fcgierr_UnsupportedRole, "The passed role (" & lt_BeginRequestBody.Role & ") is not supported."
         
         End Select

      Case Else
         ' Not a BEGIN request
         
         If mo_ActiveRequests.Exists(l_RequestKey) Then
            ' Found active request - get our state object
            Set lo_Downstream = mo_ActiveRequests.Item(l_RequestKey)

            If lo_Downstream.State <> fcgireqstate_Error Then
               ' Request not in error state.
               Select Case lt_Record.RecordType
               Case FCGI_PARAMS
                  ' PARAMS request
                  l_Offset = 0
   
                  apiOutputDebugString "Found FCGI_PARAMS"
   
                  If lt_Record.ContentLength = 0 Then
                     ' End of FCGI_PARAMS records, mark parameters object as finished
                     apiOutputDebugString "All FCGI Params content decoded. RecordPos: " & l_RecordPos & ", Bytes Total: " & BytesTotal
                     
                     lo_Downstream.Params.State = paramstate_Built
      
                  Else
                     ' Keys and Value pair length variables will be either 1 or 4 bytes long
                     ' If the first byte and &H80 <> 0 then it will be 4 bytes, otherwise 1 byte
                     ' Get Parameter Length
                     lo_Downstream.Params.State = paramstate_Building
   
                     Do While l_RecordPos + l_Offset < lt_Record.ContentLength
                        apiOutputDebugString "Records pos: " & l_RecordPos + l_Offset + 8 & ", Content Len: " & lt_Record.ContentLength
                        
                        If la_Payload(l_RecordPos + 8 + l_Offset) And &H80 Then
                           ' 4 byte length
                           apiOutputDebugString "Found "
                           
                           apiCopyMemory l_ParamLen, la_Payload(l_RecordPos + 8 + l_Offset), 4
                           l_ParamLen = apiNtohl(l_ParamLen) Xor &H80000000
   
                           l_Offset = l_Offset + 4
                        Else
                           ' 1 byte length
                           l_ParamLen = CLng(la_Payload(l_RecordPos + 8 + l_Offset))
   
                           apiOutputDebugString "Param Len: " & l_ParamLen
   
                           l_Offset = l_Offset + 1
                        End If
   
                        ' Get Value Length
                        If la_Payload(l_RecordPos + 8 + l_Offset) And &H80 Then
                           ' 4 byte length
                           apiCopyMemory l_ValueLen, la_Payload(l_RecordPos + 8 + l_Offset), 4
                           l_ValueLen = apiNtohl(l_ValueLen) Xor &H80000000
   
                           l_Offset = l_Offset + 4
                        Else
                           ' 1 byte length
                           l_ValueLen = CLng(la_Payload(l_RecordPos + 8 + l_Offset))
   
                           apiOutputDebugString "Value Len: " & l_ValueLen
   
                           l_Offset = l_Offset + 1
                        End If
   
                        If l_ParamLen > 0 Then
                           apiOutputDebugString "Param len: " & l_ParamLen
                           
                           ReDim la_Param(l_ParamLen - 1)
                           apiCopyMemory la_Param(0), la_Payload(l_RecordPos + 8 + l_Offset), l_ParamLen
                           l_Param = libCrypt.UTF8ToVBString(la_Param)
   
                           apiOutputDebugString "Param: " & l_Param
   
                           If l_ValueLen > 0 Then
                              apiOutputDebugString "Value Len: " & l_ValueLen
                           
                              ReDim la_Value(l_ValueLen - 1)
                              apiCopyMemory la_Value(0), la_Payload(l_RecordPos + 8 + l_Offset + l_ParamLen), l_ValueLen
                              l_Value = libCrypt.UTF8ToVBString(la_Value)
                              
                              apiOutputDebugString "Value: " & l_Value
                           
                           Else
                              l_Value = ""
                           End If
                           
                           Select Case UCase$(l_Param)
                           Case "PATH_INFO", "PATH_TRANSLATED", "DOCUMENT_ROOT"
                              ' Convert Unix path separator chars to Windows path separator chars
                              l_Value = Replace$(l_Value, "/", "\")
                           End Select
   
                           lo_Downstream.Params.ValueByKey(l_Param) = l_Value
   
                           apiOutputDebugString "Found Key/Value Pair: " & l_Param & " / " & l_Value
                           
                           ' Cleanup
                           'Erase la_Param
                           'Erase la_Value
                        End If
   
                        l_Offset = l_Offset + l_ParamLen + l_ValueLen
                     Loop
   
                  End If
               
               Case FCGI_STDIN
                  If lt_Record.ContentLength = 0 Then
                     ' STDIN terminator record
                     apiOutputDebugString "STDIN terminator record found - closing STDIN stream."
                     
                     lo_Downstream.Stdin.State = fcgistdinstate_Built
                     
                  Else
                     apiOutputDebugString "STDIN record found, appending content."
                     
                     ReDim la_Content(lt_Record.ContentLength - 1)
                     apiCopyMemory la_Content(0), la_Payload(l_RecordPos + 8), lt_Record.ContentLength
                     
                     lo_Downstream.Stdin.AppendContent la_Content
                     Erase la_Content
                  End If
               
               Case Else
                  ' TODO: Unsupported record type
                  Debug.Assert False
                  apiOutputDebugString "Unsupported request type: " & lt_Record.RecordType
   
               End Select
            
            Else
               ' Request in error state - skip this record
            End If
         Else
            ' Uhoh - got a non-begin record type for an unknown request!
            ' Spec says to ignore it
            apiOutputDebugString "Received non-begin record for an unknown request!"
         End If
      End Select

      apiOutputDebugString "Received Request. Version: " & lt_Record.Version & ", Type: " & lt_Record.RecordType
      apiOutputDebugString "Received Request. RequestID: " & lt_Record.RequestId
      apiOutputDebugString "Received Request. ContentLength: " & lt_Record.ContentLength & ", Padding Length: " & lt_Record.PaddingLength

      l_RecordPos = l_RecordPos + 8 + lt_Record.ContentLength + lt_Record.PaddingLength
   Loop
   
Cleanup:
   On Error Resume Next
   
   mo_ResponseTimer_Timer
   
   apiOutputDebugString "Response timing: " & libRc5Factory.C.Timing
   
   mo_ResponseTimer.Enabled = True
   
   Exit Sub

ErrorHandler:
   ' Something terrible happened, respond with STDERR and END packet
   
   apiOutputDebugString "Error in DataArrival: " & Err.Number & " " & Err.Description
   
   If lo_Downstream Is Nothing Then
      SendStdErrByInfoNoErr hSocket, lt_Record.RequestId, Err.Number, Err.Description
   Else
      SendStdErrByRequestNoErr lo_Downstream
   End If
   
   Resume Cleanup
End Sub

Private Sub mo_ResponseTimer_Timer()
         ' Build Response HTML for completed requests and send it to web server
         ' TODO: Send errors of STDERR and set END_REQUEST application/protocol status accordingly
         
         Static s_Busy As Boolean   ' Prevent re-entry

         Dim lo_Downstream As VbFcgiLib.CFcgiDownstream   ' Active request object
         Dim lo_FcgiRequest As VbFcgiLib.CFcgiRequest
         Dim lo_FcgiApp As VbFcgiLib.IFcgiApp
         Dim l_AppKey As String
         Dim l_ScriptName As String
         Dim l_StartedLoopAt As Double
         Dim l_LoopTime As Double
         Dim l_ProcessedRequest As Boolean   ' TRUE if at least one response was flushed
         Dim ii As Long
         
         Dim l_ContentLengthCgi As Boolean   ' Value of the passed CGI Content-Length parameter
         Dim l_ContentLengthStdin As Boolean ' Length of the data in the STDIN stream
         
         Dim l_ErrNumber As Long
         Dim l_ErrDesc As String
         Dim l_Erl As Long
            
10       On Error GoTo ErrorHandler

20       If s_Busy Then Exit Sub
30       s_Busy = True

40       mo_ResponseTimer.Enabled = False

50       apiOutputDebugString "In mo_ResponseTimer_Timer."

60       If mo_ActiveRequests Is Nothing Then GoTo Cleanup
70       If mo_ActiveRequests.Count = 0 Then GoTo Cleanup

80       l_StartedLoopAt = libRc5Factory.C.HPTimer
         
90       Do
100         For Each lo_Downstream In mo_ActiveRequests
110            If lo_Downstream.State = fcgireqstate_Error Then
                  ' Error occurred, send STDERR
                  apiOutputDebugString "Sending error downstream. Error #" & lo_Downstream.ErrorNumber & " " & lo_Downstream.ErrorDescription
                  
120               SendStdErrByRequestNoErr lo_Downstream
130               l_ProcessedRequest = True
                  
140            Else
                  ' No error, check for ready-state of request
150               If lo_Downstream.State <> fcgireqstate_Busy Then
160                  Select Case lo_Downstream.StdOut.State
                     Case fcgistdoutstate_Built, fcgistdoutstate_Building
                        ' STDOUT build has been finished, SO:
                        ' Flush the remaining buffer, send the end request,
                        ' and remove the object from the active collection
170                     apiOutputDebugString "Sending STDOUT for request #" & lo_Downstream.Id
               
180                     If fcgiFlushStdOut(mo_TcpServer, lo_Downstream.SocketNumber, lo_Downstream.Id, lo_Downstream.StdOut, (lo_Downstream.StdOut.State = fcgistdoutstate_Built)) Then
                           ' Flushed - send FCGI_END_REQUEST
190                        If lo_Downstream.StdOut.State = fcgistdoutstate_Built Then
200                           If Not lo_Downstream.StdOut.HasUnflushedContent Then
210                              fcgiSendEndRequest mo_TcpServer, lo_Downstream.SocketNumber, lo_Downstream.Id, 0, 0
220                           End If
230                        End If
240                     End If
                        
250                     l_ProcessedRequest = True
                        
260                     If lo_Downstream.StdOut.State = fcgistdoutstate_Built Then
270                        If Not lo_Downstream.StdOut.HasUnflushedContent Then
280                           RemoveRequestNoErr lo_Downstream
290                        End If
300                     End If
                        
310                  Case Else
                        ' STDOUT is not finished being built yet
320                     Select Case lo_Downstream.State
                        Case fcgireqstate_Initialized, fcgireqstate_Ready
330                        If lo_Downstream.Params.State = paramstate_Built Then
                              ' Parameters have all been received, check for STDIN completion
340                           On Error Resume Next
                              ' Trap error if no ContentLength sent
350                           l_ContentLengthCgi = 0
360                           l_ContentLengthCgi = lo_Downstream.Params.ValueByEnum(stdparam_ContentLength)
370                           On Error GoTo ErrorHandler
                              
380                           If l_ContentLengthCgi > 0 Then
390                              If lo_Downstream.Stdin.State = fcgistdinstate_Built Then
400                                 l_ContentLengthStdin = lo_Downstream.Stdin.ContentLength
410                                 On Error GoTo ErrorHandler
                                    
420                                 If l_ContentLengthStdin <> l_ContentLengthCgi Then
                                       ' Completed STDIN length does NOT match CGI ContentLength
430                                    Err.Raise fcgierr_ContentLengthMismatch, , "CGI content length (" & l_ContentLengthCgi & ") does not match STDIN content length (" & l_ContentLengthStdin & ")"
440                                 End If
450                              Else
460                                 Err.Raise fcgierr_NotReadyForResponse, , "Waiting on STDIN for data."
470                              End If
480                           End If
                              
490                           apiOutputDebugString "Building response for request ID #" & lo_Downstream.Id & ". Total requests in queue: " & mo_ActiveRequests.Count
                              
500                           lo_Downstream.State = fcgireqstate_Busy
                                                      
510                           l_ScriptName = lo_Downstream.Params.ValueByEnum(stdparam_ScriptName)
                              
520                           Select Case Left$(l_ScriptName, 1)
                              Case "\", "/"
530                              l_ScriptName = Mid$(l_ScriptName, 2)
540                           End Select
                              
550                           If InStr(1, l_ScriptName, "\") > 0 Then Err.Raise 70, , "Permission denied on script sub folders."
560                           If InStr(1, l_ScriptName, "/") > 0 Then Err.Raise 70, , "Permission denied on script sub folders."
                              
570                           l_AppKey = LCase$(l_ScriptName)
                              
580                           If Not envDebugMode Then
                                 ' Only allow access to apps with a .fcgi extension when not in debug mode.
590                              If Right$(l_AppKey, Len(".fcgi")) <> ".fcgi" Then
600                                 Err.Raise 70, , "Permission denied to access libraries with a .fcgi file extesion. Passed script name: " & l_ScriptName
610                              End If
620                           End If
                              
630                           If mo_FcgiApps.Exists(l_AppKey) Then
640                              Set lo_FcgiApp = mo_FcgiApps.Item(l_AppKey)
650                           Else
660                              apiOutputDebugString "Creating CFcgiApp from: " & pathBin & l_ScriptName
                                 
670                              Set lo_FcgiApp = libRc5Factory.RegFree.GetInstanceEx(pathBin & l_ScriptName, "CFcgiApp")
680                              mo_FcgiApps.Add lo_FcgiApp, l_AppKey
690                           End If
                              
700                           Set lo_FcgiRequest = New VbFcgiLib.CFcgiRequest
710                           Set lo_FcgiRequest.FcgiDownstream = lo_Downstream
                              
720                           Set mo_ActiveDownstream = lo_Downstream.FcgiResponse
                              
730                           lo_FcgiApp.ProcessRequest lo_FcgiRequest, mo_ActiveDownstream
740                        End If
750                     End Select
760                  End Select
770               End If
780            End If
               
               ' Check to see if we've surpassed the response loop threshold
790            l_LoopTime = libRc5Factory.C.HPTimer - l_StartedLoopAt
               
800            Select Case l_LoopTime
               Case Is < 0, Is > mo_ResponseTimer.Interval * 2
                  ' Short-circuit subsequent responses if time for
                  ' previous response building/sending has gone over our threshold
810               apiOutputDebugString "Response timer exceeded allotted time. Will come resume later."
                  
820               Exit For
830            End Select
               
840            apiOutputDebugString "In response loop."
               
850         Next lo_Downstream
            
            ' Keep looping if we have any active requests and we have exceed our response loop threshold
860      Loop While l_ProcessedRequest And (l_LoopTime < mo_ResponseTimer.Interval * 2) And (mo_ActiveRequests.Count > 0)
         
870      l_ErrNumber = 0
880      l_ErrDesc = ""
         
Cleanup:
890      On Error Resume Next

         ' Check for error
900      If l_ErrNumber <> 0 Then
            ' Handle error
            ' If the error is fcgierr_ResponseBuilderNotReady, then we'll try to generate the response again later
            ' Otherwise, we'll kill the request
            
910         If l_ErrNumber = fcgierr_NotReadyForResponse Then
               ' Conditions were not met for the response to be built, try again later
920            If lo_Downstream Is Nothing Then
                  ' Sanity check - shouldn't get here.
930               apiOutputDebugString "Sanity Check in mo_ResponseTimer_Timer. Request Object = Nothing after handling fcgierr_ResponseBuilderNotReady."
                  
940               Debug.Assert False
950            Else
                  ' Return the request to the Ready state
960               lo_Downstream.State = fcgireqstate_Ready
970            End If
               
980         Else
990            apiOutputDebugString "Error in response timer event, line #" & l_Erl & ": " & l_ErrNumber & " " & l_ErrDesc
            
1000           If Not lo_Downstream Is Nothing Then
1010              lo_Downstream.SetError l_ErrNumber, l_ErrDesc
                  
1020              SendStdErrByRequestNoErr lo_Downstream
1030           End If

1040           RemoveRequestNoErr lo_Downstream
1050        End If
1060     End If

1070     s_Busy = False
         
         ' Restart timer if we have any remaining active responses
1080     If Not mo_ActiveRequests Is Nothing Then
1090        If mo_ActiveRequests.Count > 0 Then
1100           mo_ResponseTimer.Enabled = True
1110        End If
1120     End If
         
1130     Exit Sub

ErrorHandler:
         ' Record error information, and resume at cleanup
1140     l_Erl = Erl
1150     l_ErrNumber = Err.Number
1160     l_ErrDesc = Err.Description
         
1170     Resume Cleanup
End Sub

Private Sub SendStdErrByInfoNoErr(ByVal p_SocketNumber As Long, ByVal p_RequestId As Integer, ByVal p_ErrorNumber As Long, ByVal p_ErrorDescription As String)
   Dim l_ProtocolStatus As Byte
   
   On Error Resume Next
   
   If p_ErrorNumber = 0 Then p_ErrorNumber = fcgierr_UnexpectedError
   If p_ErrorDescription = "" Then p_ErrorDescription = "Unexpected Error."
   
   Select Case p_ErrorNumber
   Case fcgierr_Overloaded
      l_ProtocolStatus = FCGI_OVERLOADED
   Case fcgierr_UnsupportedRole
      l_ProtocolStatus = FCGI_UNKNOWN_ROLE
   End Select
   
   If fcgiSendStdErr(mo_TcpServer, p_SocketNumber, p_RequestId, p_ErrorNumber, p_ErrorDescription) Then
      fcgiSendEndRequest mo_TcpServer, p_SocketNumber, p_RequestId, p_ErrorNumber, l_ProtocolStatus
   End If
End Sub

Private Sub SendStdErrByRequestNoErr(po_Request As CFcgiDownstream)
   On Error Resume Next
   
   If po_Request Is Nothing Then Debug.Assert False: Exit Sub
   
   apiOutputDebugString "**ERROR** " & po_Request.ErrorNumber & ", " & po_Request.ErrorDescription
   
   SendStdErrByInfoNoErr po_Request.SocketNumber, po_Request.Id, po_Request.ErrorNumber, po_Request.ErrorDescription
   
   RemoveRequestNoErr po_Request
End Sub

Private Sub RemoveRequestNoErr(po_Request As CFcgiDownstream)
   ' Remove a request from the active queue if it exists, without raising any errors
   ' Making it safe for use in an error handler
   
   Dim l_RequestKey As String
   
   On Error Resume Next
   
   If po_Request Is Nothing Then Exit Sub
   If mo_ActiveRequests Is Nothing Then Exit Sub
   If mo_ActiveRequests.Count = 0 Then Exit Sub
   
   l_RequestKey = RequestKey(po_Request.Id)
   
   If mo_ActiveRequests.Exists(l_RequestKey) Then
      mo_ActiveRequests.Remove l_RequestKey
   End If
End Sub

Private Sub mo_TcpServer_TCPAccepted(ByVal hSocket As Long)
   apiOutputDebugString "TCP connection accepted on socket #" & hSocket
End Sub
