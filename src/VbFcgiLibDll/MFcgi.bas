Attribute VB_Name = "MFcgi"
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

' This module is for FCGI related constants, types, and methods.

' TODO: Use enums instead of constants to take advantage of Intellisense

' Record Types
Public Const FCGI_BEGIN_REQUEST As Byte = 1 ' An FCGI request is beginning
Public Const FCGI_ABORT_REQUEST As Byte = 2 ' Web server wants the request to be aborted
Public Const FCGI_END_REQUEST As Byte = 3 ' An FCGI request is ending
Public Const FCGI_PARAMS As Byte = 4 ' Series of key-value pairs
Public Const FCGI_STDIN As Byte = 5 ' STDIN buffer (from requests)
Public Const FCGI_STDOUT As Byte = 6   ' STDOUT buffer (for responses)
Public Const FCGI_STDERR As Byte = 7   ' STDERR buffer (for error responses)

' Roles
Public Const FCGI_RESPONDER  As Byte = 1  ' Standard type CGI role
Public Const FCGI_AUTHORIZER As Byte = 2  ' UNSUPPORTED
Public Const FCGI_FILTER As Byte = 3   ' UNSUPPORTED

' Protocol Statuses for FCGI_END_REQUEST records
Public Const FCGI_REQUEST_COMPLETE As Byte = 0
Public Const FCGI_CANT_MPX_CONN As Byte = 1
Public Const FCGI_OVERLOADED As Byte = 2
Public Const FCGI_UNKNOWN_ROLE As Byte = 3

Public Type FCGX_RECORD
   Version As Byte
   RecordType As Byte
   RequestId As Integer   ' Must reverse bytes!
   ContentLength As Integer   ' Must reverse bytes!
   PaddingLength As Byte
   Reserved As Byte
End Type

Public Type FCGX_BEGIN_REQUEST_BODY
   Role As Integer   ' Must reverse bytes!
   Flags As Byte
   Reserved(0 To 4) As Byte
End Type

Public Type FCGX_END_REQUEST_BODY
   ApplicationStatus As Long  ' Must reverse bytes!
   ProtocolStatus As Byte
   Reserved(0 To 2) As Byte
End Type

Public Function fcgiFlushStdOut(po_TcpServer As vbRichClient5.cTCPServer, ByVal p_Socket As Long, ByVal p_RequestId As Integer, po_StdOut As CFcgiStdOut, ByVal p_CloseStream As Boolean) As Boolean
   ' Send any unsent bytes stored in the passed request's STDOUT buffer to the web server
   Dim la_Record() As Byte
   Dim la_Content() As Byte
   Dim l_Len As Integer
   Dim l_Padding As Byte
   Dim l_StartedAt As Double
   
   apiOutputDebugString "In fcgiFlushStdOut."
  
   If po_TcpServer.ConnectionCount = 0 Then
      apiOutputDebugString "No connections found in fcgiFlushStdOut. Short circuiting."
      Exit Function
   End If
   
   ReDim la_Record(7)
   
   la_Record(0) = 1  ' Version
   la_Record(1) = FCGI_STDOUT

   p_RequestId = apiNtohs(p_RequestId)
   apiCopyMemory la_Record(2), p_RequestId, 2
   
   If Not po_StdOut Is Nothing Then
      If po_StdOut.HasUnflushedContent Then
         l_StartedAt = libRc5Factory.C.HPTimer
         
         Do
            la_Content = po_StdOut.NextContentChunk
            l_Len = arraySize(la_Content)
            
            If po_StdOut.HasUnflushedContent Then
               l_Padding = 0
            Else
               l_Padding = 8 - (l_Len Mod 8)
            End If
            
            apiOutputDebugString "Sending STDOUT chunk. Len: " & l_Len
            
            If l_Len <= 0 Then
               Debug.Assert False
               ' Empty array - don't send.
               ' Should never actually get here, just a sanity check
               Exit Do
               
            Else
               l_Len = apiNtohs(l_Len)
               apiCopyMemory la_Record(4), l_Len, 2
               la_Record(6) = l_Padding
               
               If po_TcpServer.SendData(p_Socket, VarPtr(la_Record(0)), arraySize(la_Record)) Then
                  apiOutputDebugString "Flushed STDOUT record header."
                  
                  If po_TcpServer.SendData(p_Socket, VarPtr(la_Content(0)), arraySize(la_Content)) Then
                     apiOutputDebugString "Flushed STDOUT content. Length: " & arraySize(la_Content)
                     
                     If l_Padding <> 0 Then
                        ReDim la_Content(l_Padding - 1)
                        If po_TcpServer.SendData(p_Socket, VarPtr(la_Content(0)), arraySize(la_Content)) Then
                           apiOutputDebugString "Flushed padding. Length: " & l_Padding
                           fcgiFlushStdOut = True
                        Else
                           apiOutputDebugString "Could not flush padding. Length: " & l_Padding
                           fcgiFlushStdOut = False
                           Exit Do
                        End If
                        
                     Else
                        apiOutputDebugString "No padding to flush (not an error)."
                        fcgiFlushStdOut = True
                     End If
                     
                  Else
                     apiOutputDebugString "Could not flush content chunk. Length: " & arraySize(UBound(la_Content))
                     fcgiFlushStdOut = False
                     Exit Do
                  End If
               Else
                  apiOutputDebugString "Could not flush STDOUT record header. Length: " & arraySize(UBound(la_Record))
                  fcgiFlushStdOut = False
                  Exit Do
               End If
               
               Erase la_Content
            End If
            
            Select Case libRc5Factory.C.HPTimer - l_StartedAt
            Case Is < 0, Is > 0.5
               apiOutputDebugString "fcgiFlushStdOut looped too long, leaving..."
               
               Exit Do
            End Select
            
            apiOutputDebugString "In fcgiFlushStdOut loop."
            
         Loop While po_StdOut.HasUnflushedContent And fcgiFlushStdOut
         
         If Not fcgiFlushStdOut Then
            apiOutputDebugString "Failed to flush STDOUT."
            p_CloseStream = False
         End If
         
      Else
         apiOutputDebugString "STDOUT has no unflushed content."
      End If
   End If
   
   If p_CloseStream Then
      apiOutputDebugString "Closing stream in fcgiFlushStdOut."
      
      ' Send closing STDOUT/STDERR record
      la_Record(4) = 0  ' ZERO LENGTH
      la_Record(5) = 0  ' ZERO LENGTH
      la_Record(6) = 0  ' ZERO PADDING
      
      If po_TcpServer.SendData(p_Socket, VarPtr(la_Record(0)), UBound(la_Record) + 1) Then
         ' Success
         apiOutputDebugString "Sent closing STDOUT/STDERR record."
      Else
         ' Failure
         apiOutputDebugString "Could not send closing STDOUT record."
      End If
      
      fcgiFlushStdOut = True  ' We flushed something so return true
   End If

   apiOutputDebugString "Leaving fcgiFlushStdOut."
End Function

Public Function fcgiSendStdErr(po_TcpServer As vbRichClient5.cTCPServer, ByVal p_Socket As Long, ByVal p_RequestId As Integer, ByVal p_ErrorNumber As Long, ByVal p_ErrorDescription As String) As Boolean
   ' Create and send STDERR record with error message
   Dim l_Padding As Byte
   Dim la_Content() As Byte
   Dim la_Record(7) As Byte
   Dim l_RequestId As Integer
   Dim l_ContentLen As Integer
   
   On Error GoTo ErrorHandler
   
   apiOutputDebugString "In fcgiSendStdErr."
   
   If po_TcpServer.ConnectionCount = 0 Then
      apiOutputDebugString "No connections found in fcgiSendStdErr. Short-circuiting."
      Exit Function
   End If
   
   If p_Socket = 0 Then
      apiOutputDebugString "No socket # available in fcgiSendStdErr. Short-circuiting."
      Exit Function
   End If
   
   apiOutputDebugString "Sending error #" & p_ErrorNumber & " '" & p_ErrorDescription & "' in fcgiSendStdErr."
   
   la_Content = libCrypt.VBStringToUTF8(p_ErrorDescription)
   
   l_Padding = 8 - ((UBound(la_Content) + 1) Mod 8)
   
   ' Reverse integer values
   l_RequestId = apiNtohs(p_RequestId)
   l_ContentLen = apiNtohs(arraySize(la_Content))
   
   ' Build STDERR record
   la_Record(0) = 1  ' Version
   la_Record(1) = FCGI_STDERR
   apiCopyMemory la_Record(2), l_RequestId, 2
   apiCopyMemory la_Record(4), l_ContentLen, 2
   la_Record(6) = 0
   la_Record(7) = l_Padding
   
   If po_TcpServer.SendData(p_Socket, VarPtr(la_Record(0)), UBound(la_Record) + 1) Then
      If po_TcpServer.SendData(p_Socket, VarPtr(la_Content(0)), UBound(la_Content) + 1) Then
         If l_Padding > 0 Then
            ReDim la_Content(l_Padding - 1)
            If Not po_TcpServer.SendData(p_Socket, VarPtr(la_Content(0)), l_Padding) Then
               apiOutputDebugString "Failed to send payload #3 (record padding) in fcgiSendStdErr."
            End If
         End If
      
         ' Send empty STDERR record to close it
         la_Record(4) = 0
         la_Record(5) = 0
         la_Record(6) = 0
         la_Record(7) = 0
      
         If Not po_TcpServer.SendData(p_Socket, VarPtr(la_Record(0)), UBound(la_Record) + 1) Then
            apiOutputDebugString "Failed to send payload #4 (record close) in fcgiSendStdErr."
         End If
         
      Else
         apiOutputDebugString "Failed to send payload #2 (record content) in fcgiSendStdErr."
      End If
      
   Else
      apiOutputDebugString "Failed to send payload #1 (record header) in fcgiSendStdErr."
   End If

   apiOutputDebugString "Leaving fcgiSendStdErr."

   Exit Function
   
ErrorHandler:
   apiOutputDebugString "*** Error #" & Err.Number & " " & Err.Description & " in fcgiSendStdErr!"
End Function

Public Sub fcgiSendEndRequest(po_TcpServer As vbRichClient5.cTCPServer, ByVal p_Socket As Long, ByVal p_RequestId As Integer, ByVal p_ApplicationStatus As Long, ByVal p_ProtocolStatus As Byte)
   Dim l_Id As Integer
   Dim l_Len As Integer
   Dim la_Record() As Byte
   
   apiOutputDebugString "In fcgiSendEndRequest."
   
   If po_TcpServer.ConnectionCount = 0 Then
      apiOutputDebugString "No connections found in fcgiSendEndRequest. Short circuiting."
      Exit Sub
   End If
   ReDim la_Record(15)
   
   la_Record(0) = 1  ' Version
   la_Record(1) = FCGI_END_REQUEST

   p_RequestId = apiNtohs(p_RequestId)
   l_Len = apiNtohs(8)
   
   apiCopyMemory la_Record(2), p_RequestId, 2
   apiCopyMemory la_Record(4), l_Len, 2
   
   If p_ApplicationStatus <> 0 Then
      p_ApplicationStatus = apiNtohl(p_ApplicationStatus)
      apiCopyMemory la_Record(8), p_ApplicationStatus, 4
   End If
   
   If p_ProtocolStatus <> 0 Then
      la_Record(12) = p_ProtocolStatus
   End If
   
   If po_TcpServer.SendData(p_Socket, VarPtr(la_Record(0)), UBound(la_Record) + 1) Then
      apiOutputDebugString "Sent End request."
   Else
      apiOutputDebugString "Could not send End request."
   End If

   apiOutputDebugString "Leaving fcgiSendEndRequest."
End Sub


