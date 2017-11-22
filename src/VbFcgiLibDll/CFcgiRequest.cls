VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFcgiRequest"
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

' This class holds all of the information related to an FCGI request (including the response byte stream)

Public Enum e_FcgiRequestState
   fcgireqstate_Error = -1 ' There was an error processing the request
   
   fcgireqstate_Created = 0   ' Request object has just been created
   fcgireqstate_Initialized   ' Request has had initial state variable set
   fcgireqstate_Busy ' Request is off to the application for response processing
   fcgireqstate_Ready   ' Request is ready for flushing in whole or in part
End Enum

' Initial state information
Private m_State As e_FcgiRequestState  ' The current state of the request object
Private m_Id As Integer ' FCGI ID# of the request
Private m_SocketNumber As Long   ' Socket number for request/response communications with the web server

' Error information
Private m_ErrorNumber As Long ' This should be sent back through STDERR request body as the ApplicationStatus member
Private m_ErrorDescription As String

' FCGI Objects
Private mo_Params As CFcgiParams ' Collection of Key/Value pairs received from the web server
Private mo_StdIn As CFcgiStdIn   ' Byte data streamed from the web server
Private mo_StdOut As CFcgiStdOut ' Byte data to be streamed to the web server

Public Sub Initialize(ByVal p_RequestId As Integer, ByVal p_SocketNumber As Long)
   If Me.State <> fcgireqstate_Created Then Err.Raise 5, , "Can't re-initialize request."
   
   m_Id = p_RequestId
   m_SocketNumber = p_SocketNumber
   
   Me.State = fcgireqstate_Initialized
End Sub

Public Sub SetError(ByVal p_ErrorNumber As Long, ByVal p_ErrorDescription As String)
   m_ErrorNumber = p_ErrorNumber
   m_ErrorDescription = p_ErrorDescription
   
   Me.State = fcgireqstate_Error
End Sub

Public Property Get State() As e_FcgiRequestState
   State = m_State
End Property

Public Property Let State(ByVal p_State As e_FcgiRequestState)
   ' Check for invalid state changes
   If p_State = fcgireqstate_Created Then
      ' Illegal state change
      ' Can't return request to Created state - create a new object instead
      Err.Raise 5, , "Can't re-create request."
   End If
   
   If p_State <> fcgireqstate_Error Then
      If m_State = fcgireqstate_Error Then
         ' A request in an error state can not be removed from the error state
         ' Illegal state change
         Err.Raise 5, , "Can't remove request from error state."
      End If
   End If
   
   If p_State = fcgireqstate_Initialized Then
      ' A request can only be marked initialized if it is in the created or initialized states
      Select Case m_State
      Case fcgireqstate_Created, fcgireqstate_Initialized
         ' State change/status quo OK
      Case Else
         ' Illegal state change
         Err.Raise 5, , "Can't re-initialize request."
      End Select
   End If

   m_State = p_State
End Property

Public Property Get Params() As CFcgiParams
   Set Params = mo_Params
End Property

Public Property Get StdIn() As CFcgiStdIn
   Set StdIn = mo_StdIn
End Property

Public Property Get StdOut() As CFcgiStdOut
   Set StdOut = mo_StdOut
End Property

Public Property Get ErrorDescription() As String
   ErrorDescription = m_ErrorDescription
End Property

Public Property Get ErrorNumber() As Long
   ErrorNumber = m_ErrorNumber
End Property

Public Property Get Id() As Integer
   Id = m_Id
End Property

Public Property Get SocketNumber() As Long
   SocketNumber = m_SocketNumber
End Property

Private Sub Class_Initialize()
   Set mo_Params = New CFcgiParams
   Set mo_StdIn = New CFcgiStdIn
   Set mo_StdOut = New CFcgiStdOut
End Sub

