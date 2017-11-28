VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFcgiResponse"
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

Public Event SendBytes(ByVal p_RequestId As Long, pa_Bytes() As Byte)
Public Event SendError(ByVal p_RequestId As Long, ByVal p_ErrorNumber As Long, ByVal p_ErrorDescription As String)
Public Event SendFinished(ByVal p_RequestId As Long)

Private mo_Builders As VbFcgiLib.CBuilders

Private m_RequestId As Long
Private m_Finished As Boolean

Friend Property Let RequestId(ByVal p_RequestId As Long)
   m_RequestId = p_RequestId
End Property

Public Property Get Builders() As VbFcgiLib.CBuilders
   If mo_Builders Is Nothing Then Set mo_Builders = New VbFcgiLib.CBuilders
   
   Set Builders = mo_Builders
End Property

Public Sub WriteBytes(pa_Bytes() As Byte)
   ' Call this method from your FCGI App to write bytes back to the web server.
      
   If m_Finished Then Err.Raise 5, , "FCGI application finished."
   
   RaiseEvent SendBytes(m_RequestId, pa_Bytes)
End Sub

Public Sub Error(ByVal p_ErrorNumber As Long, ByVal p_ErrorDescription As String)
   ' Call this method from your FCGI app to send an error back to the web server
   
   If m_Finished Then Err.Raise 5, , "FCGI application finished."
   
   RaiseEvent SendError(m_RequestId, p_ErrorNumber, p_ErrorDescription)
   Me.Finished
End Sub

Public Sub Finished()
   ' Call this method when your FCGI app is finished responding to the web server.
   
   m_Finished = True
   
   RaiseEvent SendFinished(m_RequestId)
End Sub
