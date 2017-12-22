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

' -------------------------------------------------------------------------------
' PURPOSE
' -------------------------------------------------------------------------------
'
' This class holds references to CFcgi and CHttp classes (which themselves
' hold information received from the downstream browser and web server.
'
' A CFcgiRequest class is passed as a parameter to your FCGI application when
' IFcgiApp_ProcessRequest is called.
'
' -------------------------------------------------------------------------------
' USAGE
' -------------------------------------------------------------------------------
'
' This class is useful for getting information that was sent upstream by a
' the web browser and web server.
' For things like the request method (GET, PUT, etc...), the query parameters,
' HTTP request header fields, cookies, etc...

Private mo_Http As VbFcgiLib.CHttp
Private mo_Fcgi As VbFcgiLib.CFcgi
Private mo_Upstream As VbFcgiLib.CFcgiUpstream

Friend Property Set FcgiUpstream(po_Upstream As VbFcgiLib.CFcgiUpstream)
   Set mo_Upstream = po_Upstream
   Set mo_Http.FcgiUpstream = po_Upstream
   Set mo_Fcgi.FcgiUpstream = po_Upstream
End Property

Public Property Get Fcgi() As VbFcgiLib.CFcgi
   Set Fcgi = mo_Fcgi
End Property

Public Property Get Http() As VbFcgiLib.CHttp
   Set Http = mo_Http
End Property

Private Sub Class_Initialize()
   Set mo_Http = New VbFcgiLib.CHttp
   Set mo_Fcgi = New VbFcgiLib.CFcgi
End Sub
