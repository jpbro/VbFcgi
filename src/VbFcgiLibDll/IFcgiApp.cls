VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFcgiApp"
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
' This class defines the interface that all upstream FCGI applications must
' implement in order to be used by the VbFcgi framework.
'
' -------------------------------------------------------------------------------
' USAGE
' -------------------------------------------------------------------------------
'
' To create a class that can be used by the VbFcgi framework you must:
'
' 1) Create a new ActiveX Dll project.
'
' 2) Add a reference to VbFcgiLib to your project.
'
' 3) Rename the default "Class1" class to "CFcgiApp"
'
' 4) In the "General" code section for the "CFcgiApp" class, type:
'
'    Implements VbFcgiLib.IFcgiApp
'
' 5) Add the following code:
'
'    Private Sub IFcgiApp_ProcessRequest(po_Request As VbFcgiLib.CFcgiRequest, po_Response As VbFcgiLib.CFcgiResponse)
'
' 6) Code your web application within the IFcgiApp_ProcessRequest sub. See comments
'    in the ProcessRequest sub (below) for more details.

Public Sub ProcessRequest(po_Request As VbFcgiLib.CFcgiRequest, po_Response As VbFcgiLib.CFcgiResponse)
   ' This sub is called when an FCGI request is received by the CFgiServer class.
   
   ' The po_Request object holds information received from
   ' the browser (Http) and web server (Fcgi) that can be queried to build your dynamic web pages.
   ' For more information, see the VbFcgiLib.CFcgiRequest class.
   
   ' The po_Response object should be used by your FCGI application
   ' to pass data downstream to the web server/browser.
   ' For more information, see the VbFcgiLib.CFcgiResponse class.
End Sub

