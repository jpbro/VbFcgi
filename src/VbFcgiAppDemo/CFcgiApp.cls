VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFcgiApp"
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

Implements VbFcgiLib.IFcgiApp

' Every VbFcgiApp.dll must inclde a CFcgiApp class as a minimum.
' The CFcgiApp class must include the "Implements VbFcgiLib.IFcgiApp" directive in the General section.
' The CFcgiApp class must include all members of the IFcgiApp interface. They are:
'        IFcgiApp_BuildResponse and IFcgiApp_UpstreamObject
' After instatiation, your CFcgiApp class will be passed a VbFcgiLib.CFcgiUpstream object to IFcgiApp_UpstreamObject
'     The Upstream object should be stored and its methods should be called for sending responses and raising errors.
'     The Upstream object class methods are:
'        WriteBytes, Error, and Finished

' This is a very small sample class for building responses to FCGI requests.
' It simply spits back an HTML page with the following:
' Unicode/UTF-8 test, generated date/time, list of received parameters, approx. time taken to build response
' In reality you would parse out the QUERY_STRING and HTTP_COOKIES parameters where available
' And respond accordingly, but this is beyond the scope of this demo.

Private mo_Resp As vbRichClient5.cStringBuilder

Private mo_FcgiParams As VbFcgiLib.CFcgiParams
Private mo_FcgiStdin As VbFcgiLib.CFcgiStdIn

Private Sub IFcgiApp_ProcessRequest(po_Request As VbFcgiLib.CFcgiRequest, po_Response As VbFcgiLib.CFcgiResponse)
   Dim l_StartedBuildAt As Double
   Dim la_Header() As Byte
   Dim la_Content() As Byte
   Dim ii As Long
   
   On Error GoTo ErrorHandler

   ' Make sure that CGI parameters are complete built and the Upstream FCGI object has been set
   ' otherwise raise fcgierr_NotReadyForResponse
   ' None of these error states should ever happen
   If po_Request.Fcgi.Params.State <> paramstate_Built Then Err.Raise fcgierr_NotReadyForResponse, , "FCGI Parameters incomplete."
   
   l_StartedBuildAt = libRc5Factory.C.HPTimer

   Set mo_FcgiParams = po_Request.Fcgi.Params
   Set mo_FcgiStdin = po_Request.Fcgi.Stdin
   
   ' Build the HTML portion of the HTTP response
   With mo_Resp
      If .Length > 0 Then .Clear
      
      .AppendNL "<!DOCTYPE html><html><head><title>VbFCGI Demo App</title></head><body>"
      .AppendNL "<h1>"
      .AppendNL "Welcome to the VbFcgi Demo App!"
      .AppendNL "</h1>"
      
      .AppendNL "<p><b><a href='https://www.github.com/jpbro/VbFcgi'>Learn more about VbFcgi on GitHub.</a></b></p>"
      
      .Append "<p>The current date &amp; time on the server is: "
      .Append CStr(Now)
      .AppendNL "</p>"

      .Append "<p>VbFcgi is "
      .Append ChrW$(&HAAA&) & ChrW$(&HE01&) & ChrW$(&H671&) & ChrW$(&H188&) & ChrW$(&H47B&) & ChrW$(&H257&) & ChrW$(&HFEC9&)
      .AppendNL " capable via UTF-8!</p>"
      
      ' Append parameter key/value table
      .AppendNL "<h2>Parameters Received from Upstream Webserver:</h2>"
      
      .Append "<table>"
      For ii = 1 To mo_FcgiParams.Count - 1
         .Append "<tr><td>"
         .Append mo_FcgiParams.KeyByIndex(ii)
         .Append "</td><td>"
         .Append mo_FcgiParams.ValueByIndex(ii)
         .AppendNL "</td></tr>"
      Next ii
      .AppendNL "</table>"

      ' Append STDIN info
      .AppendNL "<p>FCGI STDIN Content Length: "
      .AppendNL mo_FcgiStdin.ContentLength
      .AppendNL "</p>"
      
      ' Append build time in fractional seconds
      .Append "<p>Response Build Time: "
      .Append Format$(libRc5Factory.C.HPTimer - l_StartedBuildAt, "0.00000")
      .AppendNL " seconds</p>"

      .Append "</body></html>"

      ' Get response content in UTF8, and clear the object
      la_Content = .ToUTF8
      
      .Clear   ' Release memory
   End With

   ' Build the HTTP header portion of the response
   la_Header = StrConv("Content-Type: text/html; charset=utf-8" & vbNewLine & "Content-Length: " & (UBound(la_Content) + 1) & vbNewLine & vbNewLine, vbFromUnicode)

   ' Signal that bytes are reading for STDOUT
   po_Response.WriteBytes la_Header

   po_Response.WriteBytes la_Content

   ' Signal that we are finished with the request/response
   po_Response.Finished

   Exit Sub

ErrorHandler:
   If Not mo_Resp Is Nothing Then
      mo_Resp.Clear
   End If
   
   If Err.Number = fcgierr_NotReadyForResponse Then
      Err.Raise Err.Number, , Err.Description
   Else
      ' Raise an error event
      po_Response.Error Err.Number, Err.Description & ", Line # " & Erl
   End If
End Sub

Private Sub Class_Initialize()
   Set mo_Resp = libRc5Factory.C.StringBuilder
End Sub

Private Sub Class_Terminate()
   Set mo_Resp = Nothing
End Sub

