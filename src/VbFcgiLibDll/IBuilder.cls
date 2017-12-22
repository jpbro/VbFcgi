VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBuilder"
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
' This class defines the interface that all VbFcgi CBuilder* classes implement.
' Classes that implement the IBuilder interface are designed to help you more
' easily build web content (including HTTP headers) using VB6.
'
' As of December 21, 2017, the following VbFcgiLib classes implement the
' IBuilder interface:
'
' CBuilderFile - helper class for sending locally stored files downstream.
'                (for example, PDFs, images, etc...)
'
' CBuilderHtml - helper class for building HTML files completely in memory.
'
' CBuilderTemplate - helper class for reading template files (usually HTML)
'                    from the local file system and detecting & replacing special
'                    "tags" to generate dynamic page data for sending downstream.
'
' -------------------------------------------------------------------------------
' USAGE
' -------------------------------------------------------------------------------
'
' In general you will not implement your own IBuilder classes, so usage is
' limited to interaction with the packaged classes that implement the
' IBuilder interface. See each IBuilder implementing class for usage specific
' to those classes.

Public Property Get HttpHeader() As VbFcgiLib.CHttpHeader
   ' Returns a IHttpHeaderFields object
End Property

Public Property Get IsFinished() As Boolean
   ' Returns TRUE if the Finish method has been called.
   ' Otherwise returns FALSE
End Property

Public Sub Finish(Optional ByVal p_ContentEncoding As VbFcgiLib.e_ContentEncoding = contentencoding_RecommendedDefault)
   ' Call this method to close the content buffer and convert it to the passed encoding.
   ' This method must be called before Length/Content properties are called.
End Sub

Public Property Get Length() As Long
   ' Returns the length of the content buffer in bytes (in the encoding passed to the Finish method)
   ' This property cannot be called before the Finish method has been called
End Property

Public Property Get Content() As Byte()
   ' Returns a byte array of all content in the buffer in the encoding passed to the Finish method
   ' This property cannot be called before the Finish method has been called
End Property

Public Sub Reset()
   ' Clear all variables and objects and/or reset them to sensible defaults
End Sub
