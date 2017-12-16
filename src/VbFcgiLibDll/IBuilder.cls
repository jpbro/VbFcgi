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
