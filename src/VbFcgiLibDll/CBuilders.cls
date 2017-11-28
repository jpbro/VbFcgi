VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBuilders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum e_ContentEncoding
   contentencoding_Unknown
   
   contentencoding_UTF8
   contentencoding_UTF16_LE
End Enum

Private mo_Html As VbFcgiLib.CBuilderHtml

Public Function Html(Optional ByVal p_CreateNew As Boolean) As VbFcgiLib.CBuilderHtml
   If (mo_Html Is Nothing) Or p_CreateNew Then Set mo_Html = New VbFcgiLib.CBuilderHtml
   
   Set Html = mo_Html
End Function
