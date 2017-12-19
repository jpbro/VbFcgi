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

Public Enum e_ContentEncoding
   contentencoding_RecommendedDefault = -1
   contentencoding_Unknown
   
   contentencoding_UTF8
   contentencoding_UTF16_LE
   contentencoding_USASCII
   contentencoding_ISO8859_1
End Enum

Public Enum e_BuilderType
   [_builder_NotInitialized] = 0
   
   builder_Html = 1 ' HTML string builder/helper
   builder_File
   
   ' Not supported yet:
   'builder_Json  ' JSON string builder/helper
   'builder_Xml   ' XML string builder/helper
   'builder_Raw   ' Raw socket to write whatever you want to
End Enum

Private Const mc_KeyIBuilder As String = "IBuilder"

Public Event Finished()

Private WithEvents mo_EventCollection As vbRichClient5.cEventCollection
Attribute mo_EventCollection.VB_VarHelpID = -1
Private mo_Builder As VbFcgiLib.IBuilder
Private m_BuilderType As VbFcgiLib.e_BuilderType

Public Function Builder(Optional ByVal p_InitializeBuilderType As VbFcgiLib.e_BuilderType) As VbFcgiLib.IBuilder
   If mo_Builder Is Nothing Then
      ' Initialize a new builder
   
      Select Case p_InitializeBuilderType
      Case builder_Html
         ' Requested an HTML builder/helper
         Set mo_Builder = New VbFcgiLib.CBuilderHtml
         
      Case builder_File
         ' Requested a file builder/helper
         Set mo_Builder = New VbFcgiLib.CBuilderFile
         
      Case Else
         Err.Raise 5, , "Unknown builder type: " & p_InitializeBuilderType
         
      End Select
   
      mo_EventCollection.Add mo_Builder, mc_KeyIBuilder
   
      m_BuilderType = p_InitializeBuilderType
      
   Else
      If p_InitializeBuilderType <> 0 Then
         If p_InitializeBuilderType <> m_BuilderType Then
            Err.Raise 5, , "Already initialized a different builder type!"
         End If
      End If
      
   End If
   
   Set Builder = mo_Builder
End Function

Private Sub Class_Initialize()
   Set mo_EventCollection = libRc5Factory.C.EventCollection
End Sub

Private Sub mo_EventCollection_EventRaised(p_Key As String, p_EventName As String, ByVal p_ParamCount As Long, p_Param1 As Variant, p_Param2 As Variant, p_Param3 As Variant, p_Param4 As Variant, p_Param5 As Variant, p_Param6 As Variant, p_Param7 As Variant, p_Param8 As Variant)
   apiOutputDebugString "In mo_EventCollection_EventRaised for " & p_Key & "." & p_EventName
   
   Select Case p_Key
   Case mc_KeyIBuilder
      Select Case LCase$(p_EventName)
      Case "finished"
         RaiseEvent Finished
      End Select
      
   Case Else
      Debug.Assert False
   End Select
End Sub
