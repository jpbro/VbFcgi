VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private WithEvents mo_Template As VbFcgiLib.CBuilderTemplate
Attribute mo_Template.VB_VarHelpID = -1

Public Sub TestTemplateInfiniteLoopProtection()
   Set mo_Template = New CBuilderTemplate
   
   mo_Template.ParseFile "This is a [[TITLE]] test."
   mo_Template.Finish contentencoding_UTF16_LE
   Debug.Print mo_Template.Content
End Sub

Private Sub mo_Template_FoundTag(ByVal p_Tag As String, po_Replacement As VbFcgiLib.CWebStringTemplate, ByRef p_DoNotReplace As Boolean)
   Select Case p_Tag
   Case "TITLE"
      po_Replacement.SkipEncodeEntities = True
      po_Replacement.OpenTags "p"
      po_Replacement.Append "HELLO [[ABC]]"
      po_Replacement.CloseAllOpenedTags
      
   Case "ABC"
      po_Replacement = "XXX [[TITLE]] XXX"
   End Select
End Sub
