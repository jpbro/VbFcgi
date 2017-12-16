VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHttpQueryParamValues"
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

Private mo_Values As vbRichClient5.cArrayList

Public Property Get Count() As Long
   Count = mo_Values.Count
End Property

Public Sub Add(ByVal p_Value As String)
   mo_Values.Add p_Value
End Sub

Public Property Get ValueByIndex(ByVal p_ZeroBasedIndex As Long) As String
   ValueByIndex = mo_Values.Item(p_ZeroBasedIndex)
End Property

Public Property Let ValueByIndex(ByVal p_ZeroBasedIndex As Long, ByVal p_Value As String)
   mo_Values.Item(p_ZeroBasedIndex) = p_Value
End Property

Private Sub Class_Initialize()
   Set mo_Values = libRc5Factory.C.ArrayList(vbString)
End Sub

Public Property Get IsValueTrueByIndex(ByVal p_ZeroBasedIndex As Long) As Boolean
   Dim l_Value As String
   Dim l_ValueLong As Long
   
   l_Value = LCase$(Me.ValueByIndex(p_ZeroBasedIndex))
   
   Select Case l_Value
   Case "1", "-1", "on", "yes", "true", "y"
      IsValueTrueByIndex = True
   Case Else
      If IsNumeric(l_Value) Then
         l_ValueLong = CLng(l_Value)
      End If
      
      IsValueTrueByIndex = (l_ValueLong <> 0)
   End Select
End Property

