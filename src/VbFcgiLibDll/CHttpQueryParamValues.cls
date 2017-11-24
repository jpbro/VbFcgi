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

