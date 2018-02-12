Attribute VB_Name = "MArrays"
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

Public Function arrayIsDimmed(p_Array As Variant) As Boolean
   Dim l_Lbound As Long
   
   If Not VarType(p_Array) And vbArray Then Err.Raise 5, , "Array required."
   
   On Error Resume Next
   Err.Clear
   l_Lbound = LBound(p_Array)
   arrayIsDimmed = (Err.Number = 0)
   Err.Clear
   
   If arrayIsDimmed Then
      If UBound(p_Array) < l_Lbound Then
         arrayIsDimmed = False
      End If
   End If
   On Error GoTo 0
End Function

Public Function arraySize(p_Array As Variant, Optional p_RaiseErrorIfUndimmed As Boolean)
   If arrayIsDimmed(p_Array) Then
      arraySize = UBound(p_Array) - LBound(p_Array) + 1
   Else
      If p_RaiseErrorIfUndimmed Then
         Err.Raise 9, , "Array not dimensioned in arraySize method."
      Else
         arraySize = -1
      End If
   End If
End Function
