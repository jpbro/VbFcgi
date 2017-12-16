Attribute VB_Name = "MArrays"
Option Explicit

Public Function arrayIsDimmed(p_Array As Variant) As Boolean
   Dim l_Lbound As Long
   
   If Not VarType(p_Array) And vbArray Then Err.Raise 5, , "Array required."
   
   On Error Resume Next
   Err.Clear
   l_Lbound = LBound(p_Array)
   arrayIsDimmed = (Err.Number = 0)
   Err.Clear
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
