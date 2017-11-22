Attribute VB_Name = "MEnv"
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

Public Function envDebugMode() As Boolean
   Static s_DebugMode As Boolean
   Static s_LastChecked As Double
   
   If libRc5Factory.C.HPTimer - s_LastChecked > 30 Then
      ' time has passed to warrant checking if debug mode is enabled
      s_DebugMode = libRc5Factory.C.FSO.FileExists(pathBin & "VbFcgi.debug")
      s_LastChecked = libRc5Factory.C.HPTimer
   End If
End Function

Public Function envRunningInIde() As Boolean
   Static s_Checked As Boolean
   Static s_InIde As Boolean
   
   If Not s_Checked Then
      s_Checked = True
      
      On Error Resume Next
      Debug.Print 1 / 0
      s_InIde = Err.Number
      On Error GoTo 0
   End If
   
   envRunningInIde = s_InIde
End Function

