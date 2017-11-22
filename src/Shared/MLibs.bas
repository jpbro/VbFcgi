Attribute VB_Name = "MLibs"
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

' DirectCOM Stuff
Private Declare Function GetInstanceEx Lib "DirectCom" (StrPtr_FName As Long, StrPtr_ClassName As Long, Optional ByVal UseAlteredSearchPath As Boolean = True) As Object
Private Declare Function GetInstanceOld Lib "DirectCom" Alias "GETINSTANCE" (FName As String, ClassName As String) As Object
Private Declare Function GETINSTANCELASTERROR Lib "DirectCom" () As String
'used, to preload DirectCOM.dll from a given Path, before we try our calls
Private Declare Function LoadLibraryW Lib "kernel32.dll" (ByVal LibFilePath As Long) As Long

'The new GetInstance-Wrapper-Proc, which is using the new DirectCOM.dll (March 2009 and newer)
'with the new Unicode-capable GetInstanceEx-Call (which now supports the AlteredSearchPath-Flag as well) -
'If you omit that optional param or set it to True, then LoadLibraryExW is used with the appropriate
'Flag. If the Param was set to False, then the behaviour is the same as with the former
'DirectCOM.dll-GETINSTANCE-Call - only that LoadLibraryW is used instead of LoadLibraryA.
'This routine also tries a fallback to the former DirectCOM.dll-GETINSTANCE-Call, in case
'you are using it against an older version of this small regfree-helper-lib.
Private Function GETINSTANCE(DllFileName As String, ClassName As String, Optional ByVal UseAlteredSearchPath As Boolean = True) As Object
   On Error Resume Next

   Set GETINSTANCE = GetInstanceEx(StrPtr(DllFileName), StrPtr(ClassName), UseAlteredSearchPath)
   If Err.Number = 453 Then      'GetInstanceEx not available, probably an older DirectCOM.dll...
      Err.Clear
      Set GETINSTANCE = GetInstanceOld(DllFileName, ClassName)      'so let's try the older GETINSTANCE-call
   End If
   If Err Then
      Dim Error As String
      Error = Err.Description
      On Error GoTo 0
      Err.Raise vbObjectError, , Error
   Else
      If GETINSTANCE Is Nothing Then
         On Error GoTo 0
         Err.Raise vbObjectError, , GETINSTANCELASTERROR()
      End If
   End If
End Function

Public Function libRc5Factory() As vbRichClient5.cFactory
   Static so_Lib As vbRichClient5.cFactory
   
   If so_Lib Is Nothing Then
      If envRunningInIde Then
         Set so_Lib = New vbRichClient5.cFactory
      Else
         Set so_Lib = GETINSTANCE(pathBin & "vbRichClient5.dll", "cFactory", True)
      End If
   End If
   
   Set libRc5Factory = so_Lib
End Function

Public Function libCrypt() As vbRichClient5.cCrypt
   Static so_Lib As vbRichClient5.cCrypt
   
   If so_Lib Is Nothing Then
      If envRunningInIde Then
         Set so_Lib = New vbRichClient5.cCrypt
      Else
         Set so_Lib = libRc5Factory.C.Crypt
      End If
   End If
   
   Set libCrypt = so_Lib
End Function

