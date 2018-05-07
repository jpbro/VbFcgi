VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFcgiStdIn"
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

' This class holds the STDIN data received from the web server
Public Enum e_FcgiStdinState
   fcgistdinstate_Error = -1  ' There was a problem with the STDIN buffer
   fcgistdinstate_Initialized   ' The buffer is ready for data to be appended
   fcgistdinstate_Building   ' Content has been and/or is being appended to the STDIN buffer
   fcgistdinstate_Built   ' We've receive all the STDIN data
End Enum

Private mo_StdIn As cStream   ' Stream of bytes received from the web server via STDIN record(s)

Private m_State As e_FcgiStdinState
Private m_FilePath As String   ' Path if the STDIN buffer is on the physical disk

Public Property Get State() As e_FcgiStdinState
   State = m_State
End Property

Public Property Let State(ByVal p_NewState As e_FcgiStdinState)
   If p_NewState = fcgistdinstate_Initialized Then Err.Raise 5, , "Can't re-initialize STDIN."
   If Me.State = fcgistdinstate_Built And p_NewState <> fcgistdinstate_Built Then Err.Raise 5, , "Can't remove STDIN from built state."
   If Me.State = fcgistdinstate_Error And p_NewState <> fcgistdinstate_Error Then Err.Raise 5, , "Can't remove STDIN from error state."

   m_State = p_NewState
End Property

Public Sub AppendContent(pa_Bytes() As Byte)
   Dim lo_DiskStream As vbRichClient5.cStream
   Dim l_ContentLength As Currency
   Dim la_Bytes() As Byte
   Dim l_AppendLength As Long
   
   Select Case Me.State
   Case fcgistdinstate_Error, fcgistdinstate_Built
      Err.Raise 5, , "Can not append to STDIN in the current state: " & Me.State
   End Select
      
   l_AppendLength = UBound(pa_Bytes) - LBound(pa_Bytes) + 1
   apiOutputDebugString "Appending content to STDIN. Size: " & l_AppendLength
   
   If l_AppendLength < 1 Then Exit Sub ' Nothing to append, Short-Circuit
   
   Me.State = fcgistdinstate_Building
   
   If m_FilePath = "" Then
      ' In memory stream - check if the existing + append length will exceed max in-memory sisze
      ' If so, move to temporary file on file system
      
      l_ContentLength = Me.ContentLength
      
      If l_ContentLength + l_AppendLength > gc_MaxStdinInMemorySize Then
         apiOutputDebugString "STDIN data exceeds maximum in-memory size of " & gc_MaxStdinInMemorySize & " bytes. Moving data to file system."
         apiOutputDebugString "Current ContentLength: " & l_ContentLength & ", Append Length: " & l_AppendLength
         
         ' Move to file system
         m_FilePath = libRc5Factory.C.FSO.GetTmpFileName("VBF")
         Set lo_DiskStream = libRc5Factory.C.FSO.CreateFileStream(m_FilePath, STRM_READWRITE + STRM_SHARE_EXCLUSIVE)
         
         apiOutputDebugString "STDIN file: " & m_FilePath
         
         If l_ContentLength > 0 Then
            apiOutputDebugString "Copying existing data from memory to disk."
            
            mo_StdIn.SetPosition 0  ' Go to beginning of stream
            mo_StdIn.ReadToByteArr la_Bytes
            lo_DiskStream.WriteFromByteArr la_Bytes
            Erase la_Bytes
         End If
         
         Set mo_StdIn = lo_DiskStream
      End If
   End If
   
   mo_StdIn.WriteFromByteArr pa_Bytes
End Sub

Public Function ContentLength() As Long
   ContentLength = mo_StdIn.GetSize
End Function

Public Function HasContent() As Boolean
   HasContent = (Me.ContentLength > 0)
End Function

Public Function Content() As Byte()
   Dim la_Bytes() As Byte
   Dim l_CurPos As Long
   
   ' Get all available content in STDIN
   
   If Me.HasContent Then
      l_CurPos = mo_StdIn.GetPosition  ' Record current stream position
      
      mo_StdIn.SetPosition 0  ' Go to beginning of stream
      mo_StdIn.ReadToByteArr la_Bytes  ' Get all bytes
      
      mo_StdIn.SetPosition l_CurPos ' Reset stream position
      
   Else
      la_Bytes = vbNullString   ' Return empty array. LBound = 0, UBound = -1
   End If

   Content = la_Bytes
End Function

Public Sub SaveToFile(ByVal p_FilePath As String)
   libRc5Factory.C.FSO.WriteByteContent p_FilePath, Me.Content
End Sub

Private Sub Class_Initialize()
   Set mo_StdIn = libRc5Factory.C.Stream
End Sub

Private Sub Class_Terminate()
   On Error Resume Next
   
   ' Remove STDIN data if it was stored on file system
   If m_FilePath <> "" Then
      If libRc5Factory.C.FSO.FileExists(m_FilePath) Then
         libRc5Factory.C.FSO.DeleteFile m_FilePath
      End If
   End If
End Sub

