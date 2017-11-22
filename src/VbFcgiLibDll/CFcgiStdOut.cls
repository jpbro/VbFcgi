VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFcgiStdOut"
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

Public Enum e_FcgiStdoutState
   fcgistdoutstate_Initialized   ' The buffer is ready for data to be appended
   fcgistdoutstate_Building   ' Content has been and/or is being appended to the STDOUT buffer
   fcgistdoutstate_Built   ' Can send FCGI_END_REQUEST after buffer flushes
End Enum

Private mo_StdOut As vbRichClient5.cArrayList
Private m_State As e_FcgiStdoutState ' The current state of this object.
Private m_ContentLength As Long  ' The current unflushed content length

Public Property Get State() As e_FcgiStdoutState
   State = m_State
End Property

Public Property Let State(ByVal p_State As e_FcgiStdoutState)
   If m_State > p_State Then Err.Raise 5, , "Can't revert state."
   
   m_State = p_State
End Property

Public Sub AppendContent(pa_Bytes As Variant)
   Dim l_Len As Long
   Dim l_Mod As Long
   Dim ii As Long
   Dim la_Bytes() As Byte
   Dim l_SafeArrayPointer As Long
   Dim l_DataPointer As Long
   
   If Me.State = fcgistdoutstate_Built Then Err.Raise 5, , "Can't append content to built STDOUT"
   Me.State = fcgistdoutstate_Building
      
   l_Len = UBound(pa_Bytes) + 1
   apiOutputDebugString "Appending content to STDOUT. Length: " & l_Len
   
   m_ContentLength = m_ContentLength + l_Len
      
   If l_Len > gc_MaxStdoutBufferChunkSize Then
      ' Length of passed content larger than our maximum STDOUT buffer chunk size
      ' So split into smaller chunks and queue those chunks
      
      ReDim la_Bytes(0 To gc_MaxStdoutBufferChunkSize - 1)
      
      apiCopyMemory l_SafeArrayPointer, ByVal VarPtr(pa_Bytes) + 8, 4
      
      If apiSafeArrayAccessData(l_SafeArrayPointer, l_DataPointer) = 0 Then
         For ii = 0 To (l_Len \ gc_MaxStdoutBufferChunkSize) - 1
            apiOutputDebugString "Appending chunk. Length: " & gc_MaxStdoutBufferChunkSize
            apiOutputDebugString "Chunk Start: " & (gc_MaxStdoutBufferChunkSize * ii)
            
            apiCopyMemory la_Bytes(0), ByVal l_DataPointer + (gc_MaxStdoutBufferChunkSize * ii), gc_MaxStdoutBufferChunkSize
            
            mo_StdOut.Queue la_Bytes
         Next ii
         
         l_Mod = l_Len Mod gc_MaxStdoutBufferChunkSize
         If l_Mod > 0 Then
            apiOutputDebugString "Appending final chunk. Length: " & l_Mod
            
            ReDim la_Bytes(0 To l_Mod - 1)
            apiCopyMemory la_Bytes(0), ByVal l_DataPointer + (gc_MaxStdoutBufferChunkSize * ii), l_Mod
      
            mo_StdOut.Queue la_Bytes
         End If
         
         apiSafeArrayUnaccessData l_SafeArrayPointer
      End If
      
   Else
      ' Length of passed content smaller or equal to our max STDOUT buffer chunk size
      ' Queue the entire chunk
      
      mo_StdOut.Queue pa_Bytes
   End If
End Sub

Public Function HasUnflushedContent() As Boolean
   HasUnflushedContent = (mo_StdOut.Count > 0)
End Function

Public Function UnflushedContentLength() As Long
   UnflushedContentLength = m_ContentLength
End Function

Public Function NextContentChunk() As Byte()
   If mo_StdOut.Count <= 0 Then Err.Raise vbObjectError, , "No unflushed content available."
   
   NextContentChunk = mo_StdOut.DeQueue
   
   m_ContentLength = m_ContentLength - (UBound(NextContentChunk) + 1)
End Function

Private Sub Class_Initialize()
   Set mo_StdOut = libRc5Factory.C.ArrayList(vbVariant)
End Sub
