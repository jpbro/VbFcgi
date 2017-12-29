VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IJson"
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

Public Property Get JsonString() As String
   ' Get the data held in the JSON collection as a VB6 UTF16LE string
End Property

Public Property Let JsonString(ByVal p_JsonString As String)
   ' Initialize a IJson object by a JSON string (VB/UTF16LE).
   ' This will replace any existing JSON data
End Property

Public Property Get JsonUtf8() As Byte()
   ' Get the data held in the JSON collection as a UTF8 Byte Array
End Property

Public Property Let JsonUtf8(pa_JsonUtf8Bytes() As Byte)
   ' Initialize a IJson object by a JSON byte array (UTF8).
   ' This will replace any existing JSON data
End Property

Public Property Get JsonRc5Collection() As vbRichClient5.cCollection
   ' Get the JSON collection
End Property

Public Property Set JsonRc5Collection(po_Collection As vbRichClient5.cCollection)
   ' Initialize a IJson object using a vbRichClient5 cCollection class (JSONArray or JSONObject type).
   ' This will replace any existing JSON data
End Property

Public Sub AddJsonObjectByKeyValuePairs(ParamArray p_KeyAndValuePairs() As Variant)
   ' Add data to the JSON collection by key and value pairs.
   ' Existing data will be left alone.
   ' For example: AddJsonObjectByKeyValuePairs "status", 200, "message", "OK"
   ' Will produce: {"status":200,"message":"OK"}
End Sub

Public Sub AddJsonArrayByValues(ParamArray p_Values() As Variant)
   ' Add to to JSON collection by an array of values.
   ' Existing data will be left alone.
   ' For Example: AddJsonArrayByValues 1,2,3,4,5
   ' Will produce: [1,2,3,4,5]
End Sub

