VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBuilderJson"
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

Implements VbFcgiLib.IBuilder
Implements VbFcgiLib.IJson

Public Event Finished()

Private WithEvents mo_HttpHeader As VbFcgiLib.CHttpHeader
Attribute mo_HttpHeader.VB_VarHelpID = -1
Private mo_Collection As vbRichClient5.cCollection
Private m_Finished As Boolean
Private ma_Content() As Byte
Private m_Length As Long
Private m_IncludeCharsetInHttpHeader As Boolean
Private m_Initialized As Boolean
Private m_StrictParsing As Boolean
Private m_SkipInitializedCheck As Boolean

Public Property Get IJsonObject() As VbFcgiLib.IJson
   Set IJsonObject = Me
End Property

Public Property Get StrictParsing() As Boolean
   StrictParsing = m_StrictParsing
End Property

Public Property Let StrictParsing(ByVal p_StrictParsing As Boolean)
   m_StrictParsing = p_StrictParsing
End Property

Public Property Get IncludeCharsetInHttpHeader() As Boolean
   IncludeCharsetInHttpHeader = m_IncludeCharsetInHttpHeader
End Property

Public Property Let IncludeCharsetInHttpHeader(ByVal p_Include As Boolean)
   m_IncludeCharsetInHttpHeader = p_Include
End Property

Private Sub Class_Initialize()
   IBuilder_Reset
End Sub

Private Sub IJson_AddJsonObjectByKeyValuePairs(ParamArray pa_KeyAndValuePairs() As Variant)
   Dim ii As Long
   Dim lo_Collection As vbRichClient5.cCollection
   Dim l_ArraySize As Long
   
   If m_Finished Then Err.Raise 5, , "Already finished."
   
   l_ArraySize = UBound(pa_KeyAndValuePairs) - LBound(pa_KeyAndValuePairs) + 1
   If l_ArraySize < 1 Then Err.Raise 5, , "At least one key and value pair is required."
   If l_ArraySize Mod 2 <> 0 Then Err.Raise 5, , "Key/value pairs must balanced (one key and one value for each pair)."
   
   Set lo_Collection = libRc5Factory.C.JSONObject
   
   For ii = LBound(pa_KeyAndValuePairs) To UBound(pa_KeyAndValuePairs) Step 2
      lo_Collection.Add pa_KeyAndValuePairs(ii + 1), pa_KeyAndValuePairs(ii)
   Next ii
   
   If mo_Collection Is Nothing Then
      Set mo_Collection = libRc5Factory.C.JSONDecodeToCollection(lo_Collection.SerializeToJSONString)
      m_Initialized = True
   Else
      If mo_Collection.Count = 0 Then
         Set mo_Collection = libRc5Factory.C.JSONDecodeToCollection(lo_Collection.SerializeToJSONString)
      Else
         mo_Collection.Add lo_Collection
      End If
   End If
End Sub

Private Sub IJson_AddJsonArrayByValues(ParamArray pa_Values() As Variant)
   Dim ii As Long
   Dim l_ArraySize As Long
   Dim lo_Collection As vbRichClient5.cCollection
   
   If m_Finished Then Err.Raise 5, , "Already finished."
   
   l_ArraySize = UBound(pa_Values) - LBound(pa_Values) + 1
   If l_ArraySize < 1 Then Err.Raise 5, , "At least one value is required."
   
   Set lo_Collection = libRc5Factory.C.JSONArray
   
   For ii = LBound(pa_Values) To UBound(pa_Values)
      lo_Collection.Add pa_Values(ii)
   Next ii
   
   If mo_Collection Is Nothing Then
      Set mo_Collection = libRc5Factory.C.JSONDecodeToCollection(lo_Collection.SerializeToJSONString)
      m_Initialized = True
   Else
      If mo_Collection.Count = 0 Then
         Set mo_Collection = libRc5Factory.C.JSONDecodeToCollection(lo_Collection.SerializeToJSONString)
      Else
         mo_Collection.Add lo_Collection
      End If
   End If
End Sub

Private Property Set IJson_JsonRc5Collection(po_JsonCollection As vbRichClient5.cCollection)
   m_SkipInitializedCheck = True
   Me.Initialize po_JsonCollection
End Property

Private Property Get IJson_JsonRc5Collection() As vbRichClient5.cCollection
   Set IJson_JsonRc5Collection = mo_Collection
End Property

Private Property Let IJson_JsonString(ByVal p_JsonString As String)
   m_SkipInitializedCheck = True
   Me.Initialize p_JsonString
End Property

Private Property Get IJson_JsonString() As String
   ' Serialize the current JSON content and return as a VB String
   ' Empty string will be return on no content
   
   If mo_Collection Is Nothing Then
      IJson_JsonString = "{}"
   Else
      IJson_JsonString = mo_Collection.SerializeToJSONString
   End If
End Property

Private Property Let IJson_JsonUtf8(pa_JsonUtf8Bytes() As Byte)
   m_SkipInitializedCheck = True
   Me.Initialize pa_JsonUtf8Bytes
End Property

Private Property Get IJson_JsonUtf8() As Byte()
   ' Serialize the current JSON content and return as a VB String
   ' Empty string will be return on no content
   
   If mo_Collection Is Nothing Then
      IJson_JsonUtf8 = stringVbToUtf8("{}")
   Else
      IJson_JsonUtf8 = mo_Collection.SerializeToJSONUTF8
   End If
End Property

Private Property Get IBuilder_Content() As Byte()
   ' You must call the Finish method before you call this property.
   ' Call this property to get your finished content as a Me.ContentEncoding format Byte Array
   
   If Not m_Finished Then Err.Raise 5, , "You must call the Finished method before requesting content."
   
   IBuilder_Content = ma_Content
End Property

Private Property Get IBuilder_HttpHeader() As VbFcgiLib.CHttpHeader
   Set IBuilder_HttpHeader = mo_HttpHeader
End Property

Private Sub IBuilder_Finish(Optional ByVal p_ContentEncoding As VbFcgiLib.e_ContentEncoding = contentencoding_RecommendedDefault)
   ' Call this method when you are finished building your JSON
   ' Once called, the Content, Length, and ContentEncoding properties will become usable.
      
   If m_Finished Then Err.Raise 5, , "Already finished."
      
   If p_ContentEncoding = contentencoding_RecommendedDefault Then p_ContentEncoding = contentencoding_UTF8
      
   Select Case p_ContentEncoding
   Case contentencoding_UTF8
      If mo_Collection Is Nothing Then
         ' No JSON content available, return empty object
         ma_Content = stringVbToUtf8("{}")
      
      Else
         ma_Content = mo_Collection.SerializeToJSONUTF8
      End If
      
   Case Else
      ' Only allow UTF-8
      Err.Raise 5, , "Only UTF-8 content encoding is accepted for JSON strings."
      
   End Select
   
   Set mo_Collection = Nothing
   m_Finished = True
   m_Length = arraySize(ma_Content)
   
   RaiseEvent Finished
End Sub

Private Property Get IBuilder_IsFinished() As Boolean
   IBuilder_IsFinished = m_Finished
End Property

Private Property Get IBuilder_Length() As Long
   ' You must call the Finish method before you call this property.
   
   ' This property returns the length of the Content byte array.
   
   If Not m_Finished Then Err.Raise 5, , "You must call the Finished method before requesting the content length."
   
   IBuilder_Length = m_Length
End Property

Private Sub IBuilder_Reset()
   m_StrictParsing = True
   m_SkipInitializedCheck = False
   
   m_Initialized = False
   m_IncludeCharsetInHttpHeader = False
   m_Length = 0
   
   m_Finished = False
   
   Erase ma_Content
   
   Set mo_HttpHeader = New VbFcgiLib.CHttpHeader
   Set mo_Collection = Nothing
End Sub

Public Sub Finish(Optional ByVal p_ContentEncoding As VbFcgiLib.e_ContentEncoding = contentencoding_RecommendedDefault)
   IBuilder_Finish p_ContentEncoding
End Sub

Public Function Content() As Byte()
   Content = IBuilder_Content
End Function

Public Function Length() As Long
   Length = IBuilder_Length
End Function

Public Function IsFinished() As Boolean
   IsFinished = IBuilder_IsFinished
End Function

Private Sub mo_HttpHeader_BeforeBuildHttpHeader()
   ' Make sure the HTTP header includes the "Content-Type: text/html"
   ' and "Content-Length" headers.
   
   Dim l_Charset As String
   
   ' Add Content-Type header if it doesn't already exist
   If Not mo_HttpHeader.Exists("Content-Type") Then
      With mo_HttpHeader.Add("Content-Type", mimeTypeFromFilePath(".json"))
         
         If m_IncludeCharsetInHttpHeader Then
            ' JSON only supports UTF-8 so we shouldn't need to include the charset
            ' But there are lots of surprises out there in the wild wild west of the world wide web.
            
            With .KeyValuePairs()
               l_Charset = "utf-8"
               
               If Not stringIsEmptyOrWhitespaceOnly(l_Charset) Then
                  .Add "charset", l_Charset
               End If
            End With
         End If
         
      End With
   End If
   
   ' Add or Replace Content-Length header if HTML build is finished
   If Me.IsFinished Then
      mo_HttpHeader.AddOrReplace "Content-Length", Me.Length
   End If
End Sub

Public Sub Initialize(p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList As Variant)
   ' p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList (how's that for a name) can be any of the following:
   ' - A JSON string (VB6 String/UTF-16LE)
   ' - A JSON Byte Array (UTF-8)
   ' - A vbRichClient5 cArrayList object
   ' - A VbRichClient5 cCollection object that has been created as a JSONObject or JSONArray
   ' - A non-byte Array of any type (it will be added to the collectionas a JSONArray)
   ' - Nothing object - an empty JSONObject collection will be created
   
   Dim lo_ArrayList As vbRichClient5.cArrayList
   Dim la_Bytes() As Byte
   Dim l_ParserError As String
   
   If m_SkipInitializedCheck Then
      m_SkipInitializedCheck = False
   Else
      If m_Initialized Then Err.Raise 5, , "Already initialized."
   End If
   
   If m_Finished Then Err.Raise 5, , "Already finished."
   
   If IsArray(p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList) Then
      If VarType(p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList) = (vbByte Or vbArray) Then
         ' Assume JSON UTF8 bytes
         la_Bytes = p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList
         Set mo_Collection = libRc5Factory.C.JSONDecodeToCollectionUTF8(la_Bytes, l_ParserError, m_StrictParsing)
         Erase la_Bytes
         
         If Not stringIsEmptyOrWhitespaceOnly(l_ParserError) Then
            Set mo_Collection = Nothing
            Err.Raise vbObjectError, , "Error parsing JSON: " & l_ParserError
         End If
      
      Else
         ' Bind the passed array to an RC5 array list
         ' We will add it to a JSON collection as a JSON Array below
         ' Move array data into a JSONArray collection
         AddToJsonCollection p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList
      End If
      
   Else
      ' We don't have an array
      
      If IsObject(p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList) Then
         If p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList Is Nothing Then
            ' Nothing object passed, so initialize an empty JSON object collection
            Set mo_Collection = libRc5Factory.C.JSONObject
         Else
            ' We have an object - test for RC5 Collection or ArrayList
            If TypeOf p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList Is vbRichClient5.cCollection Then
               ' We have an RC5 collection
               Set mo_Collection = p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList
               If Not (mo_Collection.IsJSONArray Or mo_Collection.IsJSONObject) Then
                  Set mo_Collection = Nothing
                  Err.Raise 5, , "Passed collection must be a JSON array or JSON object collection."
               End If
               
            ElseIf TypeOf p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList Is vbRichClient5.cArrayList Then
               ' We have an RC5 array list
               ' Move array list data into a JSONArray collection
               AddToJsonCollection p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList
            
            ElseIf TypeOf p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList Is VbFcgiLib.IJson Then
               ' We have an IJson object, so get a JSON RC5 cCollection class from it
               Set mo_Collection = p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList.IJson.JsonCollection
            
            Else
               ' Unusable object type detected
               Err.Raise 5, , "Unknown object type for initialize: " & TypeName(p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList)
            End If
         End If
         
      Else
         ' Passed data is not an object or array
         
         If VarType(p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList) = vbString Then
            ' Assume JSON string
            
            la_Bytes = p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList
            Set mo_Collection = libRc5Factory.C.JSONDecodeToCollection(CStr(la_Bytes), l_ParserError, m_StrictParsing)
            Erase la_Bytes
            
            If Not stringIsEmptyOrWhitespaceOnly(l_ParserError) Then
               Set mo_Collection = Nothing
               Err.Raise vbObjectError, , "Error parsing JSON: " & l_ParserError
            End If
            
         Else
            Err.Raise 5, , "Bad variable type for initialize: " & VarType(p_JsonStringOrUtf8BytesOrVbArrayOrRc5CollectionOrArrayList)
         
         End If
      End If
   End If
   
   If mo_Collection Is Nothing Then
      If lo_ArrayList Is Nothing Then
         ' This should not happen!
         ' If we get here, fix logic above to handle the case that lead here
         Err.Raise vbObjectError, , "No data available for JSON collection."
      End If
   End If
   
   m_Initialized = True
End Sub

Private Sub AddToJsonCollection(p_ArrayRc5ArrayListOrRc5Collection As Variant)
   Dim ii As Long
   Dim l_Lbound As Long
   Dim l_Ubound As Long
   Dim l_AddCollection As Boolean
   Dim l_IsArray As Boolean
            
   If IsArray(p_ArrayRc5ArrayListOrRc5Collection) Then
      l_IsArray = True
      l_Lbound = LBound(p_ArrayRc5ArrayListOrRc5Collection)
      l_Ubound = UBound(p_ArrayRc5ArrayListOrRc5Collection)
      
   Else
      If IsObject(p_ArrayRc5ArrayListOrRc5Collection) Then
         If TypeOf p_ArrayRc5ArrayListOrRc5Collection Is vbRichClient5.cArrayList Then
            l_IsArray = True
            l_Ubound = p_ArrayRc5ArrayListOrRc5Collection.Count - 1
         
         ElseIf TypeOf p_ArrayRc5ArrayListOrRc5Collection Is vbRichClient5.cCollection Then
            If p_ArrayRc5ArrayListOrRc5Collection.IsJSONArray Then
               l_AddCollection = True
            ElseIf p_ArrayRc5ArrayListOrRc5Collection.IsJSONObject Then
               l_AddCollection = True
            Else
               Err.Raise 5, , "Collections must be JSONArrays or JSONObject."
            End If
         Else
            Err.Raise 5, , "Unusable object type: " & TypeName(p_ArrayRc5ArrayListOrRc5Collection)
         End If
      Else
         Err.Raise 5, , "Unusable type: " & VarType(p_ArrayRc5ArrayListOrRc5Collection)
      End If
   End If
   
   If mo_Collection Is Nothing Then
      If l_IsArray Then
         Set mo_Collection = libRc5Factory.C.JSONArray
      Else
         Set mo_Collection = libRc5Factory.C.JSONObject
      End If
      
      m_Initialized = True
   End If
   
   If l_AddCollection Then
      mo_Collection.Add p_ArrayRc5ArrayListOrRc5Collection
   Else
      For ii = l_Lbound To l_Ubound
         If IsArray(p_ArrayRc5ArrayListOrRc5Collection(ii)) Then
            AddToJsonCollection p_ArrayRc5ArrayListOrRc5Collection(ii)
         Else
            If IsObject(p_ArrayRc5ArrayListOrRc5Collection(ii)) Then
               If TypeOf p_ArrayRc5ArrayListOrRc5Collection(ii) Is vbRichClient5.cArrayList Then
                  AddToJsonCollection p_ArrayRc5ArrayListOrRc5Collection(ii)
               ElseIf TypeOf p_ArrayRc5ArrayListOrRc5Collection(ii) Is vbRichClient5.cCollection Then
                  AddToJsonCollection p_ArrayRc5ArrayListOrRc5Collection(ii)
               Else
                  mo_Collection.Add p_ArrayRc5ArrayListOrRc5Collection(ii)
               End If
            Else
               mo_Collection.Add p_ArrayRc5ArrayListOrRc5Collection(ii)
            End If
         End If
      Next ii
   End If
End Sub
