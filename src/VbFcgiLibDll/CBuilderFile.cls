VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBuilderFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum e_EtagType
   etag_Custom = -1 ' Set the ETag property
   etag_None   ' Don't generate an etag
   
   etag_Md5Hex
   etag_Sha1Hex
   etag_Sha256Hex
   etag_Sha384Hex
   etag_Sha512Hex

   etag_Md5Base64
   etag_Sha1Base64
   etag_Sha256Base64
   etag_Sha384Base64
   etag_Sha512Base64
End Enum

Implements VbFcgiLib.IBuilder

Public Event Finished()

Private WithEvents mo_HttpHeader As VbFcgiLib.CHttpHeader
Attribute mo_HttpHeader.VB_VarHelpID = -1

Private m_Finished As Boolean
Private m_ChunkSize As Long   ' When calling content, chunks of this size will be streamed sequentially each .Content call.
                              ' If <1, then slurp and return the entire file contents
Private m_FinalChunkSize As Long ' The size of the final chunk
Private m_ChunkCount As Long  ' Total # of chunks to read/stream
Private m_ChunkIndex As Long   ' Zero based index of next chunk to read/stream
Private m_FilePath As String
Private m_FileLen As Currency
Private mo_Stream As vbRichClient5.cStream
Private m_ETag As String
Private m_NotModified As Boolean ' If true, then send 304 NOT MODIFIED downstream

Public Property Get ETag() As String
   ETag = m_ETag
End Property

Public Property Let ETag(ByVal p_ETag As String)
   m_ETag = stringRemoveWhitespace(p_ETag)
End Property

Public Property Get ChunkSize() As Long
   ChunkSize = m_ChunkSize
End Property

Public Property Get ChunkCount() As Long
   ChunkCount = m_ChunkCount
End Property

Public Property Get ChunkIndex() As Long
   ChunkIndex = m_ChunkIndex
End Property

Public Property Get FilePath() As String
   FilePath = m_FilePath
End Property

Public Property Get IsFinishedReading() As Boolean
   If m_NotModified Then
      IsFinishedReading = True
   Else
      IsFinishedReading = (m_ChunkIndex >= m_ChunkCount)
   End If
End Property

Public Sub OpenFile(ByVal p_FilePath As String, Optional ByVal p_ClientCachedEtag As String, Optional ByVal p_CalculateEtagByType As e_EtagType = etag_Sha256Hex, Optional ByVal p_ChunkSize As Long = 262144)
   Dim l_ClientCacheIsFresh As Boolean
   Dim l_EtagSourceData As String
   Dim l_LastWriteDate As Date
   
   On Error GoTo ErrorHandler
   
   If IsReadyToFinish Then Err.Raise 5, , "File already open."
   
   If Not libFso.FileExists(p_FilePath) Then Err.Raise 53, , "File not found at: " & p_FilePath
      
   libFso.GetFileAttributesEx p_FilePath, m_FileLen, , l_LastWriteDate
      
   If p_CalculateEtagByType > etag_None Then
      ' Calculate ETag
            
      ' Combine file path, length, and last modified date for hashing an Etag
      ' TODO: This is not perfect - I'd rather get a file hash but vbRichClient5 doesn't
      ' exposed an MD5/SHA methods for streamed file hashing.
      ' I've sent Olaf an email about this so hopefully we'll have it soon.
      
      l_EtagSourceData = p_FilePath & vbTab & m_FileLen & vbTab & dateVbLocalDateTimeToIso8601Utc(l_LastWriteDate)
      
      Select Case p_CalculateEtagByType
      Case etag_Md5Hex
         m_ETag = libCrypt.MD5(l_EtagSourceData, True)
         
      Case etag_Sha1Hex
         m_ETag = libCrypt.SHA1(l_EtagSourceData, True)
      
      Case etag_Sha256Hex
         m_ETag = libCrypt.SHA256(l_EtagSourceData, True)
      
      Case etag_Sha384Hex
         m_ETag = libCrypt.SHA384(l_EtagSourceData, True)
      
      Case etag_Sha512Hex
         m_ETag = libCrypt.SHA512(l_EtagSourceData, True)
      
      Case etag_Md5Base64
         m_ETag = libCrypt.Base64Enc(libCrypt.MD5(l_EtagSourceData, False))
      
      Case etag_Sha1Base64
         m_ETag = libCrypt.Base64Enc(libCrypt.SHA1(l_EtagSourceData, False))
      
      Case etag_Sha256Base64
         m_ETag = libCrypt.Base64Enc(libCrypt.SHA256(l_EtagSourceData, False))
      
      Case etag_Sha384Base64
         m_ETag = libCrypt.Base64Enc(libCrypt.SHA384(l_EtagSourceData, False))
      
      Case etag_Sha512Base64
         m_ETag = libCrypt.Base64Enc(libCrypt.SHA512(l_EtagSourceData, False))
      
      Case Else
         ' Unsupported ETag type
         Debug.Assert False
         Err.Raise 5, , "Unsupported ETag type: " & p_CalculateEtagByType
      End Select
      
      ' Compare client-side cached ETag to generate ETag.
      ' If they are the same, we should return 304 NOT MODIFIED instead of the file content
      ' This will improve performance and save on bandwidth
      If Not stringIsEmptyOrWhitespaceOnly(p_ClientCachedEtag) Then
         p_ClientCachedEtag = p_ClientCachedEtag
         
         m_NotModified = (p_ClientCachedEtag = m_ETag)
      End If
   End If
   
   If Not m_NotModified Then
      ' Client cache is stale, doesn't exist, or we don't know the local file ETag
      
      ' Open the file stream and prevent writes to it while we stream it out
      Set mo_Stream = libFso.OpenFileStream(p_FilePath, STRM_SHARE_DENY_WRITE, False)
            
      m_ChunkCount = mathIntDivideCurrency(m_FileLen, p_ChunkSize)
      m_FinalChunkSize = mathModCurrency(m_FileLen, p_ChunkSize)
      
      Select Case m_FinalChunkSize
      Case 0
         ' File size is evenly divisible by chunk size, so set final chunk size to chunk size.
         m_FinalChunkSize = m_ChunkSize
         
      Case Is > 0
         ' Increment the total chunk count by one since our chunk size doesn't evenly divide into the file size
         m_ChunkCount = m_ChunkCount + 1
         
      Case Else
         Debug.Assert False
         Err.Raise vbObjectError, , "Unexpected negative chunk mod."
      
      End Select
      
      m_ChunkSize = p_ChunkSize
      m_FilePath = p_FilePath
   End If
   
   IBuilder_Finish

   Exit Sub
   
ErrorHandler:
   Dim l_ErrNum As Long
   Dim l_ErrDesc As String
   Dim l_ErrSource As String
   
   Debug.Assert False
   l_ErrNum = Err.Number
   l_ErrDesc = Err.Description
   l_ErrSource = Err.Source
   
   IBuilder_Reset
   
   Err.Raise l_ErrNum, l_ErrSource, l_ErrDesc
End Sub

Private Sub Class_Initialize()
   IBuilder_Reset
End Sub

Private Property Get IBuilder_Content() As Byte()
   Dim la_ChunkBytes() As Byte
   
   If Me.IsFinishedReading Then Err.Raise 5, , "All file chunks have been read."
   
   If m_NotModified Then
      IBuilder_Content = la_ChunkBytes
   
   Else
      If m_ChunkIndex = m_ChunkCount - 1 Then
         ' Final chunk
         ReDim la_ChunkBytes(m_FinalChunkSize - 1)
      Else
         ReDim la_ChunkBytes(m_ChunkSize - 1)
      End If
      
      mo_Stream.ReadToByteArr la_ChunkBytes, arraySize(la_ChunkBytes)
      
      m_ChunkIndex = m_ChunkIndex + 1
   
      IBuilder_Content = la_ChunkBytes
   End If
End Property

Private Function IsReadyToFinish() As Boolean
   If Not mo_Stream Is Nothing Then
      ' Stream is open, ready to finish and send downstream
      IsReadyToFinish = True
   Else
      ' File is cached on client, ready to finish and send 304 NOT MODIFIED downstream
      IsReadyToFinish = m_NotModified
   End If
End Function

Private Sub IBuilder_Finish(Optional ByVal p_ContentEncoding As e_ContentEncoding = -1&)
   If Not IsReadyToFinish Then Err.Raise 5, , "No file has been opened, nothing to finish."
   If m_Finished Then Err.Raise 5, , "Already called finish."
      
   m_Finished = True
   RaiseEvent Finished
End Sub

Private Property Get IBuilder_HttpHeader() As CHttpHeader
   Set IBuilder_HttpHeader = mo_HttpHeader
End Property

Private Property Get IBuilder_IsFinished() As Boolean
   IBuilder_IsFinished = m_Finished
End Property

Private Property Get IBuilder_Length() As Long
   If Not m_Finished Then Err.Raise 5, , "You must call the Finished method before requesting the content length."
   
   IBuilder_Length = m_FileLen
End Property

Public Function Content() As Byte()
   Content = IBuilder_Content
End Function

Private Sub IBuilder_Reset()
   ' Reset all objects and variables
   Set mo_HttpHeader = New VbFcgiLib.CHttpHeader
   m_Finished = False
   Set mo_Stream = Nothing
   m_ChunkSize = 0
   m_ChunkCount = 0
   m_ChunkIndex = 0
   m_FilePath = ""
   m_FileLen = 0
   m_ETag = ""
   m_NotModified = False
End Sub

Private Sub mo_HttpHeader_BeforeBuildHttpHeader()
   If m_Finished Then
      ' Add Content-Type if necessary
      
      If m_NotModified Then
         ' Client cached file is fresh, send 304 NOT MODIFIED
         
         With mo_HttpHeader.AddOrReplace("Status")
            With .InitializeByKeyValuePairs()
               .Add "", "304"
            End With
         End With
                  
         With mo_HttpHeader.Add("Content-Length")
            With .InitializeByKeyValuePairs()
               .Add "", "0"
            End With
         End With
         
      Else
         ' Build HTTP header for file data
         
         If Not mo_HttpHeader.Exists("Content-Type") Then
            With mo_HttpHeader.Add("Content-Type")
               With .InitializeByKeyValuePairs()
                  .Add "", mimeTypeFromFilePath(m_FilePath)
               End With
            End With
         End If
      
         ' ETag
         If Not stringIsEmptyOrWhitespaceOnly(m_ETag) Then
            With mo_HttpHeader.AddOrReplace("ETag")
               With .InitializeByKeyValuePairs()
                  .Add "", m_ETag
               End With
            End With
         End If
         
         ' Set file size
         With mo_HttpHeader.AddOrReplace("Content-Length")
            With .InitializeByKeyValuePairs()
               .Add "", m_FileLen
            End With
         End With
      End If
   End If
End Sub
