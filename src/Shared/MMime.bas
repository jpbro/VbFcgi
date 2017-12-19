Attribute VB_Name = "MMime"
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

Private mo_MimeTypesRs As vbRichClient5.cRecordset

Public Function mimeTypeFromFilePath(ByVal p_FilePath As String) As String
   ' Pass a file path (or file name) and get back a MIME type string.
   ' Primary used for HTTP Content-Type headers.
   
   Dim lo_Db As vbRichClient5.cConnection
   Dim lo_InsertCmd As vbRichClient5.cCommand
   Dim l_TypeInfo As String
   Dim la_TypeInfo() As String
   Dim ii As Long
   
   If (mo_MimeTypesRs Is Nothing) Or envRunningInIde Then
      ' First pass (or running in IDE), So populate the MIME type/extension recordset.
      
      Set lo_Db = libRc5Factory.C.Connection(, DBCreateInMemory, , False)
      
      lo_Db.Execute "CREATE TABLE mimetypes (mimetype TEXT, extension TEXT)"
      lo_Db.Execute "CREATE UNIQUE INDEX uidx_mimetypes_extensions on mimetypes (extension ASC)"
      
      Set lo_InsertCmd = lo_Db.CreateCommand("INSERT INTO mimetypes (mimetype, extension) VALUES (?,?)")
      
      ' MIME type info string are tab delimited strings
      ' with the first element being the commonly used file extension
      ' and the second element being the MIME type identifier.
      For ii = 0 To &H7FFFFFFF
         l_TypeInfo = mimeTypeInfoByIndex(ii)
         
         If InStr(1, l_TypeInfo, vbTab) Then
            la_TypeInfo = Split(l_TypeInfo, vbTab)
         
            With lo_InsertCmd
               .SetAllParamsNull
               
               .SetText 1, la_TypeInfo(1)
               .SetText 2, la_TypeInfo(0)
               
               .Execute
            End With
         
         Else
            ' No more MIME type infos available
            Exit For
         End If
      Next ii
      
      ' Create the module level disconnected recordset
      Set mo_MimeTypesRs = lo_Db.OpenRecordset("SELECT * FROM mimetypes")
      Set mo_MimeTypesRs.ActiveConnection = Nothing
   End If
   
   ' Get the file extension only
   p_FilePath = LCase$(stringTrimWhitespace(libFso.GetFileExtension(p_FilePath)))
   
   ' Look for the file extension in our database
   If mo_MimeTypesRs.FindFirst("extension='" & Replace$(p_FilePath, "'", "''") & "'") Then
      ' Found the file extension, so grab the MIME type
      mimeTypeFromFilePath = "" & mo_MimeTypesRs.Fields("mimetype").Value
   End If
   
   If stringIsEmptyOrWhitespaceOnly(mimeTypeFromFilePath) Then
      ' MIME type not found for the file extension, default to application/octet-stream
      Debug.Assert False
      
      apiOutputDebugString "Unknown extension for MIME type: " & p_FilePath & ". Default to application/octet-stream."
      
      mimeTypeFromFilePath = "application/octet-stream"
   End If
End Function

Private Function mimeTypeInfoByIndex(ByVal p_TypeIndex As Long) As String
   Dim l_TypeInfo As String

   ' This is a small subset of all the MIME types/extensions
   ' (just ones that I have encountered more commonly).
   ' I've found a larger/more complete list out there but it is copyrighted.
   ' I've contacted the author but have yet to get permission to use the larger list.
   ' If it doesn't come I'll just add things here organically as I or users encounter them.

   Select Case p_TypeIndex
   Case 0
      ' 7-Zip
      l_TypeInfo = "7z" & vbTab & "application/x-7z-compressed"
   Case 1
      ' Advanced Audio Coding (AAC)
      l_TypeInfo = "aac" & vbTab & "audio/x-aac"
   Case 2
      ' Audio Video Interleave (AVI)
      l_TypeInfo = "avi" & vbTab & "video/x-msvideo"
   Case 3
      ' Bitmap Image File
      l_TypeInfo = "bmp" & vbTab & "image/bmp"
   Case 4
      ' Cascading Style Sheets (CSS)
      l_TypeInfo = "css" & vbTab & "text/css"
   Case 5
      ' Comma-Seperated Values
      l_TypeInfo = "csv" & vbTab & "text/csv"
   Case 6
      ' Microsoft Word
      l_TypeInfo = "doc" & vbTab & "application/msword"
   Case 7
      ' Microsoft Office - OOXML - Word Document
      l_TypeInfo = "docx" & vbTab & "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
   Case 8
      ' Autodesk Design Web Format (DWF)
      l_TypeInfo = "dwf" & vbTab & "model/vnd.dwf"
   Case 9
      ' DWG Drawing
      l_TypeInfo = "dwg" & vbTab & "image/vnd.dwg"
   Case 10
      ' AutoCAD DXF
      l_TypeInfo = "dxf" & vbTab & "image/vnd.dxf"
   Case 11
      ' Email Message
      l_TypeInfo = "eml" & vbTab & "message/rfc822"
   Case 12
      ' Microsoft Application
      l_TypeInfo = "exe" & vbTab & "application/x-msdownload"
   Case 13
      ' Flash Video
      l_TypeInfo = "f4v" & vbTab & "video/x-f4v"
   Case 14
      ' Flash Video
      l_TypeInfo = "flv" & vbTab & "video/x-flv"
   Case 15
      ' Graphics Interchange Format
      l_TypeInfo = "gif" & vbTab & "image/gif"
   Case 16
      ' HyperText Markup Language (HTML)
      l_TypeInfo = "html" & vbTab & "text/html"
   Case 17
      ' Icon Image
      l_TypeInfo = "ico" & vbTab & "image/x-icon"
   Case 18
      ' iCalendar
      l_TypeInfo = "ics" & vbTab & "text/calendar"
   Case 19
      ' Java Archive
      l_TypeInfo = "jar" & vbTab & "application/java-archive"
   Case 20
      ' Java Source File
      l_TypeInfo = "java" & vbTab & "text/x-java-source,java"
   Case 21
      ' JPEG Image
      l_TypeInfo = "jpeg" & vbTab & "image/jpeg"
   Case 22
      ' JPEG Image
      l_TypeInfo = "jpg" & vbTab & "image/jpeg"
   Case 23
      ' JPGVideo
      l_TypeInfo = "jpgv" & vbTab & "video/jpeg"
   Case 24
      ' JPEG 2000 Compound Image File Format
      l_TypeInfo = "jpm" & vbTab & "video/jpm"
   Case 25
      ' JavaScript
      l_TypeInfo = "js" & vbTab & "application/javascript"
   Case 26
      ' JavaScript Object Notation (JSON)
      l_TypeInfo = "json" & vbTab & "application/json"
   Case 27
      ' Google Earth - KML
      l_TypeInfo = "kml" & vbTab & "application/vnd.google-earth.kml+xml"
   Case 28
      ' Google Earth - Zipped KML
      l_TypeInfo = "kmz" & vbTab & "application/vnd.google-earth.kmz"
   Case 29
      ' M3U (Multimedia Playlist)
      l_TypeInfo = "m3u" & vbTab & "audio/x-mpegurl"
   Case 30
      ' Multimedia Playlist Unicode
      l_TypeInfo = "m3u8" & vbTab & "application/vnd.apple.mpegurl"
   Case 31
      ' Microsoft Access
      l_TypeInfo = "mdb" & vbTab & "application/x-msaccess"
   Case 32
      ' MIDI - Musical Instrument Digital Interface
      l_TypeInfo = "mid" & vbTab & "audio/midi"
   Case 33
      ' Quicktime Video
      l_TypeInfo = "mov" & vbTab & "video/quicktime"
   Case 34
      ' MPEG-4 Video
      l_TypeInfo = "mp4" & vbTab & "video/mp4"
   Case 35
      ' MPEG-4 Audio
      l_TypeInfo = "mp4a" & vbTab & "audio/mp4"
   Case 36
      ' MPEG Video
      l_TypeInfo = "mpg" & vbTab & "video/mpeg"
   Case 37
      ' MPEG Video
      l_TypeInfo = "mpeg" & vbTab & "video/mpeg"
   Case 38
      ' MPEG Audio
      l_TypeInfo = "mpga" & vbTab & "audio/mpeg"
   Case 39
      ' Microsoft Project
      l_TypeInfo = "mpp" & vbTab & "application/vnd.ms-project"
   Case 40
      ' OpenDocument Database
      l_TypeInfo = "odb" & vbTab & "application/vnd.oasis.opendocument.database"
   Case 41
      ' OpenDocument Chart
      l_TypeInfo = "odc" & vbTab & "application/vnd.oasis.opendocument.chart"
   Case 42
      ' OpenDocument Formula
      l_TypeInfo = "odf" & vbTab & "application/vnd.oasis.opendocument.formula"
   Case 43
      ' OpenDocument Formula Template
      l_TypeInfo = "odft" & vbTab & "application/vnd.oasis.opendocument.formula-template"
   Case 44
      ' OpenDocument Graphics
      l_TypeInfo = "odg" & vbTab & "application/vnd.oasis.opendocument.graphics"
   Case 45
      ' OpenDocument Image
      l_TypeInfo = "odi" & vbTab & "application/vnd.oasis.opendocument.image"
   Case 46
      ' OpenDocument Text Master
      l_TypeInfo = "odm" & vbTab & "application/vnd.oasis.opendocument.text-master"
   Case 47
      ' OpenDocument Presentation
      l_TypeInfo = "odp" & vbTab & "application/vnd.oasis.opendocument.presentation"
   Case 48
      ' OpenDocument Spreadsheet
      l_TypeInfo = "ods" & vbTab & "application/vnd.oasis.opendocument.spreadsheet"
   Case 49
      ' OpenDocument Text
      l_TypeInfo = "odt" & vbTab & "application/vnd.oasis.opendocument.text"
   Case 50
      ' Ogg Audio
      l_TypeInfo = "oga" & vbTab & "audio/ogg"
   Case 51
      ' Ogg Video
      l_TypeInfo = "ogv" & vbTab & "video/ogg"
   Case 52
      ' Ogg
      l_TypeInfo = "ogx" & vbTab & "application/ogg"
   Case 53
      ' Microsoft OneNote
      l_TypeInfo = "onetoc" & vbTab & "application/onenote"
   Case 54
      ' OpenType Font File
      l_TypeInfo = "otf" & vbTab & "application/x-font-otf"
   Case 55
      ' Adobe Portable Document Format
      l_TypeInfo = "pdf" & vbTab & "application/pdf"
   Case 56
      ' Portable Network Graphics (PNG)
      l_TypeInfo = "png" & vbTab & "image/png"
   Case 57
      ' Microsoft Office - OOXML - Presentation (Slideshow)
      l_TypeInfo = "ppsx" & vbTab & "application/vnd.openxmlformats-officedocument.presentationml.slideshow"
   Case 58
      ' Microsoft PowerPoint
      l_TypeInfo = "ppt" & vbTab & "application/vnd.ms-powerpoint"
   Case 59
      ' Microsoft Office - OOXML - Presentation
      l_TypeInfo = "pptx" & vbTab & "application/vnd.openxmlformats-officedocument.presentationml.presentation"
   Case 60
      ' Photoshop Document
      l_TypeInfo = "psd" & vbTab & "image/vnd.adobe.photoshop"
   Case 61
      ' Microsoft Publisher
      l_TypeInfo = "pub" & vbTab & "application/x-mspublisher"
   Case 62
      ' Quicktime Video
      l_TypeInfo = "qt" & vbTab & "video/quicktime"
   Case 63
      ' Real Audio Sound
      l_TypeInfo = "ram" & vbTab & "audio/x-pn-realaudio"
   Case 64
      ' RAR Archive
      l_TypeInfo = "rar" & vbTab & "application/x-rar-compressed"
   Case 65
      ' RealMedia
      l_TypeInfo = "rm" & vbTab & "application/vnd.rn-realmedia"
   Case 66
      ' RSS - Really Simple Syndication
      l_TypeInfo = "rss" & vbTab & "application/rss+xml"
   Case 67
      ' Scalable Vector Graphics (SVG)
      l_TypeInfo = "svg" & vbTab & "image/svg+xml"
   Case 68
      ' Adobe Flash
      l_TypeInfo = "swf" & vbTab & "application/x-shockwave-flash"
   Case 69
      ' Tar File (Tape Archive)
      l_TypeInfo = "tar" & vbTab & "application/x-tar"
   Case 70
      ' Tagged Image File Format
      l_TypeInfo = "tif" & vbTab & "image/tiff"
   Case 71
      ' Tagged Image File Format
      l_TypeInfo = "tiff" & vbTab & "image/tiff"
   Case 72
      ' BitTorrent
      l_TypeInfo = "torrent" & vbTab & "application/x-bittorrent"
   Case 73
      ' Tab Seperated Values
      l_TypeInfo = "tsv" & vbTab & "text/tab-separated-values"
   Case 74
      ' TrueType Font
      l_TypeInfo = "ttf" & vbTab & "application/x-font-ttf"
   Case 75
      ' Text File
      l_TypeInfo = "txt" & vbTab & "text/plain"
   Case 76
      ' vCard
      l_TypeInfo = "vcf" & vbTab & "text/x-vcard"
   Case 77
      ' vCalendar
      l_TypeInfo = "vcs" & vbTab & "text/x-vcalendar"
   Case 78
      ' Microsoft Visio
      l_TypeInfo = "vsd" & vbTab & "application/vnd.visio"
   Case 79
      ' Microsoft Visio 2013
      l_TypeInfo = "vsdx" & vbTab & "application/vnd.visio2013"
   Case 80
      ' Waveform Audio File Format (WAV)
      l_TypeInfo = "wav" & vbTab & "audio/x-wav"
   Case 81
      ' Open Web Media Project - Audio
      l_TypeInfo = "weba" & vbTab & "audio/webm"
   Case 82
      ' Open Web Media Project - Video
      l_TypeInfo = "webm" & vbTab & "video/webm"
   Case 83
      ' WebP Image
      l_TypeInfo = "webp" & vbTab & "image/webp"
   Case 84
      ' Microsoft Windows Media
      l_TypeInfo = "wm" & vbTab & "video/x-ms-wm"
   Case 85
      ' Microsoft Windows Media Audio
      l_TypeInfo = "wma" & vbTab & "audio/x-ms-wma"
   Case 86
      ' Microsoft Windows Metafile
      l_TypeInfo = "wmf" & vbTab & "application/x-msmetafile"
   Case 87
      ' Microsoft Windows Media Video
      l_TypeInfo = "wmv" & vbTab & "video/x-ms-wmv"
   Case 88
      ' Web Open Font Format
      l_TypeInfo = "woff" & vbTab & "application/x-font-woff"
   Case 89
      ' XHTML - The Extensible HyperText Markup Language
      l_TypeInfo = "xhtml" & vbTab & "application/xhtml+xml"
   Case 90
      ' Microsoft Excel
      l_TypeInfo = "xls" & vbTab & "application/vnd.ms-excel"
   Case 91
      ' Microsoft Office OOXML - Spreadsheet
      l_TypeInfo = "xlsx" & vbTab & "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
   Case 92
      ' XML - Extensible Markup Language
      l_TypeInfo = "xml" & vbTab & "application/xml"
   Case 93
      ' Microsoft XML Paper Specification
      l_TypeInfo = "xps" & vbTab & "application/vnd.ms-xpsdocument"
   Case 94
      ' XML Transformations
      l_TypeInfo = "xslt" & vbTab & "application/xslt+xml"
   Case 95
      ' XUL - XML User Interface Language
      l_TypeInfo = "xul" & vbTab & "application/vnd.mozilla.xul+xml"
   Case 96
      ' Zip Archive
      l_TypeInfo = "zip" & vbTab & "application/zip"
   Case Else
      ' Reached end of mime type info list by index. Return empty string to signal caller that we are done.
   End Select
   
   mimeTypeInfoByIndex = l_TypeInfo
End Function

