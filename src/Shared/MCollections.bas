Attribute VB_Name = "MCollections"
Option Explicit

Public Function collectionHasContent(po_Collection As vbRichClient5.cCollection) As Boolean
   If Not po_Collection Is Nothing Then
      collectionHasContent = (po_Collection.Count > 0)
   End If
End Function

Public Function collectionIsJsonArray(po_Collection As vbRichClient5.cCollection) As Boolean
   If Not po_Collection Is Nothing Then
      collectionIsJsonArray = po_Collection.IsJSONArray
   End If
End Function

Public Function collectionIsJsonObject(po_Collection As vbRichClient5.cCollection) As Boolean
   If Not po_Collection Is Nothing Then
      collectionIsJsonObject = po_Collection.IsJSONObject
   End If
End Function

