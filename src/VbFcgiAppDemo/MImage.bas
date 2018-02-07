Attribute VB_Name = "MImage"
Option Explicit

Public Function imageStdPictureToCairoSurface(po_StdPicture As stdole.StdPicture) As vbRichClient5.cCairoSurface
         ' Code by Olaf Schmidt. Found here: http://www.vbforums.com/showthread.php?857371-How-to-use-vbRichClient5-to-convert-VB-Picture-to-a-byte-array&p=5249553&viewfull=1#post5249553
         
         Dim lo_Cairo As vbRichClient5.cCairo
         Dim lo_CC As vbRichClient5.cCairoContext
         Dim lo_Dib As vbRichClient5.cDIB
         
10       On Error GoTo ErrorHandler
         
20       Set lo_Dib = libRc5Factory.C.DIB(, , po_StdPicture)   'create an RC5-DIBObj from the StdPic
         
30       Set lo_Cairo = libRc5Factory.C.Cairo
         
40       Set imageStdPictureToCairoSurface = lo_Cairo.CreateWin32Surface(lo_Dib.dx, lo_Dib.dy) 'ensure the Dest-Surface
         
50       lo_Dib.DrawTo imageStdPictureToCairoSurface.GetDC   'hDC-based Blitting from DIB to CairoSurface
            
         'normally we would be finished here - but the above Blt-Op left out the Alpha-Channel,
60       Set lo_CC = imageStdPictureToCairoSurface.CreateContext
70       lo_CC.Operator = CAIRO_OPERATOR_DEST_ATOP '...so we have to ensure one with a Paint-Op,
80       lo_CC.Paint 1, lo_Cairo.CreateSolidPatternLng(0, 1)   '<- which sets the Alpha-Channel to "fully opaque"

90       Exit Function
         
ErrorHandler:
100      apiOutputDebugString "*** ERROR *** " & Err.Number & " " & Err.Description & ", Line #" & Erl
End Function

