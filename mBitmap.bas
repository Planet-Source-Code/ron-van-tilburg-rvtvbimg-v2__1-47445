Attribute VB_Name = "mBitmap"
Option Explicit

'- ©2001/2003 Ron van Tilburg - All rights reserved  1.01.2001/Jun-Jul 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'mBitMap.bas
' API Functions for capturing/setting memory dc and bitmaps

Public Type Bitmap '14 bytes
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Public Type BITMAPFILEHEADER
  bfType As Integer
  bfSize As Long
  bfReserved1 As Integer
  bfReserved2 As Integer
  bfOffBits As Long
End Type

Public Type BITMAPINFOHEADER '40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Public Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Public Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER        '40&
  bmiColors(1 To 256) As RGBQUAD       '256&*4&=1024
End Type

Public Type PALETTEENTRY
  peRed As Byte
  peGreen As Byte
  peBlue As Byte
  peFlags As Byte
End Type

Public Type LOGPALETTE
  palVersion As Integer
  palNumEntries As Integer
  palPalEntry(0 To 255) As PALETTEENTRY
End Type

'Used in Creating a StdPicture
Public Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Public Type PictDesc
  size As Long
  Type As Long
  hBmp As Long
  hPal As Long
End Type

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PictDesc, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Public Declare Function GetObjectX Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Const DIB_RGB_COLORS As Long = 0         '  color table in RGBs

Public Const BI_RGB As Long = 0&
Public Const BI_RLE8 As Long = 1&
Public Const BI_RLE4 As Long = 2&
Public Const BI_BITFIELDS As Long = 3&
Public Const BI_JPEG As Long = 4&         'A GUESS
Public Const BI_PNG As Long = 5&          'A GUESS

Public Const SRCCOPY As Long = &HCC0020         ' (DWORD) dest = source
Public Const NOTSRCCOPY As Long = &H330008      ' (DWORD) dest = (NOT source)

Public Const COLORONCOLOR = 3
Public Const BLACKONWHITE = 1
Public Const HALFTONE As Long = 4

Public Const GDI_ERROR As Long = &HFFFF

Public Type POINTAPI
  x As Long
  y As Long
End Type

'Allocate a DC and Bitmap area compatible with the current display - they are NOT CONNECTED YET
Private Function AllocateDCandBitmap(ByVal Width As Long, _
                                     ByVal Height As Long, _
                                     ByRef NewMemDC As Long, _
                                     ByRef NewBitMap As Long) As Boolean

 Dim rc As Long, hDC As Long

  NewMemDC = 0
  NewBitMap = 0
  hDC = CreateDC("DISPLAY", "", "", 0)     'get a DisplayDC
  If hDC <> 0 Then
    NewMemDC = CreateCompatibleDC(hDC)
    If NewMemDC <> 0 Then
      NewBitMap = CreateCompatibleBitmap(hDC, Width, Height)
      If NewBitMap <> 0 Then                                  'we have a suitable bitmap ready for populating
        Call DeleteDC(hDC)                                    'get rid of this we dont need it anymore
        rc = 1
      End If
    End If
  End If
  If rc = 0 Then Call FreeDCandBitmap(NewMemDC, NewBitMap)
  AllocateDCandBitmap = (rc <> 0)

End Function

'Clone a DC and Bitmap area compatible with the given hDC - they are NOT CONNECTED YET
Private Function CloneDCandBitmap(ByVal hDC As Long, _
                                  ByVal Width As Long, _
                                  ByVal Height As Long, _
                                  ByRef NewMemDC As Long, _
                                  ByRef NewBitMap As Long) As Boolean

 Dim rc As Long

  NewMemDC = 0
  NewBitMap = 0
  NewMemDC = CreateCompatibleDC(hDC)
  If NewMemDC <> 0 Then
    NewBitMap = CreateCompatibleBitmap(hDC, Width, Height)
    If NewBitMap <> 0 Then               'we have a suitable bitmap ready for populating
      rc = 1
    End If
  End If
  If rc = 0 Then Call FreeDCandBitmap(NewMemDC, NewBitMap)
  CloneDCandBitmap = (rc <> 0)

End Function

'Free the memory allocated in DC and BitMap - they should not be connected
'WARNING: Dont try this with anything NOT allocated with AllocateDCandBitmap(), or CloneDCandBitmap

Public Sub FreeDCandBitmap(ByRef hDC As Long, ByRef hBitmap As Long)

  If hBitmap <> 0 Then
    Call DeleteObject(hBitmap)
    hBitmap = 0
  End If
  
  If hDC <> 0 Then
    Call DeleteDC(hDC)
    hDC = 0
  End If
End Sub

'=============================================================================================================
'================= This fills a GDI hDC and Bitmap from the data in PixBits(), is always BOTTOM UP
'=============================================================================================================
Public Function DCBitMapFromImage(ByRef PicState As Long, _
                                  ByVal Width As Long, _
                                  ByVal Height As Long, _
                                  ByRef PixBits() As Byte, _
                                  ByVal BitsPerPixel As Long, _
                                  ByRef CMap() As RGBA, _
                                  ByVal NCMapColors As Long, _
                                  ByRef DestDC As Long, _
                                  ByRef DestBitMap As Long, _
                                  Optional OpCode As Long = 0, _
                                  Optional OpParm1 As Long = 0, _
                                  Optional OpParm2 As Long = 0) As Boolean

 Dim BMI  As BITMAPINFO
 Dim rcOK As Boolean, i As Long, p As Long, RasterOp As Long, lppt As POINTAPI
 Dim NewWidth As Long, NewHeight As Long, SrcX As Long, SrcY As Long

  On Error GoTo ErrorFound

  NewWidth = Width: If OpParm1 <> 0 Then NewWidth = OpParm1
  NewHeight = Abs(Height): If OpParm2 <> 0 Then NewHeight = OpParm2
  
  If OpCode = PIC_INVERT_COLOR Then RasterOp = NOTSRCCOPY Else RasterOp = SRCCOPY

  '--------------------------------------------------------------------------------------------------------
  '------------------------------------ BitMap CREATION FROM STORAGE --------------------------------------
  '--------------------------------------------------------------------------------------------------------

  rcOK = AllocateDCandBitmap(Abs(NewWidth), Abs(NewHeight), DestDC, DestBitMap)
  If rcOK Then
    With BMI.bmiHeader
      .biSize = 40                     'sizeof(BITMAPINFOHEADER
      .biWidth = Width
      .biHeight = -Height               'assume its bottom to top, if negative otherway up
      .biPlanes = 1
      .biBitCount = BitsPerPixel       '{desired color resolution (1, 4, 8, or 16,24)}
      .biClrUsed = NCMapColors
      .biCompression = BI_RGB          '16BIT 5,6,5 will be Wrong in Display
      .biSizeImage = BMPRowModulo(Width, BitsPerPixel) * Abs(Height)
    End With

    If BitsPerPixel <= PIC_8BPP Then      'we have a Colormap to use
      For i = 1 To NCMapColors
        p = i - 1
        With BMI.bmiColors(i)
          .rgbRed = CMap(p).Red
          .rgbGreen = CMap(p).Green
          .rgbBlue = CMap(p).Blue
          .rgbReserved = CMap(p).Alpha
        End With
      Next i
    End If
  
    If OpCode = PIC_CLIP_CENTRED Then  'we wont stretch but rather we will cut
      SrcX = (Width - NewWidth) \ 2: If SrcX < 0 Then SrcX = 0
      SrcY = (Abs(Height) - NewHeight) \ 2: If SrcY < 0 Then SrcY = 0
      Width = NewWidth
      Height = NewHeight
    End If
    
    DestBitMap = SelectObject(DestDC, DestBitMap)   'select it in
    Call SetStretchBltMode(DestDC, COLORONCOLOR)  'HALFTONE)
    Call SetBrushOrgEx(DestDC, 0, 0, lppt)
    
    rcOK = (StretchDIBits(DestDC, 0, 0, Abs(NewWidth), Abs(NewHeight), _
                                  SrcX, SrcY, Width, Abs(Height), _
                                  PixBits(0), BMI, DIB_RGB_COLORS, RasterOp) <> GDI_ERROR)
    DestBitMap = SelectObject(DestDC, DestBitMap)   'unselect it again
  End If

  DCBitMapFromImage = rcOK

ErrorFound:
  On Error GoTo 0

End Function                          'the values NewDC and NewBitMap will contain the Image

'This function manipulates the image using the API functions (quite quick)
'Supports: PIC_UNMAP_COLOR, PIC_MSMAP_COLOR, PIC_IMAGE_RESIZE, PIC_INVERT_COLOR

Public Function APIOperations(ByRef PicState As Long, _
                              ByRef Width As Long, _
                              ByRef Height As Long, _
                              ByRef PixBits() As Byte, _
                              ByRef BPP As Long, _
                              ByRef CMap() As RGBA, _
                              ByRef NCMapColors As Long, _
                              ByVal OpCode As Long, _
                              Optional ByVal OpParm1 As Long = 0, _
                              Optional ByVal OpParm2 As Long = 0) As Boolean

 Dim hDC As Long, hBitmap As Long, rc As Long
 Dim NewWidth As Long, NewHeight As Long, NewBPP As Long

  NewWidth = Width: NewHeight = Height: NewBPP = BPP    'assume nothing changes

  If OpCode = PIC_RESIZE Then
    If OpParm1 > 0 Then NewWidth = OpParm1
    If OpParm2 > 0 Then NewHeight = OpParm2
  End If

  If NewWidth < 1 Then NewWidth = Width     'if <=zero revert
  If NewHeight < 1 Then NewHeight = Height  'if <=zero revert

  If OpCode = PIC_UNMAP_COLOR Then
    NewBPP = PIC_24BPP

  ElseIf OpCode = PIC_MSMAP_COLOR Then
    If OpParm1 <= PIC_1BPP Then
      OpParm1 = PIC_1BPP
     ElseIf OpParm1 <= PIC_4BPP Then
      OpParm1 = PIC_4BPP
     ElseIf OpParm1 <= PIC_8BPP Then
      OpParm1 = PIC_8BPP
     ElseIf OpParm1 <> PIC_16BPP _
        And OpParm1 <> PIC_24BPP _
        And OpParm1 <> PIC_32BPP Then
      OpParm1 = PIC_24BPP
    End If
    NewBPP = OpParm1                       'BE careful here only 1,4,8,16,24,32 OK
  End If

  'a new DC and Bitmap are allocated in this call
  rc = DCBitMapFromImage(PicState, Width, Height, PixBits(), BPP, CMap(), NCMapColors, _
                         hDC, hBitmap, OpCode, NewWidth, NewHeight)

  If rc <> 0 Then
    rc = ImageFromDCBitMap(hDC, hBitmap, PicState, Abs(NewWidth), Abs(NewHeight), _
                                         PixBits(), NewBPP, CMap(), NCMapColors)
    If rc <> 0 Then
      Width = Abs(NewWidth)
      Height = Abs(NewHeight)
      BPP = NewBPP
    End If
  End If

  Call FreeDCandBitmap(hDC, hBitmap)
  APIOperations = (rc <> 0)

End Function

'=============================================================================================================
'================= This fills PixBits() from a GDI hDC and Bitmap. It is always TOP DOWN
'=============================================================================================================
Public Function ImageFromDCBitMap(ByVal SrcDC As Long, _
                                  ByVal SrcBitmap As Long, _
                                  ByRef PicState As Long, _
                                  ByVal Width As Long, _
                                  ByVal Height As Long, _
                                  ByRef PixBits() As Byte, _
                                  ByVal BitsPerPixel As Long, _
                                  ByRef CMap() As RGBA, _
                                  ByRef NCMapColors As Long) As Boolean

 Dim BMI As BITMAPINFO
 Dim rc  As Long, i As Long, p As Long

  rc = 0
  On Error GoTo ErrorFound

  '--------------------------------------------------------------------------------------------------------
  '-----------------------------------  BitMap CAPTURE INTO STORAGE ---------------------------------------
  '--------------------------------------------------------------------------------------------------------
  If BitsPerPixel >= PIC_16BPP Then
    NCMapColors = 0
  Else
    NCMapColors = 2& ^ BitsPerPixel
  End If
  
  With BMI.bmiHeader
    .biSize = 40                              'sizeof(BITMAPINFOHEADER
    .biWidth = Width                          '{width of the bitmapclip}
    .biHeight = -Height                       '{height of the bitmapclip} make sure its top to bottom
    .biPlanes = 1
    .biBitCount = BitsPerPixel                '{desired color resolution (1, 4, 8, or 24)}
    .biClrUsed = NCMapColors
    .biCompression = BI_RGB
    .biSizeImage = BMPRowModulo(Width, BitsPerPixel) * Abs(Height)
  End With

  ReDim PixBits(0 To BMI.bmiHeader.biSizeImage - 1) As Byte       'the real image size

  rc = GetDIBits(SrcDC, SrcBitmap, 0, Abs(Height), PixBits(0), BMI, DIB_RGB_COLORS)

  If rc <> 0 Then                             'we now have the whole thing captured
    If BitsPerPixel < PIC_16BPP Then          'we have a Colormap to use
      ReDim CMap(0 To NCMapColors - 1)
      For i = 1 To NCMapColors
        p = i - 1
        With BMI.bmiColors(i)
          CMap(p).Red = .rgbRed
          CMap(p).Green = .rgbGreen
          CMap(p).Blue = .rgbBlue
          CMap(p).Alpha = .rgbReserved
        End With
      Next i
    End If
  End If
  Call SetF(PicState, IS_TOP_TO_BOTTOM)
  
ErrorFound:
  ImageFromDCBitMap = (rc <> 0)
  On Error GoTo 0

End Function

'=============================================================================================================
'Get an Image from an Object with a DC, optionally clipping a part of it
Public Function ImagefromObjhDC(ByVal Obj As Object, _
                                ByVal PicType As Long, _
                                ByRef PicState As Long, _
                                ByRef Width As Long, _
                                ByRef Height As Long, _
                                ByRef PixBits() As Byte, _
                                ByRef IntBPP As Long, _
                                ByRef CMap() As RGBA, _
                                ByRef NCMapColors As Long, _
                                Optional ByVal ClipLeft As Long = -1, _
                                Optional ByVal ClipTop As Long = -1, _
                                Optional ByVal ClipRight As Long = -1, _
                                Optional ByVal ClipBottom As Long = -1) As Boolean
  
 Dim OldScaleMode As ScaleModeConstants
 
  On Error GoTo CaptureFailed
  
  OldScaleMode = Obj.ScaleMode                      'no Scalemode is an error
  Obj.ScaleMode = vbPixels                          'we depend on this throughout
  If Obj.hDC <> 0 Then                              'will fail if No hDC or Empty
    Call ValidateClip(ClipLeft, ClipTop, ClipRight, ClipBottom, _
                      0, 0, Obj.ScaleWidth, Obj.ScaleHeight, _
                      Width, Height)
    ImagefromObjhDC = ImageFromDCClip(Obj.hDC, PicType, PicState, _
                                               Width, Height, PixBits(), IntBPP, CMap(), NCMapColors, _
                                               ClipLeft, ClipTop, ClipRight, ClipBottom)
  End If
  
CaptureFailed:
  On Error Resume Next
  Obj.ScaleMode = OldScaleMode                    'we depend on this throughout
  On Error GoTo 0
  
End Function

'--------------------------------------------------------------------------------------------------------
'================= This takes a GDI hDC, makes a copy of the bitmap and saves it to DIBits() =======
'--------------------------------------------------------------------------------------------------------
'any gDI Device Context should work here from eg. Form, PictureBox, Image, Control...
'Will not work with Printer.object (someone explain why to me please)
'It can fail for lack of memory   0=failure,1=success, NO error checking on other inputs (shielded by Class)
'--------------------------------------------------------------------------------------------------------

Public Function ImageFromDCClip(hDC As Long, _
                                ByVal PicType As Long, _
                                ByRef PicState As Long, _
                                ByRef Width As Long, _
                                ByRef Height As Long, _
                                ByRef PixBits() As Byte, _
                                ByVal IntBPP As Long, _
                                ByRef CMap() As RGBA, _
                                ByRef NCMapColors As Long, _
                                ByVal ClipLeft As Long, _
                                ByVal ClipTop As Long, _
                                ByVal ClipRight As Long, _
                                ByVal ClipBottom As Long) As Boolean

 Dim hwBitMap As Long, hwMemDC As Long         'BITMAP,hDC2
 Dim rcOK As Boolean, i As Long, p As Long
 Dim ClipWidth As Long, ClipHeight As Long

  On Error GoTo ErrorFound
  '--------------------------------------  PARAMETER VALIDATION  ------------------------------------------

  PicState = IS_VALID_PIPELINE    'clear all other bits

  'NOTE: no validation of clipping rectangle in this backend routine - MAKE SURE ITS RIGHT
  ClipWidth = ClipRight - ClipLeft
  ClipHeight = ClipBottom - ClipTop
  If ClipWidth < 1 Or ClipHeight < 1 Then GoTo ErrorFound     'Que??? its a bit small or overlapped

  Call SetF(PicState, IS_VALID_CLIP)

  '--------------------------------------------------------------------------------------------------------
  '--------------------------------------  DIB CAPTURE INTO STORAGE ---------------------------------------
  '--------------------------------------------------------------------------------------------------------
  rcOK = CloneDCandBitmap(hDC, ClipWidth, ClipHeight, hwMemDC, hwBitMap)
  If rcOK Then
    hwBitMap = SelectObject(hwMemDC, hwBitMap)    'we get a clip of the right size to fill the cloned map
    Call BitBlt(hwMemDC, 0, 0, ClipWidth, ClipHeight, hDC, ClipLeft, ClipTop, SRCCOPY)
    hwBitMap = SelectObject(hwMemDC, hwBitMap)    'hwBitMap is now a clipped copy, and unselected

'    If PicType <> PIC_BMP Then
'      rcOK = ImageFromDCBitMap(hwMemDC, hwBitMap, PicState, ClipWidth, -ClipHeight, PixBits(), IntBPP, CMap(), NCMapColors)
'      If rcOK Then Call SetF(PicState, IS_TOP_TO_BOTTOM)  'mark it right way up
'    Else
      rcOK = ImageFromDCBitMap(hwMemDC, hwBitMap, PicState, ClipWidth, ClipHeight, PixBits(), IntBPP, CMap(), NCMapColors)
'    End If

    If rcOK Then
      Width = ClipWidth
      Height = ClipHeight
      If NCMapColors <> 0 Then Call SetF(PicState, IS_CMAPPED)
    End If
  End If

ErrorFound:
  Call FreeDCandBitmap(hwMemDC, hwBitMap)
  ImageFromDCClip = rcOK
  On Error GoTo 0

End Function
  
'Capture Data info from a Passed in BitMap
Public Function CaptureBitmap(ByVal hBM As Long, _
                              ByRef PicState As Long, _
                              ByRef Width As Long, _
                              ByRef Height As Long, _
                              ByRef PixBits() As Byte, _
                              ByRef IntBPP As Long, _
                              ByRef CMap() As RGBA, _
                              ByRef NCMapColors As Long) As Boolean
  
  Dim tmpDC As Long, BMI As Bitmap
  
  If hBM <> 0 Then
    PicState = IS_VALID_PIPELINE
    Call GetObject(hBM, Len(BMI), BMI)
    With BMI
      Width = .bmWidth
      Height = .bmHeight
      If Height < 0 Then
        Height = -Height
        Call SetF(PicState, IS_TOP_TO_BOTTOM)
      End If
      IntBPP = .bmBitsPixel
    End With
    tmpDC = CreateDC("DISPLAY", "", "", 0)       'get a DisplayDC
    If tmpDC <> 0 Then
      IntBPP = PIC_24BPP                         'always capture at 24BPP
      CaptureBitmap = ImageFromDCBitMap(tmpDC, hBM, PicState, Width, Height, PixBits(), IntBPP, CMap(), NCMapColors)
      Call DeleteObject(tmpDC)
    End If
  End If
  
End Function

'Renders the current Image to the Object.hdc
Public Function RenderImage(ByVal Obj As Object, _
                            ByVal PicState As Long, _
                            ByVal Width As Long, _
                            ByVal Height As Long, _
                            ByRef PixBits() As Byte, _
                            ByVal IntBPP As Long, _
                            ByRef CMap() As RGBA, _
                            ByVal NCMapColors As Long, _
                            ByVal RenderOptions As IMG_RENDEROPTIONS, _
                            ByVal ClipLeft As Long, _
                            ByVal ClipTop As Long, _
                            ByVal ClipRight As Long, _
                            ByVal ClipBottom As Long) As Boolean  'TRUE IS GOOD

 Dim ObjhDC As Long, w As Long, H As Long, z As Long
 Dim oldBM As Long, tmpDC As Long, tmpBM As Long
 Dim OldScaleMode As ScaleModeConstants

 Dim xc As Long, yc As Long, cW As Long, ch As Long
 Dim xd As Long, yd As Long, dw As Long, dh As Long
 Dim zi As Long, zj As Long, i As Long, j As Long

  On Error GoTo RenderFailed
  
  OldScaleMode = Obj.ScaleMode                      'no Scalemode is an error
  Obj.ScaleMode = vbPixels                          'we depend on this throughout
  If Obj.hDC = 0 Then GoTo RenderFailed             'will fail if No hDC or Empty
  ObjhDC = Obj.hDC
  
  'Get the Destination Rectangle (assume all of hdc)
  xd = 0
  yd = 0
  dw = Obj.ScaleWidth
  dh = Obj.ScaleHeight
  
  'Get The Source Rectangle
  Call ValidateClip(ClipLeft, ClipTop, ClipRight, ClipBottom, 0, 0, Width, Height, cW, ch)
  xc = ClipLeft
  yc = ClipTop
  
  'If the Image is larger than the target get a clipped Bitmap else the whole thing
  If dw < Width And dh < Height Then
    Call DCBitMapFromImage(PicState, Width, Height, PixBits(), IntBPP, _
                           CMap(), NCMapColors, tmpDC, tmpBM, PIC_CLIP_CENTRED, dw, dh)
    cW = dw: ch = dh                                     'clip is this big now
  Else
    Call DCBitMapFromImage(PicState, Width, Height, PixBits(), IntBPP, CMap(), NCMapColors, tmpDC, tmpBM)
  End If
  
  If tmpDC = 0 Or tmpBM = 0 Then Exit Function           'Image probably too large, give up

  'Render according to Flags
  If RenderOptions = IRO_ASIS Then                       'Put at CurrXY, and Without scaling
    dw = cW
    dh = ch
  ElseIf FSet(RenderOptions, IRO_STRETCH) Then           'Stretch to Fit Clip Area (ie AutoCentred)
    w = dw
    H = dh
    Select Case MaskF(RenderOptions, IRO_KEEPASPECT Or IRO_INTSCALE)
     Case 0
      'do nothing extra
      
     Case IRO_KEEPASPECT
      'Image Aspect Ratio is Kept during Stretch wont necessarily fit to edges)
      If cW >= ch Then
        w = CSng(cW) * CSng(dh) / CSng(ch)
        H = CSng(ch) * CSng(w) / CSng(cW)
       ElseIf ch > cW Then
        H = CSng(ch) * CSng(dw) / CSng(cW)
        w = CSng(cW) * CSng(H) / CSng(ch)
      End If

     Case IRO_INTSCALE, IRO_KEEPASPECT Or IRO_INTSCALE
      'Image is scaled by nearest integer to StretchFit (wont necessarily fit to edges)
      If dw > cW Then
        z = dw \ cW
        If z = 0 Then z = 1
        w = cW * z
        If w > dw Then w = dw
       ElseIf cW > dw Then
        z = cW \ dw
        If z = 0 Then z = 1
        w = dw \ z
      End If
      If dh > ch Then
        z = dh \ ch
        If z = 0 Then z = 1
        H = ch * z
        If H > dh Then H = dh
       ElseIf ch > dh Then
        z = ch \ dh
        If z = 0 Then z = 1
        H = dh \ z
      End If
    End Select

    If FSet(RenderOptions, IRO_CENTRE) Then
      xd = xd + (dw - w) \ 2
      yd = yd + (dh - H) \ 2
    End If
    dw = w
    dh = H
  End If
   
  oldBM = SelectObject(tmpDC, tmpBM)
  If FSet(RenderOptions, IRO_TILE) Then                'Tile Image to Fit Clip Area (image is not Scaled)

    'We tile the image onto the Objhdc
    'The imageclip is assumed to be aligned to 0,cw,2cw, etc and 0,ch,2ch,3ch etc
    'and the destination is aligned likewise from zRPLeft,zRPTop  etc
    'We BitBlt because we dont have to scale

    If FSet(RenderOptions, IRO_CENTRE) Then   'we align the rect and image centres, then tile
      If cW <= dw Then
        zi = ((dw - cW) \ 2) Mod cW
        If zi > 0 Then zi = cW - zi
       Else
        zi = (cW - dw) \ 2
      End If
      If ch <= dh Then
        zj = ((dh - ch) \ 2) Mod ch
        If zj > 0 Then zj = ch - zj
       Else
        zj = (ch - dh) \ 2
      End If
     Else
      'we may need to further adjust i and j to account for the misalignment in images
      zi = xd Mod cW
      zj = yd Mod ch
    End If

    If cW >= dw And ch >= dh Then    'The image is bigger than the to be tiled region (easy)
      Call BitBlt(ObjhDC, xd, yd, dw, dh, tmpDC, zi, zj, SRCCOPY)
     Else

      If zi <> 0 And zj <> 0 Then    'top,left corner
        Call BitBlt(ObjhDC, xd, yd, cW - zi, ch - zj, tmpDC, zi, zj, SRCCOPY)
      End If

      If zj <> 0 Then               'top strip
        If zi = 0 Then
          For i = xd To xd + dw Step cW
            Call BitBlt(ObjhDC, i, yd, cW, ch - zj, tmpDC, 0, zj, SRCCOPY)
          Next i
         Else
          For i = xd + cW - zi To xd + dw Step cW
            Call BitBlt(ObjhDC, i, yd, cW, ch - zj, tmpDC, 0, zj, SRCCOPY)
          Next i
        End If
      End If

      If zi <> 0 Then               'left strip
        If zj = 0 Then
          For j = yd To yd + dh Step ch
            Call BitBlt(ObjhDC, xd, j, cW - zi, ch, tmpDC, zi, 0, SRCCOPY)
          Next j
         Else
          For j = yd + ch - zj To yd + dh Step ch
            Call BitBlt(ObjhDC, xd, j, cW - zi, ch, tmpDC, zi, 0, SRCCOPY)
          Next j
        End If
      End If

      'The rest (we are always clipping to bottom,right)
      If zi <> 0 Then zi = cW - zi
      If zj <> 0 Then zj = ch - zj
      For i = xd + zi To xd + dw Step cW
        For j = yd + zj To yd + dh Step ch
          Call BitBlt(ObjhDC, i, j, cW, ch, tmpDC, 0, 0, SRCCOPY)
        Next j
      Next i
    End If
   Else
    If cW > dw And ch > dh Then    'The image is bigger than the to be tiled region (easy)
      Call BitBlt(ObjhDC, xd, yd, dw, dh, tmpDC, zi, zj, SRCCOPY)
    Else
      Call StretchBlt(ObjhDC, xd, yd, dw, dh, tmpDC, xc, yc, cW, ch, SRCCOPY)
    End If
  End If
  RenderImage = True
  Call SelectObject(tmpDC, oldBM)
  
RenderFailed:
  Call FreeDCandBitmap(tmpDC, tmpBM)
  On Error Resume Next
    Obj.ScaleMode = OldScaleMode
  On Error Resume Next
    Obj.Refresh
  On Error GoTo 0
End Function

'Validate (modify) a clip against a base area, so that it is valid (but might be empty)
Public Sub ValidateClip(ByRef ClipLeft As Long, ByRef ClipTop As Long, _
                        ByRef ClipRight As Long, ByRef ClipBottom As Long, _
                        ByVal AreaLeft As Long, ByVal AreaTop As Long, _
                        ByVal AreaRight As Long, ByVal AreaBottom As Long, _
                        ByRef ClipWidth As Long, ByRef ClipHeight As Long)
 Dim z As Long

 'fit the clip to the extremes of the given area

  If ClipLeft < AreaLeft Then
    ClipLeft = AreaLeft
  ElseIf ClipLeft > AreaRight Then
    ClipLeft = AreaRight
  End If
  
  If ClipTop < AreaLeft Then
    ClipTop = AreaTop
  ElseIf ClipTop > AreaBottom Then
    ClipTop = AreaBottom
  End If
  
  If ClipRight < AreaLeft Then
    ClipRight = AreaLeft
  ElseIf ClipRight > AreaRight Then
    ClipRight = AreaRight
  End If
  
  If ClipBottom < AreaTop Then
    ClipBottom = AreaTop
  ElseIf ClipBottom > AreaBottom Then
    ClipBottom = AreaBottom
  End If
  
  'if empty make it full size
  If ClipLeft = ClipRight Then
    ClipLeft = AreaLeft
    ClipRight = AreaRight
  End If
  
  If ClipTop = ClipBottom Then
    ClipTop = AreaTop
    ClipBottom = AreaBottom
  End If
  
  'make sure its normalised
  If ClipRight < ClipLeft Then
    z = ClipLeft
    ClipLeft = ClipRight
    ClipRight = z
  End If
  
  If ClipBottom < ClipTop Then
    z = ClipTop
    ClipTop = ClipBottom
    ClipBottom = z
  End If
  
  'and now return the final size of the clip
  ClipWidth = ClipRight - ClipLeft
  ClipHeight = ClipBottom - ClipTop
End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jun-13 22:34) 95 + 328 = 423 Lines
