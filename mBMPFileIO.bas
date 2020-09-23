Attribute VB_Name = "mBMPFileIO"

Option Explicit

'- ©2001/2003 Ron van Tilburg - All rights reserved  1.01.2001/Jun-Jul 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'BMPSave.bas
' for loading (unpacking) and saving (packing) bitmaps

'02 Jun 2003 added Default ALPHA Channel as GreyCode for RGB in BMCM Save
'04 Jul 2003 added 32K and 64K mapped reading/writing, all BPPs now supported for BMPs
'            all reads result in a 24, or 32 bpp representation

'==================================== LOADING A BMP ==========================================================
'We load a bitmap (color mapped or otherwise, and return an unmapped, 24 bit Image

'These two arrays are used when rescaling 5 and 6 bit values from BMPs
Public BMP31Scale(0 To 31) As Long
Public BMP63Scale(0 To 63) As Long

Public Sub InitBMPScales()
  Dim i As Long
  
  For i = 0 To 31
    BMP31Scale(i) = (i * 510& + 31&) \ 62& '- ((i * 8&) Or (i And 7&))  '
  Next i
  
  For i = 0 To 63
    BMP63Scale(i) = (i * 510& + 63&) \ 126& '(i * 4&) Or (i And 3&)
  Next i
  
End Sub

Public Function LoadBMP(Path As String, _
                        ByRef PicState As Long, _
                        ByRef Width As Long, ByRef Height As Long, _
                        ByRef PixBits() As Byte, ByRef IntBPP As Long, _
                        ByRef CMap() As RGBA, ByRef NCMapColors As Long) As Boolean

 Dim DIBS As Long, Bitmap As Long
 Dim BMFH As BITMAPFILEHEADER
 Dim BMIH As BITMAPINFOHEADER
 Dim BMCM() As RGBQUAD
 Dim i As Long, p As Long, FileNr As Long, RMask As Long, GMask As Long, BMask As Long

  On Error GoTo ReadError
    
  FileNr = FreeFile()

  'I assume that we only have a 'Modern' bitmap ie BMIH.bisize=40, so that we have QUADRGB colors

  Open Path For Binary Access Read As #FileNr
  Get #FileNr, , BMFH
  If BMFH.bfType <> &H4D42 Then GoTo ReadError 'not a BMP
  
  PicState = IS_VALID_CLIP
  Get #FileNr, , BMIH
  With BMIH
    Width = .biWidth
    Height = .biHeight
    If Height < 0 Then
      Call SetF(PicState, IS_TOP_TO_BOTTOM)
      Height = -Height
    Else
      Call ClrF(PicState, IS_TOP_TO_BOTTOM)
    End If
    IntBPP = .biBitCount
    
    Select Case .biCompression  'See what we can Deal with
      Case BI_RLE4, BI_RLE8, BI_JPEG, BI_PNG:
        GoTo ReadError                      'not Supported (so Far)
    
      Case BI_BITFIELDS:
        Get #FileNr, , RMask
        Get #FileNr, , GMask
        Get #FileNr, , BMask
        
        If .biClrImportant <> 0 Then         'read the important colors (but ignore them)
          ReDim CMap(0 To .biClrImportant - 1)
          Get #FileNr, , CMap()
        End If
        
      Case BI_RGB:
        If IntBPP = 16 Then           'Default
          RMask = &H7C00&
          GMask = &H3E0&
          BMask = &H1F&
        ElseIf IntBPP = 32 Then       'Default
          RMask = &HFF0000
          GMask = &HFF00&
          BMask = &HFF&
        End If
    End Select
  End With

  Select Case IntBPP
    Case 1, 4, 8:
      NCMapColors = 2 ^ IntBPP
      ReDim CMap(0 To NCMapColors - 1)
      ReDim BMCM(0 To NCMapColors - 1)
      Get #FileNr, , BMCM()
      For i = 0 To NCMapColors - 1
        With CMap(i)
          .Red = BMCM(i).rgbRed
          .Green = BMCM(i).rgbGreen
          .Blue = BMCM(i).rgbBlue
          .Alpha = 0
        End With
      Next i
      
    Case 16:
      NCMapColors = 0  'do nothing, we assume this is not mapped
    
    Case 24:
      NCMapColors = 0  'do nothing, we assume this is not mapped
    
    Case 32:
      NCMapColors = 0  'do nothing, we assume this is not mapped

  End Select

  'get the bitdata
  ReDim PixBits(0 To BMPRowModulo(Width, IntBPP) * Height - 1)
  Seek #FileNr, 1 + BMFH.bfOffBits
  Get #FileNr, , PixBits()
  Close #FileNr
  LoadBMP = True
  
  On Error GoTo 0

  If IntBPP < 24 Then
    LoadBMP = UnPackBMPImage(Width, Height, PixBits(), IntBPP, CMap(), NCMapColors, RMask, GMask, BMask)
  End If
  
  'now always in 24 or 32 bpp mode
Exit Function

ReadError:
  On Error Resume Next
    Close #FileNr
    Erase PixBits(), CMap()
    NCMapColors = 0
    Width = 0
    Height = 0
    IntBPP = 0
  On Error GoTo 0

End Function

'==================================== SAVING A BMP ==============================================================

Public Function SaveBMP(Path As String, ByVal PicState As Long, _
                        ByVal Width As Long, ByVal Height As Long, _
                        ByRef PixBits() As Byte, ByRef IntBPP As Long, ByVal ReqBPP As Long, _
                        ByRef CMap() As RGBA, ByVal NCMapColors As Long) As Boolean

 Dim BMFH As BITMAPFILEHEADER
 Dim BMIH As BITMAPINFOHEADER
 Dim BMCM() As RGBQUAD
 Dim i As Long, p As Long, FileNr As Long, RMask As Long, GMask As Long, BMask As Long

  On Error GoTo WriteError
      
  Select Case ReqBPP
    Case PIC_1BPP:
      ReDim Preserve CMap(0 To 1)
      NCMapColors = 2
      Call PackBMPImage(Width, Height, PixBits(), IntBPP, ReqBPP)
    
    Case PIC_2BPP, PIC_3BPP, PIC_4BPP:
      ReDim Preserve CMap(0 To 15)
      NCMapColors = 16
      Call PackBMPImage(Width, Height, PixBits(), IntBPP, ReqBPP)
      
    Case PIC_5BPP, PIC_6BPP, PIC_7BPP, PIC_8BPP:
      ReDim Preserve CMap(0 To 255)
      NCMapColors = 256
      
    Case PIC_9BPP, PIC_12BPP, PIC_15BPP:   'already packed
      RMask = &H7C00&
      GMask = &H3E0&
      BMask = &H1F&
      
    Case PIC_16BPP:     'already packed
      RMask = &HF800&
      GMask = &H7E0&
      BMask = &H1F&
      
    Case PIC_24BPP:
      If IntBPP = 32 Then Call BMPImage32to24BPP(Width, Height, PixBits, IntBPP)
      
    Case PIC_32BPP:
      If IntBPP = 24 Then Call BMPImage24to32BPP(Width, Height, PixBits, IntBPP)
  End Select
  
  FileNr = FreeFile()
  With BMIH
    .biSize = 40                              'sizeof(BITMAPINFOHEADER
    .biWidth = Width                          '{width of the bitmapclip}
    If FSet(PicState, IS_TOP_TO_BOTTOM) Then
      .biHeight = -Height
    Else
      .biHeight = Height                        '{height of the bitmapclip} make sure its bottom to top
    End If
    .biPlanes = 1
    .biBitCount = IntBPP                      '{desired color resolution (1, 4, 8, 16, 24 or 32)}
    .biClrUsed = NCMapColors                  '2,16,256,0,0,0
    If ReqBPP >= PIC_9BPP And ReqBPP <= PIC_16BPP Then
      .biCompression = BI_BITFIELDS
    Else
      .biCompression = BI_RGB
    End If
    .biSizeImage = BMPRowModulo(Width, IntBPP) * Height
  End With

  With BMFH
    .bfType = &H4D42                                                          'as integer                   '
    .bfSize = 14& + 40& + 4& * NCMapColors + BMIH.biSizeImage                 'sizeof file  As Long
    .bfReserved1 = 0                                                          'As Integer
    .bfReserved2 = 0                                                          'As Integer
    .bfOffBits = 14& + 40& + 4& * NCMapColors                                 'As Long
    
    If ReqBPP >= PIC_9BPP And ReqBPP <= PIC_16BPP Then    'allow for masks
      .bfSize = .bfSize + 12
      .bfOffBits = .bfOffBits + 12
    End If
  End With
  
  Open Path For Binary Access Write As #FileNr
  Put #FileNr, , BMFH
  Put #FileNr, , BMIH
  If IntBPP <= PIC_8BPP Then    'write out the colormap
    ReDim BMCM(1 To NCMapColors)
    For i = 1 To NCMapColors
      p = i - 1
      With BMCM(i)
        .rgbRed = CMap(p).Red
        .rgbGreen = CMap(p).Green
        .rgbBlue = CMap(p).Blue
        .rgbReserved = RGBtoGrey(CMap(p).Red, CMap(p).Green, CMap(p).Blue)         'Default ALPHA
      End With
    Next i
    Put #FileNr, , BMCM()
  ElseIf IntBPP = PIC_16BPP Then
    Put #FileNr, , RMask
    Put #FileNr, , GMask
    Put #FileNr, , BMask
  End If
  Put #FileNr, , PixBits()
  Close #FileNr
  SaveBMP = True
  
  'revert to 24BPP format after Saving for 16BPP 5,6,5 formats (in case someone wants to display it)
  If ReqBPP = PIC_16BPP Then
    Call UnPackBMPImage(Width, Height, PixBits(), IntBPP, CMap(), NCMapColors, RMask, GMask, BMask)
  End If
  
WriteError:
  On Error GoTo 0

End Function

'=============================================================================================================
'This is the correct RowModulo for a BMP file

Public Function BMPRowModulo(ByVal Width As Long, ByVal BitsPerPixel As Long) As Long

  BMPRowModulo = (((Width * BitsPerPixel) + 31&) And Not 31&) \ 8&

End Function

'This routine takes an image in 8BPP ColorMapped format which needs to be packed into 1BPP or 4BPP format
Public Function PackBMPImage(ByVal Width As Long, ByVal Height As Long, _
                             ByRef PixBits() As Byte, _
                             ByRef IntBPP As Long, _
                             ByVal ReqBPP As Long) As Boolean  'deals with all MS formats

 Dim x As Long, y As Long, z As Long, w As Long, maskx As Long
 Dim RowMod As Long, NewRowMod As Long, p As Long, i As Integer, c As Integer

  If IntBPP <> PIC_8BPP Then Exit Function

  ' Scan the PixBits() and pack all pixels in each row of the matrix
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)  'the byte width of a given row
  If ReqBPP = PIC_1BPP Then
    NewRowMod = BMPRowModulo(Width, PIC_1BPP)                      'the byte width of a row
  Else
    NewRowMod = BMPRowModulo(Width, PIC_4BPP)                      'the byte width of a row
  End If
  
  For y = 0 To Height - 1                                      'assume bmap is right way up
    x = y * RowMod                                             'this is the byte where we will put the result
    p = y * NewRowMod
    i = 0
    Select Case ReqBPP
     Case PIC_1BPP:   '8 pixels per byte
      maskx = 128: c = 0
      For i = 0 To Width - 1
        If PixBits(x) <> 0 Then c = c Or maskx
        maskx = maskx \ 2
        If maskx = 0 Then
          PixBits(p) = c
          p = p + 1
          maskx = 128
          c = 0
        End If
        x = x + 1
      Next i
      If maskx <> 128 Then PixBits(p) = c         'the remaining bits

     Case PIC_2BPP, PIC_3BPP, PIC_4BPP: '2 pixels per byte
      w = 0
      For i = 0 To Width - 1
        If (w And 1) = 0 Then
          PixBits(p) = PixBits(x) * 16          'this guarantees it being filled in on oddlength rows
          w = w + 1
         Else
          PixBits(p) = PixBits(p) Or PixBits(x)
          w = w + 1
          p = p + 1
        End If
        x = x + 1
      Next i
    End Select
  Next y

  If ReqBPP = PIC_1BPP Then
    IntBPP = ReqBPP
  Else
    IntBPP = PIC_4BPP
  End If
  
  '  MsgBox "OK PackBMPImage"
  ReDim Preserve PixBits(0 To NewRowMod * Height - 1) As Byte
  PackBMPImage = True

End Function

'This routine takes an image in 1,4,8,16 BPP format which needs to be unpacked into 24BPP format
Public Function UnPackBMPImage(ByVal Width As Long, ByVal Height As Long, _
                               ByRef PixBits() As Byte, ByRef IntBPP As Long, _
                               ByRef CMap() As RGBA, ByVal NCMapColors As Long, _
                               ByVal RMask As Long, ByVal GMask As Long, ByVal BMask As Long) As Boolean   'deals with all MS formats

 Dim x As Long, y As Long, z As Long, w As Long, maskx As Long
 Dim RowMod As Long, NewRowMod As Long, p As Long, i As Integer, c As Integer
 Dim tmpP() As Byte
 Dim RRShift As Long, GRShift As Long, BRShift As Long
 Dim RMaxVal As Long, GMaxVal As Long, BMaxVal As Long
 
  If IntBPP <> PIC_1BPP _
  And IntBPP <> PIC_4BPP _
  And IntBPP <> PIC_8BPP _
  And IntBPP <> PIC_16BPP Then Exit Function

  ' Scan the PixBits() and unpack all pixels in each row of the matrix
  'As the new format will be Bigger or smaller (1½ to 24 times)
  'we unpack from the bottom of the image using a temporary storage array for 1 row at a time
  
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)  'the byte width of a given row
  NewRowMod = BMPRowModulo(Width, PIC_24BPP)                   'the byte width of a new row
  ReDim Preserve PixBits(0 To Height * NewRowMod - 1)          'this is what we become
  
  If IntBPP = PIC_16BPP Then
    Call MaskShift(BMask, BRShift, BMaxVal)
    Call MaskShift(GMask, GRShift, GMaxVal)
    Call MaskShift(RMask, RRShift, RMaxVal)
  End If
  
  For y = Height - 1 To 0 Step -1                              'assume bmap is right way up
    x = y * RowMod                                             'source bytes
    p = 0                                                      'dest bytes
    i = Width
    ReDim tmpP(0 To NewRowMod - 1)
    
    Select Case IntBPP
     Case PIC_1BPP:           '8 pixels per byte
      maskx = 128
      Do While i > 0
        If (PixBits(x) And maskx) <> 0 Then c = 1 Else c = 0
        With CMap(c)
          tmpP(p) = .Blue: p = p + 1
          tmpP(p) = .Green: p = p + 1
          tmpP(p) = .Red: p = p + 1
        End With
        maskx = maskx \ 2
        If maskx = 0 Then
          x = x + 1
          maskx = 128
        End If
        i = i - 1
      Loop

     Case PIC_4BPP:   '2 pixels per byte
      w = 0
      Do While i > 0
        c = PixBits(x)
        If (w And 1) = 0 Then     'first nybble
          c = (c And &HF0) \ 16
          w = w + 1
        Else
          c = c And &HF           'second nybble
          w = w + 1
          x = x + 1
        End If
        With CMap(c)
          tmpP(p) = .Blue: p = p + 1
          tmpP(p) = .Green: p = p + 1
          tmpP(p) = .Red: p = p + 1
        End With
        i = i - 1
      Loop

     Case PIC_8BPP:   '1 pixel per byte (dead easy)
      Do While i > 0
        With CMap(PixBits(x))
          tmpP(p) = .Blue: p = p + 1
          tmpP(p) = .Green: p = p + 1
          tmpP(p) = .Red: p = p + 1
        End With
        x = x + 1
        i = i - 1
      Loop

     Case PIC_16BPP:   '1 pixels per 2 bytes  'Two masks possible 5,5,5 and 5,6,5: no colormap
      If GMaxVal = 63 Then  'R5G6R5
        Do While i > 0
          w = CLng(PixBits(x)) + 256& * CLng(PixBits(x + 1))
          x = x + 2
          tmpP(p) = BMP31Scale((w And BMask) \ BRShift): p = p + 1  'Rescale to 0..255
          tmpP(p) = BMP63Scale((w And GMask) \ GRShift): p = p + 1
          tmpP(p) = BMP31Scale((w And RMask) \ RRShift): p = p + 1
          i = i - 1
        Loop
      Else                  'R5G5B5
        Do While i > 0
          w = CLng(PixBits(x)) + 256& * CLng(PixBits(x + 1))
          x = x + 2
          tmpP(p) = BMP31Scale((w And BMask) \ BRShift): p = p + 1  'Rescale to 0..255
          tmpP(p) = BMP31Scale((w And GMask) \ GRShift): p = p + 1
          tmpP(p) = BMP31Scale((w And RMask) \ RRShift): p = p + 1
          i = i - 1
        Loop
     End If
    End Select
    Call CopyMemoryRR(PixBits(y * NewRowMod), tmpP(0), NewRowMod)  'copy unpacked bytes to destination
  Next y

  IntBPP = PIC_24BPP
  Erase CMap
  NCMapColors = 0
  UnPackBMPImage = True
  
  '  MsgBox "OK UnPackBMPImage"

End Function

Private Function MaskShift(ByVal Mask As Long, ByRef RShift As Long, ByRef MaxVal As Long) As Long
  Dim i As Long
  
  RShift = 1
  Do While (RShift And Mask) = 0
    RShift = RShift + RShift
  Loop
  
  MaxVal = 1
  Mask = Mask \ RShift
  Do While (Mask And &H80&) = 0
    Mask = Mask + Mask
    MaxVal = MaxVal + MaxVal
  Loop
  MaxVal = 256 \ MaxVal - 1
End Function

'------------------- conversion between 24 and 32 bit formats ----------------------------------------------

'This routine takes an image in 32 BPP format which needs to be truncated into 24BPP format (we drop alpha)
Public Function BMPImage32to24BPP(ByVal Width As Long, ByVal Height As Long, _
                                  ByRef PixBits() As Byte, ByRef IntBPP As Long) As Boolean  'deals with all MS formats

 Dim x As Long, y As Long
 Dim RowMod As Long, NewRowMod As Long, p As Long, i As Long

  If IntBPP <> PIC_32BPP Then Exit Function

  ' Scan the PixBits() and unpack all pixels in each row of the matrix
  
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)  'the byte width of a given row
  NewRowMod = BMPRowModulo(Width, PIC_24BPP)                   'the byte width of a new row
  
  For y = 0 To Height - 1                                      'assume bmap is right way up
    x = y * RowMod                                             'this is the byte where we will put the result
    p = y * NewRowMod
    i = Width
    Do While i > 0
      PixBits(p) = PixBits(x)
      x = x + 1: p = p + 1
      PixBits(p) = PixBits(x)
      x = x + 1: p = p + 1
      PixBits(p) = PixBits(x)
      x = x + 2: p = p + 1        'skip alpha
      i = i - 1
    Loop
  Next y

  IntBPP = PIC_24BPP

  ReDim Preserve PixBits(0 To NewRowMod * Height - 1) As Byte
  BMPImage32to24BPP = True
  
  '  MsgBox "OK BMPImage32to24BPP"
  
End Function

'This routine takes an image in 24 BPP format which needs to be expanded into 32BPP format (we add alpha)
Public Function BMPImage24to32BPP(ByVal Width As Long, ByVal Height As Long, _
                                  ByRef PixBits() As Byte, ByRef IntBPP As Long) As Boolean  'deals with all MS formats

 Dim x As Long, y As Long
 Dim RowMod As Long, NewRowMod As Long, p As Long, i As Long

  If IntBPP <> PIC_24BPP Then Exit Function

  ' Scan the PixBits() and unpack all pixels in each row of the matrix
  
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)  'the byte width of a given row
  NewRowMod = BMPRowModulo(Width, PIC_32BPP)                   'the byte width of a new row
  ReDim Preserve PixBits(0 To NewRowMod * Height - 1) As Byte
  
  For y = Height To 1 Step -1                                  'assume bmap is right way up
    x = y * RowMod - 1                                         'source bytes
    p = y * NewRowMod - 1                                      'dest bytes
    i = Width
    Do While i > 0
      PixBits(p) = RGBtoGrey(PixBits(x - 2), PixBits(x - 1), PixBits(x))  'make alpha
      p = p - 1
      PixBits(p) = PixBits(x)   'R
      x = x - 1: p = p - 1
      PixBits(p) = PixBits(x)   'G
      x = x - 1: p = p - 1
      PixBits(p) = PixBits(x)   'B
      x = x - 1: p = p - 1
      i = i - 1
    Loop
  Next y

  IntBPP = PIC_32BPP

  BMPImage24to32BPP = True
  
  '  MsgBox "OK BMPImage24to32BPP"
  
End Function

'============================ UTILITY FUNCTIONS TO READ/WRITE 16 bit PIXELS =================================
'It May be useful from time to time to be able to read and write pixels, offset is automatically incremented
'Only Valid for 16 BPP Images (absolutely no error checking at all)
Public Function GetPixel16(ByRef PixBits() As Byte, ByRef Offset As Long) As RGBA
  
  Dim w As Long
  
  w = PixBits(Offset) Or 256& * PixBits(Offset + 1)
  Offset = Offset + 2
  
  With GetPixel16
    .Red = BMP31Scale((w And &H7C00&) \ 1024&)
    .Green = BMP31Scale((w And &H3E0&) \ 32&)
    .Blue = BMP31Scale((w And &H1F&))
  End With
  
End Function

Public Sub PutPixel16(ByRef PixBits() As Byte, ByRef Offset As Long, ByRef RGBAColor As RGBA)
  
  Dim w As Long
  
  With RGBAColor
    w = (.Red \ 8&) * 1024& Or (.Green \ 8&) * 32& Or (.Blue \ 8&)
    PixBits(Offset) = (w And &HFF&)
    PixBits(Offset + 1) = (w And &HFF00&) \ 256&
    Offset = Offset + 2
  End With
  
End Sub
':) Ulli's VB Code Formatter V2.16.6 (2003-Jun-13 22:35) 2 + 205 = 207 Lines
