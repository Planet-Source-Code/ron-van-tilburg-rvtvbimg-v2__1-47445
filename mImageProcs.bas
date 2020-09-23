Attribute VB_Name = "mImageProcs"
Option Explicit

'- ©2001/2003 Ron van Tilburg - All rights reserved  1.01.2001/Jun-Jul 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please
'mImageProcs.bas

' New Image processing options should be added in this file

'============================== SINGLE IMAGE OPERATIONS ======================================================

'This routine takes an image and flips it about the Y axis ie. Mirrors in the X axis
Public Function FlipImageHorz(ByVal Width As Long, ByVal Height As Long, _
                              ByVal IntBPP As Long, ByRef PixBits() As Byte) As Boolean 'deals with all MS formats

 Dim wk() As Byte, maskw As Integer, maskx As Integer, c As Integer
 Dim x As Long, y As Long, z As Long, w As Long, RowMod As Long, i As Integer

 ' Scan the PixBits() and reverse all pixels in each row of the matrix

  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)           'the byte width of a row
  ReDim wk(0 To RowMod - 1)
  For y = 0 To Height - 1                                'assume bmap is right way up
    w = ((Width - 1) * IntBPP) \ 8&                      'this is byte where the rightmost pixel is
    x = y * RowMod                                       'this is the byte where we will put the result
    Call CopyMemoryRR(wk(0), PixBits(x), RowMod)         'our working copy

    Select Case IntBPP
     Case PIC_24BPP:  ' Another for 24-bit DIBs
      Do
        PixBits(x) = wk(w): x = x + 1: w = w + 1
        PixBits(x) = wk(w): x = x + 1: w = w + 1
        PixBits(x) = wk(w): x = x + 1: w = w - 5
      Loop Until w < 0

     Case PIC_32BPP:  ' And another for 32-bit DIBs
      Do
        PixBits(x) = wk(w): x = x + 1: w = w + 1
        PixBits(x) = wk(w): x = x + 1: w = w + 1
        PixBits(x) = wk(w): x = x + 1: w = w + 1
        PixBits(x) = wk(w): x = x + 1: w = w - 7
      Loop Until w < 0
     
     Case PIC_1BPP:   '8 pixels per byte
      maskw = 2 ^ (7 - (Width - 1 - 8 * w)): maskx = 128
      Do
        c = 0
        Do
          If (wk(w) And maskw) <> 0 Then c = c Or maskx
          maskw = maskw + maskw
          If maskw > 128 Then maskw = 1: w = w - 1
          maskx = maskx \ 2
        Loop Until maskx = 0
        PixBits(x) = c
        x = x + 1
        maskx = 128
      Loop Until w < 1
      c = 0
      Do      'there are still some bits in byte0 left
        If (wk(w) And maskw) <> 0 Then c = c Or maskx
        maskw = maskw + maskw
        maskx = maskx \ 2
      Loop Until maskw > 128
      PixBits(x) = c

     Case PIC_4BPP:   '2 pixels per byte
      If (Width And 1) = 1 Then                              ' ab,cd,ef,g0  -> gf,ed,cb,a0
        Do
          PixBits(x) = (wk(w) And &HF0) Or (wk(w - 1) And &HF)
          x = x + 1: w = w - 1
        Loop Until w < 1
        PixBits(x) = wk(w) And &HF0
      Else                                                   ' ab,cd,ef,gh  -> hg,fe,dc,ba
        Do
          PixBits(x) = (16 * (wk(w) And &HF)) Or ((wk(w) And &HF0) \ 16)
          x = x + 1: w = w - 1
        Loop Until w < 0
      End If

     Case PIC_8BPP:   '1 pixel per byte    'The easiest one
      Do
        PixBits(x) = wk(w): x = x + 1: w = w - 1
      Loop Until w < 0

     Case PIC_16BPP:  ' One case for 16-bit DIBs
      Do
        PixBits(x) = wk(w): x = x + 1: w = w + 1
        PixBits(x) = wk(w): x = x + 1: w = w - 3
      Loop Until w < 0
    End Select
  Next y

  '  MsgBox "OK FlipImageHorz"
  FlipImageHorz = True

End Function

'Inplace vertical flip for all MS Formats
Public Function FlipImageVert(ByRef PicState As Long, _
                              ByVal Width As Long, ByVal Height As Long, _
                              ByVal IntBPP As Long, ByRef PixBits() As Byte) As Boolean  'deals with all MS formats

  Dim RowModulo As Long, tmp() As Byte, i As Long, s As Long, d As Long
  
  RowModulo = BMPRowModulo(Width, IntBPP)
  ReDim tmp(0 To RowModulo - 1)
  
  s = 0
  d = (Height - 1) * RowModulo
  For i = 0 To Height \ 2 - 1
    Call CopyMemoryRR(tmp(0), PixBits(s), RowModulo)
    Call CopyMemoryRR(PixBits(s), PixBits(d), RowModulo)
    Call CopyMemoryRR(PixBits(d), tmp(0), RowModulo)
    s = s + RowModulo
    d = d - RowModulo
  Next i
  Call ToggleF(PicState, IS_TOP_TO_BOTTOM)
  FlipImageVert = True
  
End Function

'Rotate Image Left 90 deg, requires a copy array, only 8,16,24,32 bPP supported
Public Function RotateImageLeft(ByRef PicState As Long, ByRef Width As Long, ByRef Height As Long, _
                                ByVal IntBPP As Long, ByRef PixBits() As Byte) As Boolean
                             
  Dim i As Long, j As Long, k As Long, p As Long, q As Long, r As Long, s As Long
  Dim OldPixBits() As Byte, OldRowMod As Long, NewRowMod As Long, Skip As Long
  
  If IntBPP >= PIC_8BPP Then
    OldPixBits() = PixBits()                    'original array copied
    NewRowMod = BMPRowModulo(Height, IntBPP)
    OldRowMod = BMPRowModulo(Width, IntBPP)
    ReDim PixBits(0 To Width * NewRowMod - 1)
    Skip = IntBPP \ 8
    
    For i = 0 To Height - 1
      p = i * OldRowMod
      q = i * Skip
      For j = 0 To Width - 1
        r = j * NewRowMod
        s = j * Skip
        For k = 0 To Skip - 1
          PixBits(r + q + k) = OldPixBits(p + s + k)
        Next k
      Next j
    Next i
    i = Width: Width = Height: Height = i
    RotateImageLeft = True
    If FSetClr(PicState, IS_TOP_TO_BOTTOM) Then
      Call FlipImageVert(PicState, Width, Height, IntBPP, PixBits())
    End If
  End If
  
End Function

'Rotate Image Right 90 deg, requires a copy array, only 8,16,24,32 BPP supported
Public Function RotateImageRight(ByRef PicState As Long, ByRef Width As Long, ByRef Height As Long, _
                                 ByVal IntBPP As Long, ByRef PixBits() As Byte) As Boolean
                             
  Dim i As Long, j As Long, k As Long, p As Long, q As Long, r As Long, s As Long
  Dim OldPixBits() As Byte, OldRowMod As Long, NewRowMod As Long, Skip As Long
  
  If IntBPP >= PIC_8BPP Then
    OldPixBits() = PixBits()                    'original array copied
    NewRowMod = BMPRowModulo(Height, IntBPP)
    OldRowMod = BMPRowModulo(Width, IntBPP)
    ReDim PixBits(0 To Width * NewRowMod - 1)
    Skip = IntBPP \ 8
    
    For i = 0 To Height - 1
      p = i * OldRowMod
      q = (Height - i - 1) * Skip
      For j = 0 To Width - 1
        r = (Width - 1 - j) * NewRowMod
        s = j * Skip
        For k = 0 To Skip - 1
          PixBits(r + q + k) = OldPixBits(p + s + k)
        Next k
      Next j
    Next i
    i = Width: Width = Height: Height = i
    RotateImageRight = True
    If FSetClr(PicState, IS_TOP_TO_BOTTOM) Then
      Call FlipImageVert(PicState, Width, Height, IntBPP, PixBits())
    End If
  End If
  

End Function

'Rotate Image 180 deg in place all MS formats are supported
Public Function RotateImage180(ByRef PicState As Long, ByRef Width As Long, ByRef Height As Long, _
                               ByVal IntBPP As Long, ByRef PixBits() As Byte) As Boolean
                               
  Call FlipImageHorz(Width, Height, IntBPP, PixBits())
  Call FlipImageVert(PicState, Width, Height, IntBPP, PixBits())
  Call ToggleF(PicState, IS_TOP_TO_BOTTOM)
  RotateImage180 = True
  
End Function

'Trim 8,16,24,32 BPP AlphaPreserved
Public Function TrimImage(ByRef Width As Long, ByRef Height As Long, _
                          ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                          ByVal LeftClip As Long, ByVal TopClip As Long, _
                          ByVal RightClip As Long, ByVal BottomClip As Long) As Boolean

  Dim i As Long, j As Long, p As Long, q As Long
  Dim OldRowMod As Long, NewRowMod As Long, Skip As Long
  Dim NewWidth As Long, NewHeight As Long
  
  If IntBPP < PIC_8BPP Then Exit Function
  
  Call ValidateClip(LeftClip, TopClip, Width - RightClip, Height - BottomClip, _
                    0, 0, Width, Height, _
                    NewWidth, NewHeight)
  
  NewRowMod = BMPRowModulo(NewWidth, IntBPP)
  OldRowMod = BMPRowModulo(Width, IntBPP)
  Skip = IntBPP \ 8
  
  p = 0
  q = TopClip * OldRowMod + LeftClip * Skip
  Skip = NewWidth * Skip
  
  For i = 0 To NewHeight - 1
    Call CopyMemoryRR(PixBits(p), PixBits(q), Skip)
    p = p + NewRowMod
    q = q + OldRowMod
  Next i
  
  ReDim Preserve PixBits(0 To NewHeight * NewRowMod - 1)
  Width = NewWidth
  Height = NewHeight
  TrimImage = True
  
End Function

'make a new sub image from main image, if sub image extends beyond original the extensions are coloured black
'16,24,32BPP AlphaPreserved
Public Function ExtractImage(ByVal Width As Long, ByVal Height As Long, _
                             ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                             ByVal atX As Long, ByVal atY As Long, _
                             ByVal XWidth As Long, ByVal XHeight As Long) As cRVTVBIMG
  
  Dim i As Long, j As Long, p As Long, q As Long
  Dim OldRowMod As Long, NewRowMod As Long, Skip As Long, I2Base As Long
                             
  If IntBPP < PIC_16BPP Then Exit Function
  
  Set ExtractImage = New cRVTVBIMG
  With ExtractImage
    Call .EraseImage(XWidth, XHeight, IntBPP)   'redimension and clear it
    I2Base = .PixBitsBasePtr
  End With
  
  OldRowMod = BMPRowModulo(Width, IntBPP)
  NewRowMod = BMPRowModulo(XWidth, IntBPP)
  Skip = IntBPP \ 8
  
  If atX < Width And atY < Height Then                          'there's something to copy
    If atX + XWidth <= Width Then Width = atX + XWidth
    If atY + XHeight <= Height Then Height = atY + XHeight
    
    For i = atY To Height - 1
      p = i * OldRowMod + Skip * atX
      q = (i - atY) * NewRowMod
      Call CopyMemoryVR(I2Base + q, PixBits(p), (Width - atX - 1) * Skip)
    Next i
  End If
  
  
End Function

'Nearest Neighbour Sampling Resize
'Resize 8,16,24,32 BPP alpha preserved, tmp copy required
Public Function ResizeImage(ByRef Width As Long, ByRef Height As Long, _
                            ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                            ByVal NewWidth As Long, ByVal NewHeight As Long) As Boolean
                            
 Dim kX As Double, kY As Double
 Dim xn As Long, yn As Long, pn As Long, qn As Long, po As Long, qo As Long
 Dim OldRowMod As Long, NewRowMod As Long, Skip As Long, OldPixBits() As Byte
 
  If IntBPP < PIC_8BPP Then Exit Function
  
  OldPixBits() = PixBits()                      'make copy of Old Array
  OldRowMod = BMPRowModulo(Width, IntBPP)
  NewRowMod = BMPRowModulo(NewWidth, IntBPP)
  Skip = IntBPP \ 8
  
  ReDim PixBits(0 To NewHeight * NewRowMod - 1)
  kX = CDbl(Width - 1) / CDbl(NewWidth - 1)     'a scaled fraction
  kY = CDbl(Height - 1) / CDbl(NewHeight - 1)
  
  pn = (NewHeight - 1) * NewRowMod
  For yn = NewHeight - 1 To 0 Step -1
    po = Int(yn * kY) * OldRowMod       'nearest
    qn = pn
    For xn = 0 To NewWidth - 1
      qo = po + Skip * Int(xn * kX)     'nearest
      PixBits(qn) = OldPixBits(qo): qn = qn + 1: qo = qo + 1
      If Skip > 1 Then PixBits(qn) = OldPixBits(qo): qn = qn + 1: qo = qo + 1
      If Skip > 2 Then PixBits(qn) = OldPixBits(qo): qn = qn + 1: qo = qo + 1
      If Skip > 3 Then PixBits(qn) = OldPixBits(qo): qn = qn + 1: qo = qo + 1
    Next xn
    pn = pn - NewRowMod
  Next yn
  
  Width = NewWidth
  Height = NewHeight
  ResizeImage = True

End Function

'Derived from Reconstructor by Peter Scale 2003 ex PSC (made into IntegerMaths version by RVT)
'NewDimension and OldDimension must be <= 2^11.5 ie. <= 2048
'Use bilinear interpolation to resize an image only works for 24 and 32bit images
'Requires a tmp copy of original, Only 24 and 32BPP are supported, Alpha is preserved
Public Function BilinearResizeImage(ByRef Width As Long, ByRef Height As Long, _
                                    ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                                    ByVal NewWidth As Long, ByVal NewHeight As Long) As Boolean

 Const SCALER As Long = 7983360          '2^8.3^4.5.7.11 ~= 2^22.928
 
 Dim kX As Double, kY As Double, tt As Double
 Dim fX As Long, fY As Long, gX As Long, gY As Long
 Dim xo As Long, yo As Long, po As Long, qo As Long, ro As Long
 Dim xn As Long, yn As Long, pn As Long, qn As Long
 Dim OldRowMod As Long, NewRowMod As Long, Skip As Long, OldPixBits() As Byte
 
  If IntBPP < PIC_24BPP Then Exit Function
  
  OldPixBits() = PixBits                    'make copy of Old Array
  OldRowMod = BMPRowModulo(Width, IntBPP)
  NewRowMod = BMPRowModulo(NewWidth, IntBPP)
  Skip = IntBPP \ 8
  
  ReDim PixBits(0 To NewHeight * NewRowMod - 1)
  kX = CDbl(Width - 1) / CDbl(NewWidth - 1)     'a scaled fraction
  kY = CDbl(Height - 1) / CDbl(NewHeight - 1)
  
  pn = (NewHeight - 1) * NewRowMod
  For yn = NewHeight - 1 To 0 Step -1
    tt = yn * kY                      ' Exact position
    yo = Int(tt)                      ' Integer position (integer part of number)
    fY = (tt - yo) * SCALER           ' Scaled Fraction part of number (integer+fraction=exact)
    gY = SCALER - fY                  ' Ones Complement of Fraction
    po = yo * OldRowMod
    qn = pn
    
    For xn = 0 To NewWidth - 1
      tt = xn * kX
      xo = Int(tt)
      fX = (tt - xo) * SCALER
      gX = SCALER - fX
      qo = po + Skip * xo
      ro = qo + OldRowMod
            
      If fX = 0 Then
        If fY = 0 Then  'Integer Rescale in X and Y
          PixBits(qn) = OldPixBits(qo): qn = qn + 1: qo = qo + 1
          PixBits(qn) = OldPixBits(qo): qn = qn + 1: qo = qo + 1
          PixBits(qn) = OldPixBits(qo): qn = qn + 1: qo = qo + 1
          If Skip > 3 Then PixBits(qn) = OldPixBits(qo): qn = qn + 1: qo = qo + 1
          
        Else            'Integer Rescale in X and interpolate Y
          PixBits(qn) = (gY * OldPixBits(qo) + fY * OldPixBits(ro)) \ SCALER
          qn = qn + 1: qo = qo + 1: ro = ro + 1
          PixBits(qn) = (gY * OldPixBits(qo) + fY * OldPixBits(ro)) \ SCALER
          qn = qn + 1: qo = qo + 1: ro = ro + 1
          PixBits(qn) = (gY * OldPixBits(qo) + fY * OldPixBits(ro)) \ SCALER
          qn = qn + 1: qo = qo + 1: ro = ro + 1
          If Skip > 3 Then
            PixBits(qn) = (gY * OldPixBits(qo) + fY * OldPixBits(ro)) \ SCALER
            qn = qn + 1: qo = qo + 1: ro = ro + 1
          End If
        End If
        
      ElseIf fY = 0 Then            'Integer Rescale in Y and interpolate X
        'b
        PixBits(qn) = (gX * OldPixBits(qo) + fX * OldPixBits(qo + Skip)) \ SCALER
        qn = qn + 1: qo = qo + 1
        'g
        PixBits(qn) = (gX * OldPixBits(qo) + fX * OldPixBits(qo + Skip)) \ SCALER
        qn = qn + 1: qo = qo + 1
        'r
        PixBits(qn) = (gX * OldPixBits(qo) + fX * OldPixBits(qo + Skip)) \ SCALER
        qn = qn + 1: qo = qo + 1
        If Skip > 3 Then
          'a
          PixBits(qn) = (gY * OldPixBits(qo) + fY * OldPixBits(ro)) \ SCALER
          qn = qn + 1: qo = qo + 1: ro = ro + 1
        End If
        
      Else  'Interpolation in X and Y
            ' Apply this formula: (1-frac) * RGB1 + frac * RGB2
            ' frac = fraction part of number <0;1); RGB1, RGB2 = red, green or blue part of color 1, 2
            ' It is applied 3 times for every part of color (2 times on X-axes and 1 times on Y-axes)
            ' The filter computes 1 point from 4 (2x2) surrounding points.
        'b
        PixBits(qn) = (gY * ((gX * OldPixBits(qo) + fX * OldPixBits(qo + Skip)) \ SCALER) _
                     + fY * ((gX * OldPixBits(ro) + fX * OldPixBits(ro + Skip)) \ SCALER)) \ SCALER
        qn = qn + 1: qo = qo + 1: ro = ro + 1
        'g
        PixBits(qn) = (gY * ((gX * OldPixBits(qo) + fX * OldPixBits(qo + Skip)) \ SCALER) _
                     + fY * ((gX * OldPixBits(ro) + fX * OldPixBits(ro + Skip)) \ SCALER)) \ SCALER
        qn = qn + 1: qo = qo + 1: ro = ro + 1
        'r
        PixBits(qn) = (gY * ((gX * OldPixBits(qo) + fX * OldPixBits(qo + Skip)) \ SCALER) _
                     + fY * ((gX * OldPixBits(ro) + fX * OldPixBits(ro + Skip)) \ SCALER)) \ SCALER
        qn = qn + 1: qo = qo + 1: ro = ro + 1
        If Skip > 3 Then
          'a
          PixBits(qn) = (gY * ((gX * OldPixBits(qo) + fX * OldPixBits(qo + Skip)) \ SCALER) _
                       + fY * ((gX * OldPixBits(ro) + fX * OldPixBits(ro + Skip)) \ SCALER)) \ SCALER
          qn = qn + 1: qo = qo + 1: ro = ro + 1
        End If
      End If
    Next xn
    pn = pn - NewRowMod
  Next yn
  Width = NewWidth
  Height = NewHeight
  BilinearResizeImage = True
  
End Function

'Inplace Sharpening of 24 and 32 bit images
Public Function SharpenImage(ByRef Width As Long, ByRef Height As Long, _
                             ByVal IntBPP As Long, ByRef PixBits() As Byte, ByVal Factor As Long) As Boolean

  SharpenImage = ConvolveImage(Width, Height, IntBPP, PixBits(), 1, Factor)

End Function

'Inplace Blurring of 24 and 32 bit images
Public Function BlurImage(ByRef Width As Long, ByRef Height As Long, _
                          ByVal IntBPP As Long, ByRef PixBits() As Byte, ByVal Factor As Long) As Boolean

  BlurImage = ConvolveImage(Width, Height, IntBPP, PixBits(), 2, Factor)

End Function

'Inplace Embossing of 24 and 32 bit images
Public Function EmbossImage(ByRef Width As Long, ByRef Height As Long, _
                            ByVal IntBPP As Long, ByRef PixBits() As Byte, ByVal Factor As Long) As Boolean

  EmbossImage = ConvolveImage(Width, Height, IntBPP, PixBits(), 3, Factor)

End Function

'Inplace Edge Detection of 24 and 32 bit images
Public Function EdgeImage(ByRef Width As Long, ByRef Height As Long, _
                          ByVal IntBPP As Long, ByRef PixBits() As Byte, Factor As Long) As Boolean

  EdgeImage = ConvolveImage(Width, Height, IntBPP, PixBits(), 4, Factor)

End Function

'Inplace Edge Detection of 24 and 32 bit images, all pixels <= Threshold presented in Black
Public Function HardEdgeImage(ByRef Width As Long, ByRef Height As Long, _
                              ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                              ByVal Factor As Long, Optional ByVal Threshold As Long = 240) As Boolean

  Dim RowMod As Long, Skip As Long
  Dim i As Long, j As Long, k As Long, p As Long
  
  If ConvolveImage(Width, Height, IntBPP, PixBits(), 4, Factor) Then
    If Threshold < 0 Or Threshold > 255 Then Threshold = 240
  
    RowMod = BMPRowModulo(Width, IntBPP)
    Skip = IntBPP \ 8
    For i = 0 To Height - 1
      p = i * RowMod
      For j = 0 To Width - 1
        If PixBits(p) <= Threshold Then PixBits(p) = 0: p = p + 1
        If PixBits(p) <= Threshold Then PixBits(p) = 0: p = p + 1
        If PixBits(p) <= Threshold Then PixBits(p) = 0: p = p + 1
        If Skip > 3 Then
          If PixBits(p) <= Threshold Then PixBits(p) = 0: p = p + 1
        End If
      Next j
    Next i
    HardEdgeImage = True
  End If
  
End Function

'Generalised Convolution, only 24 and 32BPP
'Sharpen (1),Blur(2), Emboss(3), Edge(4) Convolutions 24 and 32bit
'Sharpen [-1  -1]         Blur [1   1]      Edge[-1  -1]   Emboss [+1  -1]
' w = w +|  +4  |/6        w=w-|  4  |/6     w= |  +4  |     w =w+|      |/6
'        [-1  -1]              [1   1]          [-1  -1]          [-1  +1]
Private Function ConvolveImage(ByRef Width As Long, ByRef Height As Long, _
                             ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                             ByVal OpCode As Long, Factor As Long) As Boolean
                                          
  Const cvSHARPEN  As Long = 1
  Const cvBLUR     As Long = 2
  Const cvEMBOSS   As Long = 3
  Const cvEDGE     As Long = 4
  
  Dim Row0() As Byte, Row1() As Byte, Row2() As Byte, w As Long
  Dim RowMod As Long, Skip As Long
  Dim i As Long, j As Long, k As Long, ps As Long, pd As Long, qp As Long, qn As Long, qs As Long
  Dim f11 As Long, f00 As Long, f02 As Long, f20 As Long, f22 As Long, fsum As Long
  
  If IntBPP < PIC_24BPP Then Exit Function
  If Factor < 1 Or Factor > 9 Then Factor = 1
  
  Select Case OpCode
    Case cvSHARPEN: f11 = 4: f00 = -1: f02 = -1: f20 = -1: f22 = -1: fsum = 6
    Case cvBLUR:    f11 = 4: f00 = -1: f02 = -1: f20 = -1: f22 = -1: fsum = -6
    Case cvEMBOSS:  f11 = 0: f00 = -2: f02 = 2:  f20 = -2: f22 = 2:  fsum = 3
    Case cvEDGE:    f11 = 4: f00 = -1: f02 = -1: f20 = -1: f22 = -1: fsum = 1
  End Select
  
  RowMod = BMPRowModulo(Width, IntBPP)
  Skip = IntBPP \ 8
  
  ReDim Row0(0 To RowMod - 1), Row1(0 To RowMod - 1), Row2(0 To RowMod - 1)
  Call CopyMemoryRR(Row0(0), PixBits(0), RowMod)
  Call CopyMemoryRR(Row1(0), PixBits(RowMod), RowMod)
    
  ps = RowMod
  For i = 1 To Height - 2
    Call CopyMemoryRR(Row2(0), PixBits(ps + RowMod), RowMod) 'get next row
    For j = 1 To Width - 2
      qs = j * Skip: pd = ps + qs: qp = qs - Skip: qn = qs + Skip
      For k = 1 To Skip
        If OpCode = cvEDGE Then
          w = (Factor * (f00 * CLng(Not Row0(qp)) + f20 * CLng(Not Row2(qp)) _
                                  + f11 * CLng(Not Row1(qs)) _
                       + f02 * CLng(Not Row0(qn)) + f22 * CLng(Not Row2(qn)))) \ fsum
        Else
          w = CLng(Row1(qs)) _
             + Factor * ((f00 * CLng(Row0(qp)) + f20 * CLng(Row2(qp)) _
                               + f11 * CLng(Row1(qs)) _
                        + f02 * CLng(Row0(qn)) + f22 * CLng(Row2(qn)))) \ fsum
        End If
        If w < 0 Then w = 0 Else If w > 255 Then w = 255
        If OpCode = cvEDGE Then w = 255 - w
        PixBits(pd) = w
        pd = pd + 1: qs = qs + 1: qp = qp + 1: qn = qn + 1
      Next k
    Next j
    Row0() = Row1() 'copy row up
    Row1() = Row2() 'copy row up
    ps = ps + RowMod
  Next i
  ConvolveImage = True
  
End Function

'============================== DUAL IMAGE OPERATIONS ======================================================

':) Ulli's VB Code Formatter V2.16.6 (2003-Jun-13 22:36) 10 + 89 = 99 Lines
