Attribute VB_Name = "mColorProcs"
Option Explicit


'- ©2001/2003 Ron van Tilburg - All rights reserved  1.01.2001/Jun-Jul 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please
'mColorProcs.bas

' New Color processing options should be added in this file

'======================= COLOR OPERATIONS ====================================================================

'InPlace Invert all colors - works for all mapped and unmapped types
Public Function InvertImageColor(ByRef Width As Long, ByRef Height As Long, _
                                 ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                                 ByRef CMap() As RGBA, ByVal NCMapColors) As Boolean

  Dim i As Long, w As Integer
  
  If NCMapColors > 0 Then   'is mapped
    For i = 0 To NCMapColors - 1
      With CMap(i)
        .Red = Not .Red
        .Green = Not .Green
        .Blue = Not .Blue
        .Alpha = RGBtoGrey(.Red, .Green, .Blue)
      End With
    Next i
  Else                      'isnt mapped
    Select Case IntBPP
      Case PIC_16BPP:       'Both Formats 555 and 565
        For i = 0 To UBound(PixBits) Step 2
          w = Not (PixBits(i) Or 256& * PixBits(i + 1))
          PixBits(i) = (w And &HFF)
          PixBits(i + 1) = (w And &HFF00) \ 256
        Next i

      Case Else           'everything else
        For i = 0 To UBound(PixBits)
          PixBits(i) = Not (PixBits(i))
        Next
    End Select
  End If
  
  InvertImageColor = True
  
End Function

'Replace colours in and around the TargetColor with ReplacementColor
Public Function ReplaceImageColor(ByRef Width As Long, ByRef Height As Long, _
                                  ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                                  ByRef CMap() As RGBA, ByVal NCMapColors, _
                                  ByVal TargetRGBColor As Long, _
                                  ByVal SearchRadius As Long, _
                                  ByVal ReplacementRGBColor As Long) As Boolean
  
  Dim i As Long, j As Long, p As Long, tc As RGBA, rc As RGBA
  Dim RowMod As Long, Skip As Long, SearchRange As Long, Pixel As RGBA
  
  SearchRadius = Abs(SearchRadius)
  If SearchRadius > 255 Then SearchRadius = 255
  
  tc = GetRGBA(TargetRGBColor)
  rc = GetRGBA(ReplacementRGBColor)
  SearchRange = SearchRadius * SearchRadius
  
  If NCMapColors > 0 Then   'is mapped
    For i = 0 To NCMapColors - 1
      If ColorMatch(SearchRange, CMap(i), tc) Then
        CMap(i) = rc
        With CMap(i)
          .Alpha = RGBtoGrey(.Red, .Green, .Blue)
        End With
      End If
    Next i
  Else                      'isnt mapped
    RowMod = BMPRowModulo(Width, IntBPP)
    Skip = IntBPP \ 8
    
    Select Case IntBPP
      Case PIC_24BPP, PIC_32BPP:
        For i = 0 To Height - 1
          p = i * RowMod
          For j = 0 To Width - 1
            Pixel.Blue = PixBits(p)
            Pixel.Green = PixBits(p + 1)
            Pixel.Red = PixBits(p + 2)
            If Skip > 3 Then Pixel.Alpha = PixBits(p + 3)
            If ColorMatch(SearchRange, Pixel, tc) Then
              PixBits(p) = rc.Blue:  p = p + 1
              PixBits(p) = rc.Green: p = p + 1
              PixBits(p) = rc.Red:   p = p + 1
              If Skip > 3 Then PixBits(p) = rc.Alpha: p = p + 1
            Else
              p = p + Skip
            End If
          Next j
        Next i
      
      Case PIC_16BPP:
        For i = 0 To Height - 1
          p = i * RowMod
          For j = 0 To Width - 1
            Pixel = GetPixel16(PixBits(), (p))
            If ColorMatch(SearchRange, Pixel, tc) Then
              Call PutPixel16(PixBits(), p, rc)
            Else
              p = p + Skip
            End If
          Next j
        Next i
    End Select
  End If
  
  ReplaceImageColor = True

End Function

Private Function ColorMatch(ByVal Range As Long, ByRef Color1 As RGBA, ByRef Color2 As RGBA) As Boolean
  
  Dim v As Long
  
  With Color1
    If Range = 0 Then                   'is it an exact match
      If .Blue = Color2.Blue Then
        If .Green = Color2.Green Then
          If .Red = Color2.Red Then
            ColorMatch = True
          End If
        End If
      End If
    Else                                'is it a near match
      v = CLng(.Red) - CLng(Color2.Red)
      Range = Range - v * v
      If Range > 0 Then
        v = CLng(.Green) - CLng(Color2.Green)
        Range = Range - v * v
        If Range > 0 Then
          v = CLng(.Blue) - CLng(Color2.Blue)
          Range = Range - v * v
          If Range >= 0 Then
            ColorMatch = True
          End If
        End If
      End If
    End If
  End With
  'if not found then FALSE
End Function

'Equalize image according to greyscale, a temp array is used for speed 24 and 32bpp images
'in 32BPP images the value of Alpha is Unchanged
Public Function EqualizeImageColor(ByRef Width As Long, ByRef Height As Long, _
                                   ByVal IntBPP As Long, ByRef PixBits() As Byte) As Boolean
                                     
  Const NSTEPS As Long = 1023
  
  Dim RowMod As Long, tmpMod As Long, Skip As Long, tmpG() As Integer
  Dim i As Long, j As Long, k As Long
  Dim p As Long, q As Long, grey(0 To NSTEPS) As Long
  Dim r As Long, g As Long, b As Long, v As Long, maxV As Long, minV As Long, meanV As Long
  
  If IntBPP < PIC_24BPP Then Exit Function
  
  RowMod = BMPRowModulo(Width, IntBPP)
  tmpMod = BMPRowModulo(Width, PIC_8BPP)
  ReDim tmpG(0 To Height * tmpMod - 1)
  
  Skip = IntBPP \ 8
  
  'count grey values
  
  For i = 0 To Height - 1
    p = i * RowMod
    q = i * tmpMod
    For j = 0 To Width - 1
      r = PixBits(p): p = p + 1
      g = PixBits(p): p = p + 1
      b = PixBits(p): p = p + Skip - 2
      v = (19652& * r + 38584 * g + 7493& * b) \ 16384&    'Recast for 0-1023
      tmpG(q) = v: q = q + 1
      grey(v) = grey(v) + 1
    Next j
  Next i
  
  'accumulate greys and form histogram, then differences from expected value
  For i = 1 To NSTEPS
    grey(i) = grey(i) + grey(i - 1)
  Next i
  
  q = grey(0)
  p = grey(NSTEPS) - q
  For i = 0 To NSTEPS
    grey(i) = (NSTEPS * (grey(i) - q)) \ p - i   'the delta to apply
  Next i

  'now fix the components of the colors
  For i = 0 To Height - 1
    p = i * RowMod
    q = i * tmpMod
    For j = 0 To Width - 1
      If grey(tmpG(q)) <> 0 Then
        For k = 1 To 3
          v = (CLng(PixBits(p)) * (tmpG(q) + grey(tmpG(q)))) \ tmpG(q)
          If v < 0 Then v = 0 Else If v > 255 Then v = 255
          PixBits(p) = v: p = p + 1
        Next k
        p = p + Skip - 3
      Else
        p = p + Skip
      End If
      q = q + 1
    Next j
  Next i

  EqualizeImageColor = True
  
End Function

'Idea Drawn from pnmgamma.c - perform gamma correction on a portable pixmap
' Copyright (C) 1991 by Bill Davidson and Jef Poskanzer. netpbm 10.15
'all formats, if 32BPP alpha is untouched
'If gamma value passed in <0 then The inverse is applied  + ==> brighter - ==> darker
Public Function GammaImageColor(ByRef Width As Long, ByRef Height As Long, _
                                ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                                ByRef CMap() As RGBA, ByVal NCMapColors, _
                                Optional ByVal RGamma As Single = 1, _
                                Optional ByVal GGamma As Single = 1, _
                                Optional ByVal BGamma As Single = 1) As Boolean
                               
  Dim RGammaTab(0 To 255) As Byte, GGammaTab(0 To 255) As Byte, BGammaTab(0 To 255) As Byte, g As Single
    
  g = Abs(RGamma)
  If g < 1 Then g = 1
  If g > 2.4 Then g = 2.4
  Call BuildsRGBGamma(RGammaTab(), g, RGamma < 0)
  
  g = Abs(GGamma)
  If g < 1 Then g = 1
  If g > 2.4 Then g = 2.4
  Call BuildsRGBGamma(GGammaTab(), g, GGamma < 0)
  
  g = Abs(BGamma)
  If g < 1 Then g = 1
  If g > 2.4 Then g = 2.4
  Call BuildsRGBGamma(BGammaTab(), g, BGamma < 0)
  
  GammaImageColor = ColorAdjust(Width, Height, IntBPP, PixBits(), CMap(), NCMapColors, _
                                RGammaTab(), GGammaTab(), BGammaTab())
End Function
   
'Build a gamma table of size maxval+1 for the IEC SRGB gamma transfer function (Standard IEC 61966-2-1).
'  'gamma' must be 2.4 for true SRGB
Private Sub BuildsRGBGamma(ByRef Table() As Byte, ByVal Gamma As Single, ByVal InvGamma As Boolean)
    
  Dim OneOverGamma As Double, i As Long, LinearCutoff As Long, LinearExpansion As Double
  Dim Normalized As Double, v As Long
  
  OneOverGamma = 1# / Gamma
 
  ' This transfer function is linear for sample values 0
  ' .. 255*.040405 and an exponential for larger sample values.
  ' The exponential is slightly stretched and translated, though,
  ' unlike the popular pure exponential gamma transfer function.

  LinearCutoff = CLng(255# * 0.040405 + 0.5)
  LinearExpansion = (1.055 * (0.040405 ^ OneOverGamma) - 0.055) / 0.040405

  If InvGamma Then
    For i = 0 To LinearCutoff
      Table(i) = CLng(i / LinearExpansion)
    Next i
    Do While i <= 255
      Normalized = CDbl(i) / 255#                                           ' sample value normalized to 0..1
      v = CLng(255# * (((Normalized + 0.055) / 1.055) ^ Gamma) + 0.5)       ' denormalize, round
      If v > 255 Then v = 255                                               ' clamp
      Table(i) = v
      i = i + 1
    Loop
  Else
    For i = 0 To LinearCutoff
      Table(i) = CLng(i * LinearExpansion)
    Next i
    Do While i <= 255
      Normalized = CDbl(i) / 255#                                           ' sample value normalized to 0..1
      v = CLng(255# * (1.055 * (Normalized ^ OneOverGamma) - 0.055) + 0.5)  ' denormalize, round
      If v > 255 Then v = 255                                               ' clamp
      Table(i) = v
      i = i + 1
    Loop
  End If
  
End Sub

Public Function BrightnessImageColor(ByRef Width As Long, ByRef Height As Long, _
                                     ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                                     ByRef CMap() As RGBA, ByVal NCMapColors, _
                                     Optional ByVal RFactor As Single = 0, _
                                     Optional ByVal GFactor As Single = 0, _
                                     Optional ByVal BFactor As Single = 0) As Boolean
                               
  Dim RBrightTab(0 To 255) As Byte, GBrightTab(0 To 255) As Byte, BBrightTab(0 To 255) As Byte
  Dim i As Long, rf As Long, gf As Long, bf As Long
    
  If RFactor < -1 Then RFactor = -1 Else If RFactor > 1 Then RFactor = 1
  If GFactor < -1 Then GFactor = -1 Else If GFactor > 1 Then GFactor = 1
  If BFactor < -1 Then BFactor = -1 Else If BFactor > 1 Then BFactor = 1
  
  rf = Int(255 * RFactor)
  gf = Int(255 * GFactor)
  bf = Int(255 * BFactor)
  
  For i = 0 To 255
    RBrightTab(i) = i + (rf * i * (255 - i)) \ 65025
    GBrightTab(i) = i + (gf * i * (255 - i)) \ 65025
    BBrightTab(i) = i + (bf * i * (255 - i)) \ 65025
  Next i
  
  BrightnessImageColor = ColorAdjust(Width, Height, IntBPP, PixBits(), CMap(), NCMapColors, _
                                     RBrightTab(), GBrightTab(), BBrightTab())
  
End Function

'Given Adjustment Tables, adjust the colours of the Image (called by Gamma Correct, and Brightness Correct)

Private Function ColorAdjust(ByRef Width As Long, ByRef Height As Long, _
                             ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                             ByRef CMap() As RGBA, ByVal NCMapColors, _
                             ByRef RTable() As Byte, _
                             ByRef GTable() As Byte, _
                             ByRef BTable() As Byte) As Boolean
  
  Dim i As Long, j As Long, p As Long, Pixel As RGBA
  Dim RowMod As Long, Skip As Long
  
  If NCMapColors > 0 Then   'is mapped
    For i = 0 To NCMapColors - 1
      With CMap(i)
        .Red = RTable(.Red)
        .Green = GTable(.Green)
        .Blue = BTable(.Blue)
        .Alpha = RGBtoGrey(.Red, .Green, .Blue)
      End With
    Next i
  Else                      'isnt mapped
    RowMod = BMPRowModulo(Width, IntBPP)
    Skip = IntBPP \ 8
    
    Select Case IntBPP
      Case PIC_24BPP, PIC_32BPP:
        For i = 0 To Height - 1
          p = i * RowMod
          For j = 0 To Width - 1
            PixBits(p) = BTable(PixBits(p)): p = p + 1
            PixBits(p) = GTable(PixBits(p)): p = p + 1
            PixBits(p) = RTable(PixBits(p)): p = p + 1
            If Skip > 3 Then p = p + 1
          Next j
        Next i
      
      Case PIC_16BPP:
        For i = 0 To Height - 1
          p = i * RowMod
          For j = 0 To Width - 1
            Pixel = GetPixel16(PixBits(), (p))
            Pixel.Red = RTable(Pixel.Red)
            Pixel.Green = GTable(Pixel.Green)
            Pixel.Blue = BTable(Pixel.Blue)
            Call PutPixel16(PixBits(), p, Pixel)
          Next j
        Next i
    End Select
  End If
  
  ColorAdjust = True
  
End Function

'put a colored border around the Image 16,24 and 32bpp
'for 32BPP the alpha from FrameRGBColor is used
Public Function FrameImage(ByRef Width As Long, ByRef Height As Long, _
                           ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                           ByVal LeftRight As Long, ByVal TopBottom As Long, _
                           ByVal FrameRGBColor As Long) As Boolean

  Dim i As Long, j As Long, p As Long
  Dim RowMod As Long, Skip As Long, cc As RGBA
  
  If IntBPP < PIC_16BPP Then Exit Function
  
  If LeftRight <= 0 Or LeftRight > Width \ 2 Then LeftRight = 1
  If TopBottom <= 0 Or TopBottom > Height \ 2 Then TopBottom = 1
  
  RowMod = BMPRowModulo(Width, IntBPP)
  Skip = IntBPP \ 8

  cc = GetRGBA(FrameRGBColor)
  
  Select Case IntBPP
    Case PIC_24BPP, PIC_32BPP:
      For i = 0 To Height - 1
        p = i * RowMod
        If i < TopBottom Or i >= Height - TopBottom Then
          For j = 0 To Width - 1
            PixBits(p) = cc.Blue:  p = p + 1
            PixBits(p) = cc.Green: p = p + 1
            PixBits(p) = cc.Red:   p = p + 1
            If Skip > 3 Then PixBits(p) = cc.Alpha: p = p + 1
          Next j
        Else
          For j = 0 To LeftRight - 1
            PixBits(p) = cc.Blue:  p = p + 1
            PixBits(p) = cc.Green: p = p + 1
            PixBits(p) = cc.Red:   p = p + 1
            If Skip > 3 Then PixBits(p) = cc.Alpha: p = p + 1
          Next j
          p = i * RowMod + (Width - LeftRight) * Skip
          For j = 0 To LeftRight - 1
            PixBits(p) = cc.Blue:  p = p + 1
            PixBits(p) = cc.Green: p = p + 1
            PixBits(p) = cc.Red:   p = p + 1
            If Skip > 3 Then PixBits(p) = cc.Alpha: p = p + 1
          Next j
        End If
      Next i
      
    Case PIC_16BPP:
      For i = 0 To Height - 1
        p = i * RowMod
        If i < TopBottom Or i >= Height - TopBottom Then
          For j = 0 To Width - 1
            Call PutPixel16(PixBits(), p, cc)
          Next j
        Else
          For j = 0 To LeftRight - 1
            Call PutPixel16(PixBits(), p, cc)
          Next
          p = i * RowMod + (Width - LeftRight) * Skip
          For j = 0 To LeftRight - 1
            Call PutPixel16(PixBits(), p, cc)
          Next
        End If
      Next i
  End Select
  
  FrameImage = True

End Function

