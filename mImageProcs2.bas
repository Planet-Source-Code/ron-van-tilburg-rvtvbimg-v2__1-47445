Attribute VB_Name = "mImageProcs2"
Option Explicit

'mImageProcs2.bas
'- ©2003 Ron van Tilburg - All rights reserved  29 July 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

' Image Combination processing options should be added in this file

'----------------------------------------------------------------------------------------------------------
'Global in this Module for Processing speed - so we dont need to pass them around between modules
Private P1      As RGBA   'Pixel1
Private P2      As RGBA   'Pixel2
Private AMask   As RGBA   'Alpha Mask
Private OpCode  As Long   'Combination Opcode

'Blend as ( P1 (255 - AMask) + P2 * AMask) \ 255  for each component of RGB, A is unchanged
Public Sub AlphaBlendPixels()
  
  Dim v As Long, w As Long
  
  With P1
    v = .Red: w = P2.Red:     .Red = (v * (Not AMask.Red) + w * AMask.Red) \ 255&
    v = .Green: w = P2.Green: .Green = (v * (Not AMask.Green) + w * AMask.Green) \ 255&
    v = .Blue: w = P2.Blue:   .Blue = (v * (Not AMask.Blue) + w * AMask.Blue) \ 255&
  End With
  
End Sub

'UnBlend Result = (255*P1 - AMask*P2)\(NOT AMask); if Mask=255 then 255 is returned
Public Sub InvAlphaBlendPixels()  'Inverse ie Subtract an Alpha Amount of P2
  
  Dim v As Long, w As Long
  
  With P1
    If AMask.Red = 255 Then
      .Red = 255
    Else
      v = .Red: w = P2.Red
      v = (255 * v - w * AMask.Red) \ (Not AMask.Red)
      If v < 0 Then v = 0 Else If v > 255 Then v = 255
      .Red = v
    End If
    
    If AMask.Green = 255 Then
      .Green = 255
    Else
      v = .Green: w = P2.Green
      v = (255 * v - w * AMask.Green) \ (Not AMask.Green)
      If v < 0 Then v = 0 Else If v > 255 Then v = 255
      .Green = v
    End If
    
    If AMask.Blue = 255 Then
      .Blue = 255
    Else
      v = .Blue: w = P2.Blue
      v = (255 * v - w * AMask.Blue) \ (Not AMask.Blue)
      If v < 0 Then v = 0 Else If v > 255 Then v = 255
      .Blue = v
    End If
  End With
  
End Sub

'P1 = (NOT) Opcode( (NOT) P1, (NOT) P2 )
Public Sub CombinePixels()
  
  Dim r1 As Long, g1 As Long, b1 As Long
  Dim r2 As Long, g2 As Long, b2 As Long
  Dim OpCode2 As Long, v As Long
  
  If (OpCode And PCC_NOT1) = 0 Then
    r1 = P1.Red: g1 = P1.Green: b1 = P1.Blue
  Else
    r1 = Not P1.Red: g1 = Not P1.Green: b1 = Not P1.Blue
  End If
  
  If (OpCode And PCC_NOT2) = 0 Then
    r2 = P2.Red: g2 = P2.Green: b2 = P2.Blue
  Else
    r2 = Not P2.Red: g2 = Not P2.Green: b2 = Not P2.Blue
  End If
  
  If (OpCode And 128) = 0 Then    'is not a separation
    Select Case OpCode And &HF
      Case PCC_COPY:
        r1 = r2: g1 = g2: b1 = b2
      
      Case PCC_BLACK:
        r1 = 0:  g1 = 0:  b1 = 0
      
      Case PCC_WHITE:
        r1 = 255: g1 = 255: b1 = 255
      
      Case PCC_ALT:
        If r1 > 127 Then r1 = r2 'r1 = 255 Else r1 = 0
        If g1 > 127 Then g1 = g2 'g1 = 255 Else g1 = 0
        If b1 > 127 Then b1 = b2 'b1 = 255 Else b1 = 0
        
      Case PCC_AND:
        r1 = r1 And r2
        g1 = g1 And g2
        b1 = b1 And b2
      
      Case PCC_OR:
        r1 = r1 Or r2
        g1 = g1 Or g2
        b1 = b1 Or b2
      
      Case PCC_XOR:
        r1 = r1 Xor r2
        g1 = g1 Xor g2
        b1 = b1 Xor b2
      
      Case PCC_ADD:
        r1 = r1 + r2:   If r1 > 255 Then r1 = 255
        g1 = g1 + g2:   If g1 > 255 Then g1 = 255
        b1 = b1 + b2:   If b1 > 255 Then b1 = 255
      
      Case PCC_SUB:
        r1 = r1 - r2:   If r1 < 0 Then r1 = 0
        g1 = g1 - g2:   If g1 < 0 Then g1 = 0
        b1 = b1 - b2:   If b1 < 0 Then b1 = 0
      
      Case PCC_MOD:
        r1 = (r1 * r2) \ 255&
        g1 = (g1 * g2) \ 255&
        b1 = (b1 * b2) \ 255&
      
      Case PCC_MOD2:
        r1 = (2 * r1 * r2) \ 255&: If r1 > 255 Then r1 = 255
        g1 = (2 * g1 * g2) \ 255&: If g1 > 255 Then g1 = 255
        b1 = (2 * b1 * b2) \ 255&: If b1 > 255 Then b1 = 255
      
      Case PCC_ADDM:
        r1 = r1 + r2 - (r1 * r2) \ 255&
        g1 = g1 + g2 - (g1 * g2) \ 255&
        b1 = b1 + b2 - (b1 * b2) \ 255&
      
      Case PCC_MIN:
        If r1 >= b2 Then r1 = r2
        If g1 >= b2 Then g1 = g1
        If b1 >= b2 Then b1 = b2
        
      Case PCC_MAX:
        If r1 <= r2 Then r1 = r2
        If g1 <= g2 Then g1 = g2
        If b1 <= b2 Then b1 = b2
        
      Case PCC_AVE:
        r1 = (r1 + r2) \ 2
        g1 = (g1 + g2) \ 2
        b1 = (b1 + b2) \ 2
        
      Case PCC_DOT:   'scaled to 0..255
        r1 = Sqr((r1 * r1 + r2 * r2) \ 2)
        g1 = Sqr((g1 * g1 + g2 * g2) \ 2)
        b1 = Sqr((b1 * b1 + b2 * b2) \ 2)
        
      Case Else:   'DOES NOTHING
    End Select
  Else
    v = r1
    If g1 < v Then v = g1
    If b1 < v Then v = b1   'v is the smallest of r,g,b
    r1 = r1 - v
    g1 = g1 - v
    b1 = b1 - v
    
    Select Case OpCode2        'separations
      Case PSC_RED:
        g1 = 0: b1 = 0
        
      Case PSC_GREEN:
        r1 = 0: b1 = 0
        
      Case PSC_BLUE:
        r1 = 0: g1 = 0
        
      Case PSC_YELLOW:
        b1 = 0
        If r1 > g1 Then r1 = g1 Else g1 = r1
        
      Case PSC_MAGENTA:
        g1 = 0
        If r1 > b1 Then r1 = b1 Else b1 = r1
        
      Case PSC_CYAN:
        r1 = 0
        If b1 > g1 Then b1 = g1 Else g1 = b1
        
      Case PSC_BLACK:     'min(r, g, b)
        r1 = v: g1 = v: b1 = v
      
      Case Else:   'DOES NOTHING
    End Select
  End If
  
  With P1
    .Red = r1: .Green = g1: .Blue = b1
    If (OpCode And PCC_NOTR) = PCC_NOTR Then
      .Red = Not .Red: .Green = Not .Green: .Blue = Not .Blue
    End If
  End With
  
End Sub

'-----------------------------------------------------------------------------------------------------------
'Blend original image with a color as (1-A)*Image + A*Color
'OR Combine original image with a color as Image OPCode Color
Public Function CombineImageColor(ByRef Width As Long, ByRef Height As Long, _
                                  ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                                  ByRef CMap() As RGBA, ByVal NCMapColors, _
                                  ByVal OpCode_RGBAMask As Long, ByVal RGBColor As Long, _
                                  Optional ByVal AlphaBlending As Long = 0) As Boolean

  Dim i As Long, j As Long, p As Long
  Dim RowMod As Long, Skip As Long
  
  P2 = GetRGBA(RGBColor)
  
  If AlphaBlending Then
    AMask = GetRGBA(OpCode_RGBAMask)
  Else
    OpCode = OpCode_RGBAMask
  End If
  
  If NCMapColors > 0 Then   'is mapped
    For i = 0 To NCMapColors - 1
      P1 = CMap(i)
      If AlphaBlending = 1 Then
        Call AlphaBlendPixels
      ElseIf AlphaBlending = -1 Then
        Call InvAlphaBlendPixels
      Else
        Call CombinePixels
      End If
      With CMap(i)
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
            P1.Blue = PixBits(p)
            P1.Green = PixBits(p + 1)
            P1.Red = PixBits(p + 2)
            If Skip > 3 Then P1.Alpha = PixBits(p + 3)
            If AlphaBlending = 1 Then
              Call AlphaBlendPixels
            ElseIf AlphaBlending = -1 Then
              Call InvAlphaBlendPixels
            Else
              Call CombinePixels
            End If
            PixBits(p) = P1.Blue:  p = p + 1
            PixBits(p) = P1.Green: p = p + 1
            PixBits(p) = P1.Red:   p = p + 1
            If Skip > 3 Then PixBits(p) = P1.Alpha: p = p + 1
          Next j
        Next i
      
      Case PIC_16BPP:
        For i = 0 To Height - 1
          p = i * RowMod
          For j = 0 To Width - 1
            P1 = GetPixel16(PixBits(), (p))
            If AlphaBlending = 1 Then
              Call AlphaBlendPixels
            ElseIf AlphaBlending = -1 Then
              Call InvAlphaBlendPixels
            Else
              Call CombinePixels
            End If
            Call PutPixel16(PixBits(), p, P1)
          Next j
        Next i
    End Select
  End If
  
  CombineImageColor = True
                                     
End Function

'If Image2 is smaller than Image1 it can be Tiled until Image1 is completely blended
'If Image2 is larger than Image1 then only the Image1 part is blended from top left
'RGBAMask => each component is used in blend separately, for all the same fill in each as a grey
'EITHER Blend original image with a color as Image1 = (1-A)*Image1 + A*Image2
'OR   Combine original image with a color as Image1=Image1 OpCode Image2
Public Function Combine2Images(ByVal Width As Long, ByVal Height As Long, _
                               ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                               ByVal Image2 As cRVTVBIMG, _
                               ByVal OpCode_RGBAMask As Long, _
                               Optional ByVal AutoTile As Boolean = False, _
                               Optional ByVal AlphaBlending As Long = 0) As Boolean
  
  Dim i1 As Long, j1 As Long, i2 As Long, j2 As Long, p As Long
  Dim RowMod1 As Long, Skip1 As Long
  Dim RowMod2 As Long, Skip2 As Long, Width2 As Long, Height2 As Long

  If IntBPP < PIC_16BPP Then Exit Function                'Image  must be 16,24 or 32 BPP
  If Image2.BitsPerPixel < PIC_8BPP Then Exit Function    'Image2 must be 8,16,24 or 32 BPP
  
  RowMod1 = BMPRowModulo(Width, IntBPP)
  Skip1 = IntBPP \ 8
  With Image2
    RowMod2 = .RowModulo
    Skip2 = .BitsPerPixel \ 8
    Width2 = .Width
    Height2 = .Height
  End With
  
  If AlphaBlending Then
    AMask = GetRGBA(OpCode_RGBAMask)
  Else
    OpCode = OpCode_RGBAMask
  End If
    
  If Not AutoTile Then                      'constrain blend to the smallest area
    If Width2 < Width Then Width = Width2
    If Height2 < Height Then Height = Height2
  End If
  
  With Image2
    Select Case IntBPP
      Case PIC_24BPP, PIC_32BPP:
        i2 = 0
        For i1 = 0 To Height - 1
          p = i1 * RowMod1
          j2 = 0
          For j1 = 0 To Width - 1
            P1.Blue = PixBits(p)
            P1.Green = PixBits(p + 1)
            P1.Red = PixBits(p + 2)
            If Skip1 > 3 Then P1.Alpha = PixBits(p + 3)
            P2 = .GetPixelz(i2 * RowMod2 + Skip2 * j2)
            If AlphaBlending = 1 Then
              Call AlphaBlendPixels
            ElseIf AlphaBlending = -1 Then
              Call InvAlphaBlendPixels
            Else
              Call CombinePixels
            End If
            PixBits(p) = P1.Blue: p = p + 1
            PixBits(p) = P1.Green: p = p + 1
            PixBits(p) = P1.Red: p = p + 1
            If Skip1 > 3 Then PixBits(p) = P1.Alpha: p = p + 1
            j2 = j2 + 1
            If j2 >= Width2 Then j2 = 0
          Next j1
          i2 = i2 + 1
          If i2 >= Height2 Then i2 = 0
        Next i1
      
      Case PIC_16BPP:  'this is quite a bit slower
        i2 = 0
        For i1 = 0 To Height - 1
          p = i1 * RowMod1
          j2 = 0
          For j1 = 0 To Width - 1
            P1 = GetPixel16(PixBits(), (p))
            P2 = .GetPixelz(i2 * RowMod2 + Skip2 * j2)
            If AlphaBlending = 1 Then
              Call AlphaBlendPixels
            ElseIf AlphaBlending = -1 Then
              Call InvAlphaBlendPixels
            Else
              Call CombinePixels
            End If
            Call PutPixel16(PixBits(), p, P1)
            j2 = j2 + 1
            If j2 >= Width2 Then j2 = 0
          Next j1
          i2 = i2 + 1
          If i2 >= Height2 Then i2 = 0
        Next i1
    End Select
  End With
  
  Combine2Images = True

End Function

'AlphaBlend Image1 with Image2 where Image3 is an alpha mask - grey or color
'Images 2 and Mask can AutoTile if they are smaller in dimension from Image1

Public Function AlphaBlend3Images(ByRef Width As Long, ByRef Height As Long, _
                                  ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                                  ByVal Image2 As cRVTVBIMG, ByVal ImageA As cRVTVBIMG, _
                                  Optional ByVal AutoTile As Boolean = False) As Boolean
  
  Dim i1 As Long, j1 As Long, i2 As Long, j2 As Long, iA As Long, jA As Long, p As Long
  Dim RowMod1 As Long, Skip1 As Long
  Dim RowMod2 As Long, Skip2 As Long, Width2 As Long, Height2 As Long
  Dim RowModA As Long, SkipA As Long, WidthA As Long, HeightA As Long

  If IntBPP < PIC_16BPP Then Exit Function                'This image must be 16,24 or 32BPP
  If Image2.BitsPerPixel < PIC_8BPP Then Exit Function    'Image2 may be 8,16,24 or 32 BPP
  If ImageA.BitsPerPixel < PIC_8BPP Then Exit Function    'ImageA may be 8,16,24 or 32 BPP
    
  RowMod1 = BMPRowModulo(Width, IntBPP)
  Skip1 = IntBPP \ 8
  
  With Image2
    RowMod2 = .RowModulo
    Skip2 = .BitsPerPixel \ 8
    Width2 = .Width
    Height2 = .Height
  End With
  
  With ImageA
    RowModA = .RowModulo
    SkipA = .BitsPerPixel \ 8
    WidthA = .Width
    HeightA = .Height
  End With
  
  If Not AutoTile Then                        'constrain blend to the smallest area
    If Width2 < Width Then Width = Width2
    If Height2 < Height Then Height = Height2
    If WidthA < Width Then Width = WidthA
    If HeightA < Height Then Height = HeightA
  End If
  
  Select Case IntBPP
    Case PIC_24BPP, PIC_32BPP:  'inline for speed
      i2 = 0: iA = 0
      For i1 = 0 To Height - 1
        p = i1 * RowMod1
        j2 = 0: jA = 0
        For j1 = 0 To Width - 1
          P1.Blue = PixBits(p)
          P1.Green = PixBits(p + 1)
          P1.Red = PixBits(p + 2)
          If Skip1 > 3 Then P1.Alpha = PixBits(p + 3)
          P2 = Image2.GetPixelz(i2 * RowMod2 + Skip2 * j2)
          AMask = ImageA.GetPixelz(iA * RowModA + SkipA * jA)
          Call AlphaBlendPixels
          PixBits(p) = P1.Blue: p = p + 1
          PixBits(p) = P1.Green: p = p + 1
          PixBits(p) = P1.Red: p = p + 1
          If Skip1 > 3 Then PixBits(p) = P1.Alpha: p = p + 1
          j2 = j2 + 1
          If j2 >= Width2 Then j2 = 0
          jA = jA + 1
          If jA >= WidthA Then jA = 0
        Next j1
        i2 = i2 + 1
        If i2 >= Height2 Then i2 = 0
        iA = iA + 1
        If iA >= HeightA Then iA = 0
      Next i1
      
    Case PIC_16BPP:     'quite a bit slower
      i2 = 0: iA = 0
      For i1 = 0 To Height - 1
        p = i1 * RowMod1
        j2 = 0: jA = 0
        For j1 = 0 To Width - 1
          P1 = GetPixel16(PixBits(), (p))
          P2 = Image2.GetPixelz(i2 * RowMod2 + Skip2 * j2)
          AMask = ImageA.GetPixelz(iA * RowModA + SkipA * jA)
          Call AlphaBlendPixels
          Call PutPixel16(PixBits(), p, P1)
          j2 = j2 + 1
          If j2 >= Width2 Then j2 = 0
          jA = jA + 1
          If jA >= WidthA Then jA = 0
        Next j1
        i2 = i2 + 1
        If i2 >= Height2 Then i2 = 0
        iA = iA + 1
        If iA >= HeightA Then iA = 0
      Next i1
  End Select
  
  AlphaBlend3Images = True

End Function

'Concatenate Two Images, or rather join a second image to the first at the named edge
'Both Images Must be 24 or 32 BPP, alpha is preserved
Public Function SpliceImage(ByRef Width As Long, ByRef Height As Long, _
                            ByVal IntBPP As Long, ByRef PixBits() As Byte, _
                            ByVal Image2 As cRVTVBIMG, _
                            Optional ByVal WhichEdge As PIC_JOIN_EDGES = PJE_RIGHT) As Boolean

  Dim i As Long, j As Long, p As Long, q As Long, r As Long
  Dim OldPixBits() As Byte, OldRowMod1 As Long, OldRowMod2 As Long, I2Base As Long
  Dim NewRowMod As Long, NewWidth As Long, NewHeight As Long, Skip As Long

  'must have the same format, and be unmapped
  If IntBPP < PIC_24BPP Then Exit Function
  If Image2.BitsPerPixel < PIC_24BPP Then Exit Function
  If IntBPP <> Image2.BitsPerPixel Then Exit Function
  
  If WhichEdge < PJE_LEFT Or WhichEdge > PJE_BOTTOM Then WhichEdge = PJE_RIGHT

  Select Case WhichEdge
    Case PJE_LEFT, PJE_RIGHT:
      NewWidth = Width + Image2.Width
      
    Case PJE_TOP, PJE_BOTTOM:
      If Width > Image2.Width Then
        NewWidth = Width
      Else
        NewWidth = Image2.Width
      End If
  End Select
      
  NewRowMod = BMPRowModulo(NewWidth, IntBPP)

  Select Case WhichEdge
    Case PJE_LEFT, PJE_RIGHT:
      If Height > Image2.Height Then
        NewHeight = Height
      Else
        NewHeight = Image2.Height
      End If
      
    Case PJE_TOP, PJE_BOTTOM:
      NewHeight = Height + Image2.Height
  End Select
  
  OldPixBits() = PixBits()  'take a copy of this one
  ReDim PixBits(0 To NewHeight * NewRowMod - 1)     'the new image
  
  'now copy in the respective old pixels
  Skip = IntBPP \ 8
  OldRowMod1 = BMPRowModulo(Width, IntBPP)
  OldRowMod2 = BMPRowModulo(Image2.Width, IntBPP)
  I2Base = Image2.PixBitsBasePtr
  
  Select Case WhichEdge
    Case PJE_LEFT:
      For i = 0 To NewHeight - 1
        p = i * OldRowMod2
        q = i * OldRowMod1
        r = i * NewRowMod
        If i < Image2.Height Then Call CopyMemoryRV(PixBits(r), I2Base + p, OldRowMod2)               'Image2
        If i < Height Then Call CopyMemoryRR(PixBits(r + OldRowMod2), OldPixBits(q), OldRowMod1)      'Image
      Next i
      
    Case PJE_TOP:
      For i = 0 To Image2.Height - 1                            'copy in Image2
        p = i * OldRowMod2
        r = i * NewRowMod
        Call CopyMemoryRV(PixBits(r), I2Base + p, OldRowMod2)
      Next i
      For i = 0 To Height - 1                                   'copy in Image
        q = i * OldRowMod1
        r = (i + Image2.Height) * NewRowMod
        Call CopyMemoryRR(PixBits(r), OldPixBits(q), OldRowMod1)
      Next i
    
    Case PJE_RIGHT:
      For i = 0 To NewHeight - 1
        p = i * OldRowMod2
        q = i * OldRowMod1
        r = i * NewRowMod
        If i < Height Then Call CopyMemoryRR(PixBits(r), OldPixBits(q), OldRowMod1)                   'Image
        If i < Image2.Height Then Call CopyMemoryRV(PixBits(r + OldRowMod1), I2Base + p, OldRowMod2)  'Image2
      Next i
      
    Case PJE_BOTTOM:
      For i = 0 To Height - 1                                   'copy in Image
        q = i * OldRowMod1
        r = i * NewRowMod
        Call CopyMemoryRR(PixBits(r), OldPixBits(q), OldRowMod1)
      Next i
      For i = 0 To Image2.Height - 1                            'copy in Image2
        p = i * OldRowMod2
        r = (i + Height) * NewRowMod
        Call CopyMemoryRV(PixBits(r), I2Base + p, OldRowMod2)
      Next i
  End Select
    
  Width = NewWidth
  Height = NewHeight
  SpliceImage = True

End Function


