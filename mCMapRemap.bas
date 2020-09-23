Attribute VB_Name = "mCMAPRemap"
Option Explicit


'- ©2001/2003 Ron van Tilburg - All rights reserved  1.01.2001/Jun-Jul 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'CRemap - Remapping colors by various methods, Simple, Dithered, SED Dithers, also ColorMap Mutations

'these routines make extensive use of HistCmap.bas for three functions,
'  InitColorMappingHistogram()
'  FreeColorMappingHistogram()
'  MatchColorbyHistogram()

Private Type PixelError
  SRErr As Long
  SGErr As Long
  SBErr As Long
  Count As Long
End Type

Private Type GamutParms         'Gamut Range Parameters
  IsFixedRegular As Boolean     'TRUE if Gamut is generated with Fixed Proportions of R,G,B
  nR             As Long        'Nr of Reds - 1
  nG             As Long        'Nr of Greens - 1 (and Nr Greys-1 for Greys)
  nB             As Long        'Nr of Blues - 1
  sR             As Long        'Red Scaler
  sG             As Long        'Green Scaler (and Grey for Greys)
  sB             As Long        'Blue Scaler
  shR            As Long        '(nB+1)*(nG+1)
  shG            As Long        '(nB+1)
  dDiv           As Long        'Divisor
  dHalf          As Long        'Divisor\2
  PostScale      As Boolean     'TRUE if dithered value needs rescaling
  PSScale()      As Long        'Post Scaling Array
End Type

Public Gamut As GamutParms      'Used in All Remap Routines

Public Sub AnalyseGamut(ByVal CMAPMode As IMG_CMAPMODES, ByRef CMap() As RGBA, ByVal NCMapColors As Long)

 Dim i As Long, j As Long, k As Long, r As Long, g As Long, b As Long
 Dim cr(0 To 255) As Integer, cg(0 To 255) As Integer, cb(0 To 255) As Integer

  With Gamut
    .PostScale = False
    
    Select Case CMAPMode
      Case PIC_FIXED_CMAP_C8, PIC_FIXED_CMAP_C16, PIC_FIXED_CMAP_C32, PIC_FIXED_CMAP_C64, _
           PIC_FIXED_CMAP_C128, PIC_FIXED_CMAP_C256, PIC_FIXED_CMAP_MS256, PIC_FIXED_CMAP_INET:
        .IsFixedRegular = True
        
      Case PIC_FIXED_VMAP_C512:    'for mapping 16bpp,24bpp and 32bpp to 9bpp 3,3,3
        .IsFixedRegular = True
        .nR = 7:     .nG = 7:    .nB = 7
        .sR = 1793:  .sG = 1793: .sB = 1793
        .shR = 1024: .shG = 32
        .dDiv = 65536: .dHalf = 32768
        .PostScale = True
        ReDim .PSScale(0 To 7)
        For i = 0 To 7
          .PSScale(i) = (i * 62& + 7&) \ 14&
        Next i
        Exit Sub
        
      Case PIC_FIXED_VMAP_C4K:    'for mapping 16bpp,24bpp and 32bpp to 12bpp 4,4,4
        .IsFixedRegular = True
        .nR = 15:    .nG = 15:   .nB = 15
        .sR = 3841:  .sG = 3841: .sB = 3841
        .shR = 1024: .shG = 32
        .dDiv = 65536: .dHalf = 32768
        .PostScale = True
        ReDim .PSScale(0 To 15)
        For i = 0 To 15
          .PSScale(i) = (i * 62& + 15&) \ 30&
        Next i
        Exit Sub
        
      Case PIC_FIXED_VMAP_C32K:    'for mapping 16bpp,24bpp and 32bpp to 15bpp 5,5,5
        .IsFixedRegular = True
        .nR = 31:    .nG = 31:   .nB = 31
        .sR = 7937:  .sG = 7937: .sB = 7937
        .shR = 1024: .shG = 32
        .dDiv = 65536: .dHalf = 32768
        Exit Sub
        
      Case PIC_FIXED_VMAP_C64K:    'for mapping 24bpp and 32bpp to 16bpp 5,6,5
        .IsFixedRegular = True
        .nR = 31:    .nG = 63:    .nB = 31
        .sR = 7937:  .sG = 16129: .sB = 7937
        .shR = 2048: .shG = 32
        .dDiv = 65536: .dHalf = 32768
        Exit Sub
        
      Case PIC_FIXED_CMAP_GREY:
        .IsFixedRegular = False
        .nG = NCMapColors - 1
        .sG = 256 * .nG + 1
        .dDiv = 65536: .dHalf = 32768
        Exit Sub
        
      Case Else:    'BW,C4 or Variable
        .IsFixedRegular = False
    End Select
    
    For i = 0 To NCMapColors - 1    'count the incidences of a given value
      r = CMap(i).Red
      g = CMap(i).Green
      b = CMap(i).Blue
      cr(r) = cr(r) + 1
      cg(g) = cr(g) + 1
      cb(b) = cb(b) + 1
    Next i
   
    .nR = -1
    .nG = -1
    .nB = -1

    For i = 0 To 255                    'now tally and collect, distinct values
      If cr(i) <> 0 Then .nR = .nR + 1
      If cg(i) <> 0 Then .nG = .nG + 1
      If cb(i) <> 0 Then .nB = .nB + 1
    Next i

    If .nR < 1 Then .nR = 1 'nr of reds-1
    If .nG < 1 Then .nG = 1 'nr of greens-1
    If .nB < 1 Then .nB = 1 'nr of blues-1

    If FSet(CMAPMode, PIC_FIXED_CMAP) Then    'we dither with 257 values on Fixed Maps
      .sR = 256 * .nR + 1
      .sG = 256 * .nG + 1
      .sB = 256 * .nB + 1
      .dDiv = 65536
    Else                                      'and 64 values on variable colour Maps
      .sR = 64 * .nR + 1
      .sG = 64 * .nG + 1
      .sB = 64 * .nB + 1
      .dDiv = 16384
    End If
    
    .dHalf = .dDiv \ 2
    .shG = .nB + 1
    .shR = .shG * (.nG + 1)
  End With

End Sub

'======================= SIMPLE COLOUR REPLACEMENT =============================================================

'This routine takes a new colormap, and the original PixBit array and remaps the array a pixel at a time
'to the best fitting color by direct substitution. This should really be dithered to minimise the
'error propagation of approximate colormaps. There may be unused (unmappable) colors in the colormap
'the original Pixel array will be resized to the equivalent of a 8BPP array, correctly size without padding
'every entry will be the index into the map

Public Sub SimpleMapColors(ByVal Width As Long, ByVal Height As Long, _
                           ByRef PixBits() As Byte, _
                           ByVal IntBPP As Long, _
                           ByRef CMap() As RGBA, _
                           ByVal NCMapColors As Long, _
                           ByVal CMAPMode As Long)

 Dim x As Long, y As Long, z As Long, w As Long, Skip As Long, RowMod As Long, NewRowMod As Long
 Dim r As Long, g As Long, b As Long, wColor As Long, p As Long, CIndex As Long
 Dim PixErr() As PixelError

  If IntBPP < PIC_16BPP Then Exit Sub           'will only work for unmapped DIBs

  ' Initialize Remapping variables
  Call AnalyseGamut(CMAPMode, CMap(), NCMapColors)                  'Set Color Gamut Parms
  If FClr(CMAPMode, PIC_GMAP) Then
    If Not Gamut.IsFixedRegular Then Call InitColorMappingHistogram(CMap(), NCMapColors)
  End If

  If FClr(CMAPMode, PIC_FIXED_CMAP) Then
    ReDim PixErr(0 To NCMapColors - 1)                              'Init Sum of Error arrays
  End If

  Skip = IntBPP \ 8                                           'size of a pixel in bytes
  RowMod = (UBound(PixBits) - LBound(PixBits) + 1) \ Height         'the byte width of a row
  
  If FSet(CMAPMode, PIC_VMAP) Then
    NewRowMod = BMPRowModulo(Width, PIC_16BPP)
  Else
    NewRowMod = BMPRowModulo(Width, PIC_8BPP)
  End If
  
  For y = 0 To Height - 1                                           'assume bmap is right way up
    z = y * RowMod
    p = y * NewRowMod                                               'this is where we put the new byte
    w = z + Skip * (Width - 1)
    For x = z To w Step Skip        'pixel 0,1,2,3 in a row

      Select Case IntBPP
       Case PIC_24BPP:                                'for 24-bit DIBs
        b = PixBits(x)
        g = PixBits(x + 1)
        r = PixBits(x + 2)
       
       Case PIC_16BPP:                                'for 16-bit DIBs (rescale to 0..255)
        wColor = PixBits(x) + PixBits(x + 1) * 256&
        b = BMP31Scale((wColor And &H1F&))
        g = BMP31Scale((wColor And &H3E0&) \ 32&)
        r = BMP31Scale((wColor And &H7C00&) \ 1024&)
       
       Case PIC_32BPP:                                'for 32-bit DIBs
        b = PixBits(x)
        g = PixBits(x + 1)
        r = PixBits(x + 2)
      End Select

      With Gamut
        If FSet(CMAPMode, PIC_GMAP) Then
          CIndex = ((.sG * RGBtoGrey(r, g, b) + .dHalf) \ .dDiv) 'calculate the grey
        
        ElseIf Not .IsFixedRegular Then
          CIndex = MatchColorbyHistogram(r, g, b)                   'find the nearest color to the given color
          
        Else                                                        'calculate the index
          r = (.sR * r + .dHalf) \ .dDiv
          g = (.sG * g + .dHalf) \ .dDiv
          b = (.sB * b + .dHalf) \ .dDiv
          
          If .PostScale Then
            r = .PSScale(r): g = .PSScale(g): b = .PSScale(b)
          End If
          
          CIndex = r * .shR + g * .shG + b
        End If
      End With
      
      If FClr(CMAPMode, PIC_FIXED_CMAP) Then
        With PixErr(CIndex)
          .SRErr = .SRErr + r - CMap(CIndex).Red    'Error Sums per color
          .SGErr = .SGErr + g - CMap(CIndex).Green
          .SBErr = .SBErr + b - CMap(CIndex).Blue
          .Count = .Count + 1
        End With
      End If

      If FSet(CMAPMode, PIC_VMAP) Then
        PixBits(p) = (CIndex And &HFF&)
        p = p + 1
        PixBits(p) = (CIndex And &HFF00&) \ 256&
        p = p + 1
      Else
        PixBits(p) = CIndex
        p = p + 1
      End If
    Next x
  Next y

  'OK everything is now mapped so lets resize the PixBits array
  ReDim Preserve PixBits(0 To NewRowMod * Height - 1) As Byte

  If FClr(CMAPMode, PIC_GMAP) Then
    If Not Gamut.IsFixedRegular Then Call FreeColorMappingHistogram
  End If

  'After mapping we now adjust the color map to improve overall pixel color matching
  If FClr(CMAPMode, PIC_FIXED_CMAP) And NCMapColors > 32 Then
    Call PostAdjustCMap(CMap(), NCMapColors, PixErr())
  End If

  '  MsgBox "OK SimpleRemap"

End Sub

'==============================================================================================================
'==================================== GENERATING COLOUR MAPS ==================================================
'==============================================================================================================
'A pass through here will independently correct the colours found by Quantizing by summing the mapping errors
Public Sub OptimiseColors(ByVal Width As Long, ByVal Height As Long, _
                          ByRef PixBits() As Byte, _
                          ByVal IntBPP As Long, _
                          ByRef CMap() As RGBA, _
                          ByVal NCMapColors As Long, _
                          ByVal CMAPMode As Long)

 Dim x As Long, y As Long, z As Long, w As Long, Skip As Long, RowMod As Long, NewRowMod As Long
 Dim r As Long, g As Long, b As Long, wColor As Long, CIndex As Long, v As Long, PSkip As Long
 Dim PixErr() As PixelError

  If IntBPP < PIC_16BPP Then Exit Sub                    'will only work for unmapped DIBs

  If FSet(CMAPMode, PIC_FIXED_CMAP) Then Exit Sub                'and will only work for variable CMaps

  ' Initialize Remapping variables
  Call InitColorMappingHistogram(CMap(), NCMapColors)

  'Init Sum of Error arrays
  ReDim PixErr(0 To NCMapColors - 1)

  ' Scan the PixBits() and build the octree
  Skip = (IntBPP \ 8)                                             'size of a pixel in bytes
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)           'the byte width of a row
  
  PSkip = (Width * Height) \ 81920        'Effectively read as if size was 320*256
  If (PSkip And 1) = 0 Then PSkip = PSkip - 1
  If PSkip < 1 Then PSkip = 1

  For y = 0 To Height - 1                                               'assume bmap is right way up
    z = y * RowMod
    w = z + Skip * (Width - 1)
    For x = z To w Step Skip        'pixel 0,1,2,3 in a row

      If v = 0 Then
        Select Case IntBPP
         Case PIC_24BPP:                                ' for 24-bit DIBs
          b = PixBits(x)
          g = PixBits(x + 1)
          r = PixBits(x + 2)
         
         Case PIC_16BPP:                                'for 16-bit DIBs (rescale to 0..255)
          wColor = PixBits(x) + PixBits(x + 1) * 256&
          b = BMP31Scale((wColor And &H1F&))
          g = BMP31Scale((wColor And &H3E0&) \ 32&)
          r = BMP31Scale((wColor And &H7C00&) \ 1024&)
  
         Case PIC_32BPP:                                ' for 32-bit DIBs
          b = PixBits(x)
          g = PixBits(x + 1)
          r = PixBits(x + 2)
        End Select
  
        CIndex = MatchColorbyHistogram(r, g, b)     'find the nearest color to the given color
        With PixErr(CIndex)
          .SRErr = .SRErr + r - CMap(CIndex).Red    'Error Sums per color
          .SGErr = .SGErr + g - CMap(CIndex).Green
          .SBErr = .SBErr + b - CMap(CIndex).Blue
          .Count = .Count + 1
        End With
        v = PSkip
      Else
        v = v - 1
      End If
    Next x
  Next y

  Call FreeColorMappingHistogram
  Call PostAdjustCMap(CMap(), NCMapColors, PixErr())

End Sub

'After mapping (or counting) we have a measure of the nett mapping errors by colour
'Here we can postoptimize the palette by adjusting thos colors by an average of the nett mapping error
Private Sub PostAdjustCMap(ByRef CMap() As RGBA, ByVal NCMapColors As Long, ByRef PixErr() As PixelError)

 Dim i As Long, r As Long, g As Long, b As Long, e As Long

  For i = 0 To NCMapColors - 1
    With PixErr(i)
      If .Count = 0 Then .Count = 1
      r = CMap(i).Red + .SRErr \ .Count
      g = CMap(i).Green + .SGErr \ .Count
      b = CMap(i).Blue + .SBErr \ .Count
    End With
    If r < 0 Then r = 0 Else If r > 255 Then r = 255
    If g < 0 Then g = 0 Else If g > 255 Then g = 255
    If b < 0 Then b = 0 Else If b > 255 Then b = 255
    CMap(i).Red = r
    CMap(i).Green = g
    CMap(i).Blue = b
    CMap(i).Alpha = RGBtoGrey(r, g, b)
  Next i
  
  If NCMapColors = 4 Then     'we attempt to adjust the colours towards the Fixed Poles
    Call MoveTowards(CMap(0), vbBlack)
    Call MoveTowards(CMap(1), vbCyan)
    Call MoveTowards(CMap(2), vbMagenta)
    Call MoveTowards(CMap(3), vbYellow)
  ElseIf NCMapColors = 8 Then
    Call MoveTowards(CMap(0), vbBlack)
    Call MoveTowards(CMap(1), vbBlue)
    Call MoveTowards(CMap(2), vbGreen)
    Call MoveTowards(CMap(3), vbCyan)
    Call MoveTowards(CMap(4), vbRed)
    Call MoveTowards(CMap(5), vbMagenta)
    Call MoveTowards(CMap(6), vbYellow)
    Call MoveTowards(CMap(7), vbWhite)
  End If
End Sub

Private Sub MoveTowards(ByRef c As RGBA, ByVal Pole As Long)
  Dim pm As Long, px As Long
  
  With c
    'find min
    pm = .Red
    If .Green < pm Then pm = .Green
    If .Blue < pm Then pm = .Blue
    'find max
    px = .Red
    If .Green > px Then px = .Green
    If .Blue > px Then px = .Blue
    px = 255 - px
    
    Select Case Pole
      Case vbBlack:   .Red = .Red - pm: .Green = .Green - pm: .Blue = .Blue - pm
      Case vbRed:     .Red = .Red + px: .Green = .Green - pm: .Blue = .Blue - pm
      Case vbGreen:   .Red = .Red - pm: .Green = .Green + px: .Blue = .Blue - pm
      Case vbYellow:  .Red = .Red + px: .Green = .Green + px: .Blue = .Blue - pm
      Case vbBlue:    .Red = .Red - pm: .Green = .Green - pm: .Blue = .Blue + px
      Case vbCyan:    .Red = .Red - pm: .Green = .Green + px: .Blue = .Blue + px
      Case vbMagenta: .Red = .Red + px: .Green = .Green - pm: .Blue = .Blue + px
      Case vbWhite:   .Red = .Red + px: .Green = .Green + px: .Blue = .Blue + px
    End Select
    .Alpha = RGBtoGrey(.Red, .Green, .Blue)
  End With
End Sub

'a generic color to grey function
Public Function RGBtoGrey(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long

  RGBtoGrey = (19595& * r + 38470 * g + 7471& * b) \ 65536    'Luminance

End Function

'Modified YUV - shifted to 0..255 intervals
Public Sub RGBtoYUV(ByRef y As Long, ByRef u As Long, ByRef v As Long, _
                    ByVal r As Long, ByVal g As Long, ByVal b As Long)

                                                                  'The JFIF RGB to YUV Matrix for $00010000 = 1.0
  y = (19595& * r + 38470 * g + 7471& * b) \ 65536                '[Y]   [ 19595   38469   7471][R]
  u = (11056& * r - 21712& * g + 32768 * b + 5536560) \ 65536     '[U] = [-11056  -21712  32768][G]
  v = (32768 * r - 27440& * g - 5328& * b + 8355840) \ 63356      '[V]   [ 32768  -27440  -5328][B]

End Sub

Public Sub YUVtoRGB(ByRef r As Long, ByRef g As Long, ByRef b As Long, _
                    ByVal y As Long, ByVal u As Long, ByVal v As Long)

  If u > 127 Then u = u - 128
  If v > 127 Then v = v - 128
                                           'The inverse of the JFIF RGB to YUV Matrix for $00010000 = 1.0
  r = 65496 * y + 91880 * v                '[Y]   [65496        0   91880][R]
  g = 65533 * y - 22580& * u - 46799 * v   '[U] = [65533   -22580  -46799][G]
  b = 65537 * y + 116128 * u - 8& * v      '[V]   [65537   116128      -8][B]

End Sub

Public Sub CMaptoGrey(ByRef CMap() As RGBA)  'change the RGB palette to CIY greys

 Dim i As Long, z As Long

  For i = LBound(CMap) To UBound(CMap)
    z = RGBtoGrey(CLng(CMap(i).Red), CLng(CMap(i).Green), CLng(CMap(i).Blue))
    CMap(i).Alpha = z
    CMap(i).Red = z
    CMap(i).Green = z
    CMap(i).Blue = z
  Next i

End Sub

Public Sub GenFixedMap(ByRef CMAPMode As Long, ByRef CMap() As RGBA, ByRef NCMapColors As Long, ByRef BitsPerPixel As Long)

 Dim i As Long, fcm() As Variant
        
  If FSet(CMAPMode, PIC_GMAP) Then
    Call GenGreyMap(CMap(), NCMapColors, BitsPerPixel)
  Else
    Select Case CMAPMode
     Case PIC_FIXED_CMAP_BW:      'Turned into GREY(2)
      BitsPerPixel = 1
      NCMapColors = 2
      CMAPMode = PIC_FIXED_CMAP_GREY
      Call GenGreyMap(CMap(), NCMapColors, BitsPerPixel)
  
     Case PIC_FIXED_CMAP_C4:
      BitsPerPixel = 2
      NCMapColors = 4
      ReDim CMap(0 To 3)
      fcm() = Array(0, 0, 0, 0, 255, 255, 255, 0, 255, 255, 255, 0) 'KCMY
      For i = 0 To 11 Step 3
        CMap(i \ 3).Red = fcm(i)
        CMap(i \ 3).Green = fcm(i + 1)
        CMap(i \ 3).Blue = fcm(i + 2)
      Next i
  
     Case PIC_FIXED_CMAP_C8:
      BitsPerPixel = 3
      NCMapColors = 8
      Call GenCMap(CMap(), NCMapColors, 2, 2, 2)  '8
  
     Case PIC_FIXED_CMAP_C16:
      BitsPerPixel = 4
      NCMapColors = 16
      Call GenCMap(CMap(), NCMapColors, 2, 4, 2) '16
  '    Call GenCMap(CMap(), NCMapColors, 2, 3, 2, 64, 128, 192, 224) '12 +4
  
     Case PIC_FIXED_CMAP_VGA:
      BitsPerPixel = 4
      NCMapColors = 16
      ReDim CMap(0 To 15)
      fcm() = Array(0, 0, 0, 128, 0, 0, 0, 128, 0, 128, 128, 0, _
                    0, 0, 128, 128, 0, 128, 0, 128, 128, _
                    192, 192, 192, 128, 128, 128, _
                    255, 0, 0, 0, 255, 0, 255, 255, 0, _
                    0, 0, 255, 255, 0, 255, 0, 255, 255, 255, 255, 255) 'KRGYBMC/2,3W/4,W/2,RGYBMCW
      For i = 0 To 47 Step 3
        CMap(i \ 3).Red = fcm(i)
        CMap(i \ 3).Green = fcm(i + 1)
        CMap(i \ 3).Blue = fcm(i + 2)
      Next i
  
     Case PIC_FIXED_CMAP_C32:
      BitsPerPixel = 5
      NCMapColors = 32
  '    Call GenCMap(CMap(), NCMapColors, 4, 4, 2) '32
      Call GenCMap(CMap(), NCMapColors, 3, 3, 3) '27
  '    Call GenCMap(CMap(), NCMapColors, 3, 3, 3, 64, 96, 160, 192, 224) '27+5
  
     Case PIC_FIXED_CMAP_C64:
      BitsPerPixel = 6
      NCMapColors = 64
      Call GenCMap(CMap(), NCMapColors, 4, 4, 4)    '64
  
     Case PIC_FIXED_CMAP_C128:
      BitsPerPixel = 7
      NCMapColors = 128
   '   Call GenCMap(CMap(), NCMapColors, 5, 5, 5, 96, 160, 224) '125+3
      Call GenCMap(CMap(), NCMapColors, 5, 5, 5)  '125
  
     Case PIC_FIXED_CMAP_INET:
      BitsPerPixel = 8
      NCMapColors = 256
      Call GenCMap(CMap(), NCMapColors, 6, 6, 6)   '216
  
     Case PIC_FIXED_CMAP_C256:
      BitsPerPixel = 8
      NCMapColors = 256
      Call GenCMap(CMap(), NCMapColors, 6, 7, 6)   '252
  
     Case PIC_FIXED_CMAP_MS256:
      BitsPerPixel = 8
      NCMapColors = 256
      Call GenCMap(CMap(), NCMapColors, 8, 8, 4)   '256
      
     Case Else:
      BitsPerPixel = 8
      NCMapColors = 256
      Call GenCMap(CMap(), NCMapColors, 6, 7, 6)   '252 'DEFAULT
    End Select
  End If
End Sub

'Copy as much of UserCmap into CMap that will fit inside
Public Sub GenUserCMap(ByRef CMap() As RGBA, ByRef NCMapColors As Long, ByRef UserCMap() As RGBA)
  Dim size As Long
  
  size = UBound(UserCMap) - LBound(UserCMap) + 1
  If NCMapColors = 0 Then NCMapColors = size
  If size > 0 Then
    ReDim CMap(0 To NCMapColors - 1)
    If size <= NCMapColors Then
      Call CopyMemoryRR(CMap(0), UserCMap(LBound(UserCMap)), 4 * size)
    Else
      Call CopyMemoryRR(CMap(0), UserCMap(LBound(UserCMap)), 4 * NCMapColors)
    End If
  End If
  
End Sub

Private Sub GenGreyMap(ByRef CMap() As RGBA, ByRef NCMapColors As Long, ByRef BitsPerPixel As Long)

 Dim i As Long, z As Long

  NCMapColors = 2 ^ BitsPerPixel
  ReDim CMap(0 To NCMapColors - 1)

  'Fill Colour Map with Greys
  For i = 0 To NCMapColors - 1
    z = (256 * i) \ (NCMapColors - 1)
    If z > 255 Then z = 255
    CMap(i).Alpha = z
    CMap(i).Red = z
    CMap(i).Green = z
    CMap(i).Blue = z
  Next i

End Sub

Private Sub GenCMap(ByRef CMap() As RGBA, ByVal NCMapColors As Long, _
                    ByVal nR As Integer, ByVal nG As Integer, ByVal nB As Integer, _
                    ParamArray Greys())

 Dim p As Integer, r As Integer, g As Integer, b As Integer, sR As Integer, sG As Integer, sB As Integer

  ReDim CMap(0 To NCMapColors - 1)
  nR = nR - 1
  nG = nG - 1
  nB = nB - 1
  p = 0
  For r = 0 To nR
    sR = (r * 256) \ nR
    If sR > 255 Then sR = 255
    For g = 0 To nG
      sG = (g * 256) \ nG
      If sG > 255 Then sG = 255
      For b = 0 To nB
        sB = (b * 256) \ nB
        If sB > 255 Then sB = 255
        CMap(p).Red = sR
        CMap(p).Green = sG
        CMap(p).Blue = sB
        CMap(p).Alpha = RGBtoGrey(sR, sG, sB)
        p = p + 1
      Next b
    Next g
  Next r

  For g = LBound(Greys) To UBound(Greys)
    CMap(p).Alpha = Greys(g)
    CMap(p).Red = Greys(g)
    CMap(p).Green = Greys(g)
    CMap(p).Blue = Greys(g)
    p = p + 1
  Next g

End Sub

Public Function ShrinkCMap(ByRef PixBits() As Byte, ByVal PixelWidth As Long, _
                           ByRef CMap() As RGBA, ByVal BitsPerPixel As Long) As Long

 Dim i As Long, j As Long, k As Long, NC As Long, nnc As Long
 Dim Idx() As Integer, cc() As Long

  ShrinkCMap = BitsPerPixel
  NC = UBound(CMap) - LBound(CMap) + 1            'the current number of colours in CMAP
  ReDim Idx(0 To NC - 1) As Integer               'the indices of old and new pixels
  ReDim cc(0 To NC - 1) As Long                   'the colour count of colours used

  'count the number of each sort of pixel
  For i = LBound(PixBits) To UBound(PixBits)  'PixelWidth is 4 or 8 only
    If PixelWidth = 8 Then
      j = PixBits(i)
      cc(j) = cc(j) + 1
     Else
      j = (PixBits(i) And &HF0&) \ 16&
      cc(j) = cc(j) + 1
      j = (PixBits(i) And &HF&)
      cc(j) = cc(j) + 1
    End If
  Next i

  'now set up the idx array, newcolor=idx(oldcolor)
  j = 0
  For i = 0 To NC - 1
    If cc(i) > 0 Then     'its been used
      Idx(i) = j          'and is kept
      j = j + 1
     Else
      Idx(i) = -1
    End If
  Next i

  nnc = j
  If nnc <= (NC \ 2) Then     'we did reduce the number of colors enough to gain something
    For i = 0 To NC - 1       'so adjust the cmap
      j = Idx(i)
      k = i
      If j >= 0 And j < k Then       'move it up
        CMap(j) = CMap(k)
      End If
    Next i
    i = 1
    j = 0              'find the new palette size
    Do
      j = j + 1
      i = i + i
    Loop Until i > nnc
    ReDim Preserve CMap(0 To i - 1)    '2^j colours
    ShrinkCMap = j                             'New Bits Per Pixel

    'now the jolly task of remapping the Pixels
    For i = LBound(PixBits) To UBound(PixBits)  'PixelWidth is 4 or 8 only
      If PixelWidth = 8 Then
        PixBits(i) = Idx(PixBits(i))
       Else
        j = Idx((PixBits(i) And &HF0&) \ 16&)
        k = Idx((PixBits(i) And &HF&))
        PixBits(i) = j * 16 + k           'the new value
      End If
    Next i
  End If

End Function

'==========================  CONVERSION TO GREY OF ALL COLOURS (SOME GREYS MAY NOT TURN UP) ===================

'This routine takes a BMP and turns all its pixels into grey shades
Public Sub MakePixelsGrey(ByVal Width As Long, ByVal Height As Long, _
                          ByRef PixBits() As Byte, ByVal IntBPP As Long)

 Dim x As Long, y As Long, z As Long, w As Long, Skip As Long, RowMod As Long
 Dim r As Long, g As Long, b As Long, wColor As Long, p As Long, v As Long

  If IntBPP < PIC_16BPP Then Exit Sub           'will only work for unmapped DIBs

  Skip = (IntBPP \ 8)                                             'size of a pixel in bytes
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)           'the byte width of a row

  For y = 0 To Height - 1                                               'assume bmap is right way up
    p = y * RowMod                                                      'this is where we put the new byte
    z = y * RowMod
    w = z + Skip * Width - 1
    For x = z To w Step Skip        'pixel 0,1,2,3 in a row

      Select Case IntBPP
       Case PIC_24BPP:                                  'for 24-bit DIBs
        b = PixBits(x)
        g = PixBits(x + 1)
        r = PixBits(x + 2)

       Case PIC_16BPP:                                'for 16-bit DIBs (rescale to 0..255)
        wColor = PixBits(x) + PixBits(x + 1) * 256&
        b = BMP31Scale((wColor And &H1F&))
        g = BMP31Scale((wColor And &H3E0&) \ 32&)
        r = BMP31Scale((wColor And &H7C00&) \ 1024&)

       Case PIC_32BPP:                                  'for 32-bit DIBs
        b = PixBits(x)
        g = PixBits(x + 1)
        r = PixBits(x + 2)
      End Select

      v = RGBtoGrey(r, g, b)                            'The conversion

      Select Case IntBPP
       Case PIC_24BPP:                                  'for 24-bit DIBs
        PixBits(p) = v
        p = p + 1
        PixBits(p) = v
        p = p + 1
        PixBits(p) = v
        p = p + 1

       Case PIC_16BPP:                                  'for 16-bit DIBs
        v = (v * 31&) \ 255&
        wColor = (v * 32& + v) * 32& + v
        PixBits(p) = wColor And &HFF
        p = p + 1
        PixBits(p) = (wColor And &HFF00) \ 256&
        p = p + 1

       Case PIC_32BPP:                                 'for 32-bit DIBs
        PixBits(p) = v
        p = p + 1
        PixBits(p) = v
        p = p + 1
        PixBits(p) = v
        p = p + 1
        PixBits(p) = v  'ALPHA
        p = p + 1
      End Select

    Next x
  Next y

  '  MsgBox "OK MakePixelsGrey"

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jun-07 18:05) 21 + 526 = 547 Lines
