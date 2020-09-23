Attribute VB_Name = "mSEDDither"
Option Explicit

'mSEDDither.bas - Remapping colors by various Serpentine Error Diffusion Dithers

'- ©2001/2003 Ron van Tilburg - All rights reserved  1.01.2001/Jun-Jul 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'this module makes extensive use of HistCmap.bas for three functions,
'  InitColorMappingHistogram()
'  FreeColorMappingHistogram()
'  MatchColorbyHistogram()
' and AnalyseGamut() in CRemap.bas

Private Type PixelError
  RSum As Long
  GSum As Long
  BSum As Long
End Type

Private Type SEDCoeffs
  w1 As Long
  w2 As Long
  w3 As Long
  w4 As Long
End Type

Private SEDCoeffs(-255 To 255) As SEDCoeffs   'Error propagation arrays for SED Dithering

'==============================================================================================================
'================== SERPENTINE ERROR DIFFUSION ================================================================
'==============================================================================================================

'This routine takes a new colormap, and the original PixBit array and remaps the array a pixel at a time
'to the best fitting color by Serpentine Error Diffusion Dithering. This minimises the error propagation of
'approximate colormaps. There may be unused (unmappable) colors in the colormap
'the original Pixel array will be resized to the equivalent of a 8BPP array, correctly size without padding
'every entry will be the index into the map - this routine is fairly involved but delivers excellent results

Public Sub SEDDitherMapColors(ByVal Width As Long, ByVal Height As Long, _
                              ByRef PixBits() As Byte, _
                              ByVal IntBPP As Long, _
                              ByRef CMap() As RGBA, _
                              ByVal NCMapColors As Long, _
                              ByVal CMAPMode As Long, _
                              ByVal DitherMode As Long)

 Const SED_SCALE As Long = 4096&  'errors are scaled by this much to limit roundoff error (was 128)

 Dim x As Long, Skip As Long, RowMod As Long, NewRowMod As Long

 Dim r As Long, g As Long, b As Long, wColor As Long
 Dim sR As Long, sG As Long, sB As Long, aw As Long, p As Long

 Dim Col0 As Long, Col1 As Long, Col2 As Long, LimitCol As Long
 Dim Row As Long, CIndex As Long, SED_LtoR As Boolean

 Dim DelR As SEDCoeffs, DelG As SEDCoeffs, DelB As SEDCoeffs    'scaled errors
 Dim CurrErr() As PixelError, NextErr() As PixelError           'accumulators

  If IntBPP < PIC_16BPP Then Exit Sub           'will only work for unmapped DIBs

  Call GenSEDErrLUT(DitherMode, SED_SCALE)                 'Initialize Remapping Coefficients Lookup Tables
  Call AnalyseGamut(CMAPMode, CMap(), NCMapColors)         'Set Color Gamut Parms
  
  If FClr(CMAPMode, PIC_GMAP) Then
    If Not Gamut.IsFixedRegular Then Call InitColorMappingHistogram(CMap(), NCMapColors) 'Setup Inverse Colormap
  End If

  ' Establish the amounts needed to scan the bitmap
  Skip = IntBPP \ 8                                               'size of a pixel in bytes
  RowMod = (UBound(PixBits) - LBound(PixBits) + 1) \ Height             'the byte width of a row
  
  If FSet(CMAPMode, PIC_VMAP) Then
    NewRowMod = BMPRowModulo(Width, PIC_16BPP)
  Else
    NewRowMod = BMPRowModulo(Width, PIC_8BPP)
  End If

  aw = Width + 2              'Initialize Error Diffusion error vectors.
  ReDim CurrErr(0 To aw - 1)

  SED_LtoR = True             'ie moving Left to Right, FALSE is moving right to left

  For Row = 0 To Height - 1

    ReDim NextErr(0 To aw - 1)    'Clear Next Error Arrays

    If SED_LtoR Then
      Col0 = 0
      LimitCol = Width
     Else
      Col0 = Width - 1
      LimitCol = -1
    End If

    'We need to be cunning about this overwriting trick here. We have a serpentine movement LtoR then RtoL
    'On the first row we will at worst (16bit case) use half the bytes of the first original pixel row
    'On the second row we will therefore start writing 1 byte before the last pixel to be used on the first
    'RtoL pass - hence we havent clobbered anything we need later - I just thought Id share that :-)

    p = Row * NewRowMod + Col0       'where we will start putting the row of mapped bytes
    x = Row * RowMod + Col0 * Skip   'where the first pixel will be - assume bmap is right way up

    Do      'remap a pixel
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
      
      Col1 = Col0 + 1
      Col2 = Col0 + 2

      If FSet(CMAPMode, PIC_GMAP) Then          'Short version for Grey Scales, and BW is one of them

        sG = RGBtoGrey(r, g, b) + CurrErr(Col1).GSum \ SED_SCALE
        If sG < 0 Then sG = 0 Else If sG > 255 Then sG = 255
        With Gamut
          CIndex = (.sG * sG + .dHalf) \ .dDiv             'CALCULATE match for this color
        End With
        
        DelG = SEDCoeffs(sG - CMap(CIndex).Green)              'Propagate error terms. (using LUTs)

        If SED_LtoR Then
          CurrErr(Col2).GSum = CurrErr(Col2).GSum + DelG.w1
          NextErr(Col0).GSum = NextErr(Col0).GSum + DelG.w2
          NextErr(Col1).GSum = NextErr(Col1).GSum + DelG.w3

          If DitherMode <> PIC_DITHER_SED3 Then   '4 coeffs
            NextErr(Col2).GSum = NextErr(Col2).GSum + DelG.w4
          End If
         Else
          CurrErr(Col0).GSum = CurrErr(Col0).GSum + DelG.w1
          NextErr(Col2).GSum = NextErr(Col2).GSum + DelG.w2
          NextErr(Col1).GSum = NextErr(Col1).GSum + DelG.w3

          If DitherMode <> PIC_DITHER_SED3 Then   '4 coeffs
            NextErr(Col0).GSum = NextErr(Col0).GSum + DelG.w4
          End If
        End If
       Else  '-------------------------------------------------------------- Full color mapping

        With CurrErr(Col1)
          sR = r + .RSum \ SED_SCALE
          sG = g + .GSum \ SED_SCALE
          sB = b + .BSum \ SED_SCALE
        End With

        If sR < 0 Then sR = 0 Else If sR > 255 Then sR = 255
        If sG < 0 Then sG = 0 Else If sG > 255 Then sG = 255
        If sB < 0 Then sB = 0 Else If sB > 255 Then sB = 255

        If Not Gamut.IsFixedRegular Then
          CIndex = MatchColorbyHistogram(sR, sG, sB)      'find the nearest color to the given color
        Else                                              'calculate the index
          With Gamut
            r = (.sR * sR + .dHalf) \ .dDiv
            g = (.sG * sG + .dHalf) \ .dDiv
            b = (.sB * sB + .dHalf) \ .dDiv
            
            If .PostScale Then
              r = .PSScale(r): g = .PSScale(g): b = .PSScale(b)
            End If
            
            CIndex = r * .shR + g * .shG + b
          End With
        End If

        If FSet(CMAPMode, PIC_VMAP) Then
          If CMAPMode = PIC_FIXED_VMAP_C64K Then
            DelR = SEDCoeffs(sR - (CIndex And &HF800&) \ 256&)        'The New error terms
            DelG = SEDCoeffs(sG - (CIndex And &H7C0&) \ 8&)
            DelB = SEDCoeffs(sB - (CIndex And &H1F&) * 8&)
          Else
            DelR = SEDCoeffs(sR - (CIndex And &H7C00&) \ 128&)        'The New error terms
            DelG = SEDCoeffs(sG - (CIndex And &H3E0&) \ 4&)
            DelB = SEDCoeffs(sB - (CIndex And &H1F&) * 8&)
          End If
        Else
          DelR = SEDCoeffs(sR - CMap(CIndex).Red)                   'The New error terms
          DelG = SEDCoeffs(sG - CMap(CIndex).Green)
          DelB = SEDCoeffs(sB - CMap(CIndex).Blue)
        End If
        
        If SED_LtoR Then                                  'Propagate error terms. (using LUTs)
          With CurrErr(Col2)
            .RSum = .RSum + DelR.w1
            .GSum = .GSum + DelG.w1
            .BSum = .BSum + DelB.w1
          End With

          With NextErr(Col0)
            .RSum = .RSum + DelR.w2
            .GSum = .GSum + DelG.w2
            .BSum = .BSum + DelB.w2
          End With

          With NextErr(Col1)
            .RSum = .RSum + DelR.w3
            .GSum = .GSum + DelG.w3
            .BSum = .BSum + DelB.w3
          End With

          If DitherMode <> PIC_DITHER_SED3 Then   '4th coeff
            With NextErr(Col2)
              .RSum = .RSum + DelR.w4
              .GSum = .GSum + DelG.w4
              .BSum = .BSum + DelB.w4
            End With
          End If
         Else    'RtoL
          With CurrErr(Col0)
            .RSum = .RSum + DelR.w1
            .GSum = .GSum + DelG.w1
            .BSum = .BSum + DelB.w1
          End With

          With NextErr(Col2)
            .RSum = .RSum + DelR.w2
            .GSum = .GSum + DelG.w2
            .BSum = .BSum + DelB.w2
          End With

          With NextErr(Col1)
            .RSum = .RSum + DelR.w3
            .GSum = .GSum + DelG.w3
            .BSum = .BSum + DelB.w3
          End With

          If DitherMode <> PIC_DITHER_SED3 Then   '4th coeff
            With NextErr(Col0)
              .RSum = .RSum + DelR.w4
              .GSum = .GSum + DelG.w4
              .BSum = .BSum + DelB.w4
            End With
          End If
        End If
      End If

      If FSet(CMAPMode, PIC_VMAP) Then
        PixBits(p) = (CIndex And &HFF&)
        p = p + 1
        PixBits(p) = (CIndex And &HFF00&) \ 256&
        p = p + 1
        Col0 = Col0 + 1
        x = x + Skip
      Else
        PixBits(p) = CIndex  'yes we can finally store something, the rest is housekeeping
        If SED_LtoR Then
          Col0 = Col0 + 1
          p = p + 1
          x = x + Skip
         Else
          Col0 = Col0 - 1
          p = p - 1
          x = x - Skip
        End If
      End If
    Loop While Col0 <> LimitCol

    'copy the error arrays up for next line
    CurrErr() = NextErr()

    If FClr(CMAPMode, PIC_VMAP) Then
      SED_LtoR = Not SED_LtoR                   'and reverse direction every second row (not for 32K Mode)
    End If
  Next Row

  'OK everything is now mapped so lets resize the PixBits array
  ReDim Preserve PixBits(0 To NewRowMod * Height - 1) As Byte

  'all dynamically assigned arrays die on exit
  If FClr(CMAPMode, PIC_GMAP) Then
    If Not Gamut.IsFixedRegular Then Call FreeColorMappingHistogram
  End If

  '  MsgBox "OK SEDDitherRemap"

End Sub

'Generation of LU Tables (nett 12% faster than calculating on the fly)
Private Sub GenSEDErrLUT(ByVal DitherMode As Long, ByVal SEDScale As Long)
  Dim dd() As Byte
  
  dd = LoadResData(DitherMode Or &H4000, "CUSTOM")        'Error Diffusion Coefficients
  Call CopyMemoryRR(SEDCoeffs(-255), dd(0), 1 + UBound(dd))
End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jun-07 18:05) 29 + 303 = 332 Lines
