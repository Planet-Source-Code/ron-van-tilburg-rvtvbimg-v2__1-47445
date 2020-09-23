Attribute VB_Name = "mMatDither"
Option Explicit

'mMatDither.bas - Remapping colors by various Matrix Dithers

'- ©2001/2003 Ron van Tilburg - All rights reserved  1.01.2001/Jun-Jul 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'this module makes extensive use of HistCmap.bas for three functions,
'  InitColorMappingHistogram()
'  FreeColorMappingHistogram()
'  MatchColorbyHistogram()
' Gamut in Cremap.bas for Gamut Parameters

'================== MATRIX DITHERING ====================================

Private DMat() As Long  'a dither matrix against which values are compared
Private DModX  As Long  'Width of DMat
Private DModY  As Long  'Height of DMat

'======================= MATRIX DITHERING =====================================================================

'This routine takes a new colormap, and the original PixBit array and remaps the array a pixel at a time
'to the best fitting color by Dithering against a a dither matrix. This minimises the error propagation of
'approximate colormaps. There may be unused (unmappable) colors in the colormap
'the original Pixel array will be resized to the equivalent of a 8BPP array, correctly size without padding
'every entry will be the index into the map

Public Sub MatDitherMapColors(ByVal Width As Long, ByVal Height As Long, _
                              ByRef PixBits() As Byte, _
                              ByVal IntBPP As Long, _
                              ByRef CMap() As RGBA, _
                              ByVal NCMapColors As Long, _
                              ByVal CMAPMode As Long, _
                              ByVal DitherMode As Long)

 Dim x As Long, y As Long, z As Long, w As Long, Skip As Long, RowMod As Long, NewRowMod As Long
 Dim r As Long, g As Long, b As Long, wColor As Long, p As Long
 Dim CIndex As Long, dval As Long, dcol As Long, drow As Long

  If IntBPP < PIC_16BPP Then Exit Sub           'will only work for unmapped DIBs

  ' Initialize Remapping variables
  Call AnalyseGamut(CMAPMode, CMap(), NCMapColors)
  
  If FClr(CMAPMode, PIC_GMAP) Then
    If Not Gamut.IsFixedRegular Then Call InitColorMappingHistogram(CMap(), NCMapColors)
  End If
  
  Call SetDitherMatrix(DitherMode, CMAPMode)
  
  Skip = IntBPP \ 8                                             'size of a pixel in bytes
  RowMod = (UBound(PixBits) - LBound(PixBits) + 1) \ Height           'the byte width of a row
  
  If FSet(CMAPMode, PIC_VMAP) Then
    NewRowMod = BMPRowModulo(Width, PIC_16BPP)
  Else
    NewRowMod = BMPRowModulo(Width, PIC_8BPP)
  End If

  drow = 0
  For y = 0 To Height - 1                                             'assume bmap is right way up
    z = y * RowMod
    p = y * NewRowMod                                                 'this is where we put the new byte
    w = z + Skip * (Width - 1)
    dcol = 0

    For x = z To w Step Skip                            'pixel 0,1,2,3 in a row

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

      dval = DMat(drow, dcol)                           'find the dither threshold
      
      With Gamut
        If FSet(CMAPMode, PIC_GMAP) Then                    '------------ Grey Dithering ---------------
          CIndex = (.sG * RGBtoGrey(r, g, b) + dval) \ .dDiv
        Else                                                '-------------- Full Color -----------------
          r = (.sR * r + dval) \ .dDiv
          g = (.sG * g + dval) \ .dDiv
          b = (.sB * b + dval) \ .dDiv
                    
          If .IsFixedRegular Then                           'Index is calculable
            If .PostScale Then
              r = .PSScale(r): g = .PSScale(g): b = .PSScale(b)
            End If
            CIndex = r * .shR + g * .shG + b
          Else                                              'Lookup nearest color
            r = (r * 256) \ .nR: If r > 255 Then r = 255
            g = (g * 256) \ .nG: If g > 255 Then g = 255
            b = (b * 256) \ .nB: If b > 255 Then b = 255
            CIndex = MatchColorbyHistogram(r, g, b)
          End If
        End If
      End With
      
      If FSet(CMAPMode, PIC_VMAP) Then
        PixBits(p) = (CIndex And &HFF&)
        p = p + 1
        PixBits(p) = (CIndex And &HFF00&) \ 256&
        p = p + 1
      Else
        PixBits(p) = CIndex
        p = p + 1
      End If

      dcol = dcol + 1
      If dcol = DModX Then dcol = 0
    Next x

    drow = drow + 1
    If drow = DModY Then drow = 0
  Next y

  'OK everything is now mapped so lets resize the PixBits array
  ReDim Preserve PixBits(0 To NewRowMod * Height - 1) As Byte

  If FClr(CMAPMode, PIC_GMAP) Then
    If Not Gamut.IsFixedRegular Then Call FreeColorMappingHistogram
  End If
  '  MsgBox "OK DitheredRemap"

End Sub

'------------------------- Setup of Dither Matrices and Gamut scaling -----------------------------------------
'See mMatDitherSlow.bas for the code that made the resfiles

Private Sub SetDitherMatrix(ByVal WhichDither As Long, ByVal CMAPMode As Long)

 Dim i As Long, j As Long, k As Long, m As Long, dd() As Byte
  
  dd = LoadResData(WhichDither Or &H4000, "CUSTOM")
  k = Sqr(UBound(dd) + 1)
  ReDim DMat(0 To k - 1, 0 To k - 1)
  If FSet(CMAPMode, PIC_FIXED_CMAP) Then             'Fixed=256, Variable=64
    For i = 0 To UBound(dd)
      DMat(i \ k, i Mod k) = 256& * CLng(dd(i))
    Next i
  Else
    For i = 0 To UBound(dd)
      DMat(i \ k, i Mod k) = 64& * (CLng(dd(i)) \ 4)
    Next i
  End If
  DModX = k
  DModY = k
End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jun-07 18:05) 28 + 254 = 282 Lines
