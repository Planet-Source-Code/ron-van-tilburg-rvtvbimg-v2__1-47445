Attribute VB_Name = "mPNMFileIO"
Option Explicit

'mPNMFileIO.bas - PAM, PBM, PGM, and PPM File IO

'- ©2003 Ron van Tilburg - All rights reserved  15.Jun.2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'Adapted to VB from C sources of Netpbm10.15  June 2003 RVT

'Adapted From the Netpbm10.15 release (GPL licenses apply) this is used to read and write NetPBM files
'As PPM files can have a color resolution of up to 3*16 bits, and we can only process and show 24 Bit Color
'we cannot display these correctly natively but we can read the extended formats
'so we shift them down to 24 bit (slowing us down a bit on Load)

'PNM files are always Loaded into Memory as 24bit files, then converted to/from at Load/Save

Private Enum PNM_Type
  PNM_PBM = &H3150    'P1
  PNM_PGM = &H3250    'P2
  PNM_PPM = &H3350    'P3
  PNM_RPBM = &H3450   'P4
  PNM_RPGM = &H3550   'P5
  PNM_RPPM = &H3650   'P6
  PNM_PAM = &H3750    'P7  - Recognised but only read if PBM,PGM or PPM in disguise
End Enum

Private PXType    As Integer  'PX File Type
Private PXRow()   As Byte     'Used to Read/Write/Reformat on the fly
Private PXBytes() As Byte     'Used as a temp string array
Private FileNr    As Long

'============================================ LOAD ===============================================================
'We can Load
'   P1 - .PBM   Plain Format
'   P2 - .PGM   Plain Format 0..255
'   P3 - .PPM   Plain Format 0..Maxval
'   P4 - .RPBM  Binary
'   P5 - .RPGM  Binary
'   P6 - .RPPM  Binary
'   P7 - .PAM   PBM,PGM,PPM, extensions(RGB,RGBA etc etc)

Public Function LoadPNM(Path As String, PicState As Long, _
                        ByRef Width As Long, ByRef Height As Long, _
                        ByRef PixBits() As Byte, ByRef IntBPP As Long, _
                        ByRef CMap() As RGBA, ByRef NCMapColors As Long) As Boolean

 Dim RowModulo As Long, MaxVal As Long, z As Byte

  On Error GoTo LoadPNMFailed

  FileNr = FreeFile()
  Open Path For Binary Access Read As #FileNr

  Get #FileNr, , PXType
  Get #FileNr, , z
  Select Case PXType
   Case PNM_PBM, PNM_PGM, PNM_PPM, PNM_RPBM, PNM_RPGM, PNM_RPPM:     'We have a PNM file we can read
    Width = GetIntVal()
    Height = GetIntVal()

    If PXType <> PNM_PBM And PXType <> PNM_RPBM Then
      MaxVal = GetIntVal()
    End If

    Call SetF(PicState, IS_TOP_TO_BOTTOM)
    IntBPP = PIC_24BPP
    RowModulo = BMPRowModulo(Width, PIC_24BPP)
    NCMapColors = 0

    ReDim PixBits(0 To Height * RowModulo - 1)  'The Target

    Select Case PXType
     Case PNM_PBM:                                                        '1 and 0 as text
      Call ReadPBMBodyPlain(Width, Height, RowModulo, PixBits())

     Case PNM_RPBM:                                                       'binary rows of bits (trailing 0s)
      Call ReadPBMBodyBinary(Width, Height, RowModulo, PixBits())

     Case PNM_PGM:                                                        '0 to 65535 as text
      Call ReadPGMBodyPlain(Width, Height, RowModulo, PixBits(), MaxVal)

     Case PNM_RPGM:                                                       '0 to 65535
      Call ReadPGMBodyBinary(Width, Height, RowModulo, PixBits(), MaxVal)

     Case PNM_PPM:                                                        '0 to 65535 as text
      Call ReadPPMBodyPlain(Width, Height, RowModulo, PixBits(), MaxVal)

     Case PNM_RPPM:                                                       '0 to 65535
      Call ReadPPMBodyBinary(Width, Height, RowModulo, PixBits(), MaxVal)
    End Select

   Case PNM_PAM:   'Dont support this YET
    GoTo LoadPNMFailed    'Cant read this

   Case Else:
    GoTo LoadPNMFailed    'Not a PNM file
  End Select

  LoadPNM = True
  'always come back as 24BPP

LoadPNMFailed:              'We will end up here on all errors whether by data or read action
  Close #FileNr
  On Error GoTo 0

End Function

Private Sub ReadPBMBodyPlain(ByVal Width As Long, ByVal Height As Long, ByVal RowModulo As Long, _
                             ByRef PixBits() As Byte)
                             
 'all values are 1 byte but only the smallest Bit counts (1=BLACK)

 Dim i As Long, j As Long, p As Long, v As Byte

  For i = 0 To Height - 1
    p = i * RowModulo
    j = Width
    Do While j > 0
      Get #FileNr, , v
      If v = &H30 Or v = &H31 Then
        If v = &H31 Then v = 0 Else v = 255
        PixBits(p) = v: p = p + 1
        PixBits(p) = v: p = p + 1
        PixBits(p) = v: p = p + 1
        j = j - 1
      End If
    Loop
  Next i

End Sub

Private Sub ReadPBMBodyBinary(ByVal Width As Long, ByVal Height As Long, ByVal RowModulo As Long, _
                              ByRef PixBits() As Byte)
                              
  'all bit values are packed ala BMP 2 Color  (1=BLACK)

 Dim i As Long, j As Long, p As Long, q As Long, v As Byte, w As Long, shift As Byte

  ReDim PXRow(0 To (Width + 7) \ 8 - 1)

  For i = 0 To Height - 1
    Get #FileNr, , PXRow()
    p = i * RowModulo
    q = -1
    j = Width
    shift = 128
    Do While j > 0
      If shift = 128 Then q = q + 1
      v = Not PXRow(q)                    'INVERT
      Do While shift > 0
        w = ((v And shift) \ shift) * 255
        PixBits(p) = w: p = p + 1
        PixBits(p) = w: p = p + 1
        PixBits(p) = w: p = p + 1
        shift = shift \ 2
        j = j - 1
      Loop
      shift = 128
    Loop
  Next i

End Sub

Private Sub ReadPGMBodyPlain(ByVal Width As Long, ByVal Height As Long, ByVal RowModulo As Long, _
                             ByRef PixBits() As Byte, ByVal MaxVal As Long)
                             
  'all values are n byte text in size

 Dim i As Long, j As Long, p As Long, v As Long

  For i = 0 To Height - 1
    p = i * RowModulo
    j = Width
    Do While j > 0
      v = GetIntVal()
      If MaxVal > 255 Then v = (255 * v) \ MaxVal
      If v > 255 Then v = 255
      PixBits(p) = v: p = p + 1
      PixBits(p) = v: p = p + 1
      PixBits(p) = v: p = p + 1
      j = j - 1
    Loop
  Next i

End Sub

Private Sub ReadPGMBodyBinary(ByVal Width As Long, ByVal Height As Long, ByVal RowModulo As Long, _
                              ByRef PixBits() As Byte, ByVal MaxVal As Long)
                              
  'all values are 1 or 2 bytes in size

 Dim i As Long, j As Long, p As Long, q As Long, v As Long, hi As Byte, lo As Byte

  If MaxVal <= 255 Then
    ReDim PXRow(0 To Width - 1)
   Else
    ReDim PXRow(0 To 2 * Width - 1)
  End If

  For i = 0 To Height - 1
    Get #FileNr, , PXRow()
    p = i * RowModulo
    q = 0
    j = Width
    Do While j > 0
      If MaxVal <= 255 Then
        v = PXRow(q)
        q = q + 1
       Else
        v = (255 * (PXRow(q) * 256 + PXRow(q + 1))) \ MaxVal  'Rescaled to 0..255
        If v > 255 Then v = 255
        q = q + 2
      End If
      PixBits(p) = v: p = p + 1
      PixBits(p) = v: p = p + 1
      PixBits(p) = v: p = p + 1
      j = j - 1
    Loop
  Next i

End Sub

Private Sub ReadPPMBodyPlain(ByVal Width As Long, ByVal Height As Long, ByVal RowModulo As Long, _
                             ByRef PixBits() As Byte, ByVal MaxVal As Long)
                             
  'all values are n byte text in size

 Dim i As Long, j As Long, p As Long, r As Long, g As Long, b As Long

  For i = 0 To Height - 1
    p = i * RowModulo
    j = Width
    Do While j > 0
      r = GetIntVal()
      g = GetIntVal()
      b = GetIntVal()
      If MaxVal > 255 Then
        r = (255 * r) \ MaxVal
        g = (255 * g) \ MaxVal
        b = (255 * b) \ MaxVal
      End If
      If r > 255 Then r = 255
      If g > 255 Then g = 255
      If b > 255 Then b = 255
      PXRow(p) = b: p = p + 1
      PXRow(p) = g: p = p + 1
      PXRow(p) = r: p = p + 1
      j = j - 1
    Loop
  Next i

End Sub

Private Sub ReadPPMBodyBinary(ByVal Width As Long, ByVal Height As Long, ByVal RowModulo As Long, _
                              ByRef PixBits() As Byte, ByVal MaxVal As Long)

 'Each Pixel is up to 16 Bits each of R,G,B

 Dim i As Long, j As Long, p As Long, q As Long, r As Long, g As Long, b As Long

  If MaxVal <= 255 Then                 'its almost the same as a 24Bitmap
    ReDim PXRow(0 To 3 * Width - 1)

    For i = 0 To Height - 1
      Get #FileNr, , PXRow()

      For j = 0 To 3 * Width - 1 Step 3   'swap r,b
        r = PXRow(j)
        b = PXRow(j + 2)
        PXRow(j) = b
        PXRow(j + 2) = r
      Next j
      Call CopyMemoryRR(PixBits(i * RowModulo), PXRow(0), 3 * Width)
    Next i

   Else
    ReDim PXRow(0 To 6 * Width - 1)

    For i = 0 To Height - 1
      Get #FileNr, , PXRow()
      p = i * RowModulo
      q = 0
      j = Width
      Do While j > 0
        r = (255 * (PXRow(q) * 256 + PXRow(q + 1))) \ MaxVal  'Rescaled to 0..255
        If r > 255 Then r = 255
        q = q + 2

        g = (255 * (PXRow(q) * 256 + PXRow(q + 1))) \ MaxVal  'Rescaled to 0..255
        If g > 255 Then g = 255
        q = q + 2

        b = (255 * (PXRow(q) * 256 + PXRow(q + 1))) \ MaxVal  'Rescaled to 0..255
        If b > 255 Then b = 255
        q = q + 2

        PixBits(p) = b: p = p + 1
        PixBits(p) = g: p = p + 1
        PixBits(p) = r: p = p + 1
        j = j - 1
      Loop
    Next i
  End If

End Sub

'===========================================================================================================

Private Function GetIntVal() As Long   'get an integer value expressed as a string

 Dim i As Long, v As Byte, InNumber As Boolean

  Do
    Get #FileNr, , v
    If v >= &H30 And v <= &H39 Then
      InNumber = True
      i = 10 * i + v - 48
     Else
      If InNumber Then Exit Do  'stop at first non digit after digits
    End If
  Loop
  GetIntVal = i

End Function

Private Function PutIntVal(ByVal v As Long)    'put an integer value expressed as a string int PXBytes()

 Dim i As Long, w As Long, p As Long

  ReDim PXBytes(0 To 9)
  i = 1
  Do While i < v
    i = i * 10
  Loop
  i = i \ 10
  If i = 0 Then i = 1
  Do Until i = 0
    PXBytes(p) = v \ i + 48
    v = v Mod i
    i = i \ 10
    p = p + 1
  Loop
  ReDim Preserve PXBytes(0 To p - 1)

End Function

'============================================ SAVE ===============================================================
'We write only P4 - RPBM, P5 - RPGM, P6 - RPPM  ie. Binary Files
'If NCMapColors=0 we are 24Bit, else we are Mapped 8 Bit

Public Function SavePNM(Path As String, _
                        ByVal Width As Long, ByVal Height As Long, _
                        ByRef PixBits() As Byte, ByVal IntBPP As Long, _
                        ByRef CMap() As RGBA, ByVal NCMapColors As Long, ByVal CMAPMode As Long) As Boolean

 Dim RowModulo As Long

  On Error GoTo SavePNMFailed
  FileNr = FreeFile()
  Open Path For Binary Access Write As #FileNr

  If IntBPP = PIC_1BPP Then
    PXType = PNM_RPBM
   Else
    If IntBPP <= 8 And FSet(CMAPMode, PIC_GMAP) Then
      PXType = PNM_RPGM
     Else
      PXType = PNM_RPPM
    End If
  End If

  Put #FileNr, , PXType
  Put #FileNr, , CByte(10)

  Call PutIntVal(Width)
  Put #FileNr, , PXBytes()
  Put #FileNr, , CByte(32)

  Call PutIntVal(Height)
  Put #FileNr, , PXBytes()
  Put #FileNr, , CByte(10)

  If PXType <> PNM_PBM And PXType <> PNM_RPBM Then   'put MaxVal into header, we always use 255
    Call PutIntVal(255)
    Put #FileNr, , PXBytes()
    Put #FileNr, , CByte(10)
  End If

  If NCMapColors = 0 Then
    RowModulo = BMPRowModulo(Width, PIC_24BPP)
   Else
    RowModulo = BMPRowModulo(Width, PIC_8BPP)
  End If

  Select Case PXType
   Case PNM_RPBM: Call WritePBMBodyBinary(Width, Height, RowModulo, PixBits, CMap(), NCMapColors)
   Case PNM_RPGM: Call WritePGMBodyBinary(Width, Height, RowModulo, PixBits, CMap(), NCMapColors)
   Case PNM_RPPM: Call WritePPMBodyBinary(Width, Height, RowModulo, PixBits, CMap(), NCMapColors)
  End Select
  SavePNM = True

SavePNMFailed:
  Close #FileNr
  On Error GoTo 0

End Function

'---------------------------------------------------------------------------------------------------------------------
Private Sub WritePBMBodyBinary(ByVal Width As Long, ByVal Height As Long, ByVal RowModulo As Long, _
                               ByRef PixBits() As Byte, ByRef CMap() As RGBA, ByVal NCMapColors As Long)

 Dim i As Long, j As Long, p As Long, b As Long, maskx As Byte, Limit As Long

  Limit = (Width + 7) \ 8
  ReDim PXRow(0 To Limit - 1)   '8 bits to a byte traiing 0s

  For i = 0 To Height - 1
    maskx = 128: b = 0: p = 0
    For j = i * RowModulo To i * RowModulo + Width - 1
      If PixBits(j) = 0 Then b = b Or maskx                 '0=WHITE
      maskx = maskx \ 2
      If maskx = 0 Then
        PXRow(p) = b
        p = p + 1
        maskx = 128
        b = 0
      End If
    Next j
    If maskx <> 128 Then PXRow(p) = b         'the remaining bits
    Put #FileNr, , PXRow()
  Next i

End Sub

Private Sub WritePGMBodyBinary(ByVal Width As Long, ByVal Height As Long, ByVal RowModulo As Long, _
                               ByRef PixBits() As Byte, ByRef CMap() As RGBA, ByVal NCMapColors As Long)

 Dim i As Long, j As Long, p As Long, g As Long

  ReDim PXRow(0 To Width - 1)           '8 bits for Grey

  For i = 0 To Height - 1
    p = i * RowModulo
    For j = 0 To Width - 1
      If NCMapColors = 0 Then
        g = PixBits(p): p = p + 3
       Else
        g = CMap(PixBits(p)).Green: p = p + 1
      End If
      PXRow(j) = g
    Next j
    Put #FileNr, , PXRow()
  Next i

End Sub

Private Sub WritePPMBodyBinary(ByVal Width As Long, ByVal Height As Long, ByVal RowModulo As Long, _
                               ByRef PixBits() As Byte, ByRef CMap() As RGBA, ByVal NCMapColors As Long)

 Dim i As Long, j As Long, p As Long, r As Long, g As Long, b As Long, Limit As Long

  Limit = 3 * Width
  ReDim PXRow(0 To Limit - 1)       '8 bits to each of r,g,b

  For i = 0 To Height - 1
    p = i * RowModulo
    j = 0
    Do
      If NCMapColors = 0 Then
        b = PixBits(p): p = p + 1
        g = PixBits(p): p = p + 1
        r = PixBits(p): p = p + 1
       Else
        b = CMap(PixBits(p)).Blue
        g = CMap(PixBits(p)).Green
        r = CMap(PixBits(p)).Red
        p = p + 1
      End If
      PXRow(j) = r: j = j + 1
      PXRow(j) = g: j = j + 1
      PXRow(j) = b: j = j + 1
    Loop Until j = Limit

    Put #FileNr, , PXRow()
  Next i

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jun-15 15:26) 27 + 451 = 478 Lines
