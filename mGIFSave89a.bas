Attribute VB_Name = "mGIFSave89a"
Option Explicit

' mGIFSave.bas  -  master file for writing GIF files

'- ©2001/2003 Ron van Tilburg - All rights reserved  1.01.2001/Jun-Jul 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

' from the C copyright ©1997 Ron van Tilburg 25.12.1997
' VB copyright ©2000 Ron van Tilburg 24.12.2000     'what xmas holidays are good for <:-)
' and copyrights of the original C code from which this is derived are given in the body
' Documentation of GIF structures is from the GIF standard as attached as html documents
' All copyrights applying there continue to apply

' Unisys Corp believes it has the Copyright on all LZW algorithms for GIF files. If it worries you then
' dont use this code. Read the HTML standards for the owner of the copyright of GIFs and its usability

' Start at the bottom of this file at the SaveGIF function and work upwards

' General Disclaimer: I think this all works ok (but it needs some exercise to prove it) but you use and
' adapt it at your own risk. However reference to my authorship would be appreciated if you use it in
' the public domain.  Ron van Tilburg Xmas 2000.

'Transparent GIF support for colour other than index 0
' Pietro Cecchi & RVT 12 Oct 2001

' GIF structures    (not actually Used)

Private Type GifScreenHdr
  Width As Integer
  Height As Integer
  MCR0Pix As Byte     'see below
  BC As Byte
  Aspect As Byte
End Type

Private Type GifImageHdr
  Left As Integer
  Top As Integer
  Width  As Integer
  Height As Integer
  MIPixBits As Byte
  CodeSize As Byte
End Type

'GLOBAL VRIABLES for the Encoding Routines ============================================================

'**************************************************************************'
'  FROM GIFCOMPR.C       - GIF Image compression routines
'
'  Lempel-Ziv compression based on 'compress'.  GIF modifications by
'  David Rowley (mgardi@watdcsu.waterloo.edu)
'
'*************************************************************************
' an Integer must be able to hold 2**BITS values of type int, and also -1

Private Const MAXBITS      As Integer = 12              ' user settable max - bits/code
Private Const MAXBITSHIFT  As Integer = 2 ^ MAXBITS
Private Const MAXMAXCODE   As Integer = 2 ^ MAXBITS     ' should NEVER generate this code
Private Const HASHTABSIZE  As Integer = 5003            ' 80% occupancy

' GIF Image compression - modified 'compress'
'
' Based on: compress.c - File compression ala IEEE Computer, June 1984.
'
' By Authors:  Spencer W. Thomas       (decvax!harpo!utah-cs!utah-gr!thomas)
'              Jim McKie               (decvax!mcvax!jim)
'              Steve Davies            (decvax!vax135!petsd!peora!srd)
'              Ken Turkowski           (decvax!decwrl!turtlevax!ken)
'              James A. Woods          (decvax!ihnp4!ames!jaw)
'              Joe Orost               (decvax!vax135!petsd!joe)
' VB code by   Ron van Tilburg          rivit@f1.net.au

Private nBits As Integer                    ' number of bits/code
Private MaxCode As Integer                   ' maximum code, given nBits

'-define MAXCODE(nBits) (((Integer)1 << (nBits)) - 1)    '=masks(nBits)

Private HashTab(0 To HASHTABSIZE - 1) As Long
Private CodeTab(0 To HASHTABSIZE - 1) As Integer

' To save much memory, we overlay the table used by compress() with those used by decompress().
' The tab_prefix table is the same size and type as the codetab.  The tab_suffix table needs
' 2**MAXBITS characters.  We get this from the beginning of HashTab.  The output stack uses the rest
' of HashTab, and contains characters.  There is plenty of room for any possible stack
' (stack used to be 8000 characters).

'-define tab_prefixof(i) CodeTabOf(i)
'-define tab_suffixof(i) ((byte*)(HashTab))[i]
'-define de_stack        ((byte*)&tab_suffixof((Integer)1<<MAXBITS))

Private FirstFree     As Integer        ' first unused entry

' block compression parameters -- after all codes are used up, and compression rate changes, start over.
Private isCleared     As Boolean
Private Offset        As Integer
Private In_Count      As Long           ' length of input
Private Out_Count     As Long           ' - of codes output (for debugging)

' Algorithm:  use open addressing double hashing (no chaining) on the prefix code / next character
' combination.  We do a variant of Knuth's algorithm D (vol. 3, sec. 6.4) along with G. Knott's
' relatively-prime secondary probe.  Here, the modular division first probe is gives way to a faster
' exclusive-or manipulation.  Also do block compression with an adaptive reset, whereby the code table
' is cleared when the compression ratio decreases, but after the table fills.  The variable-length output
' codes are re-sized at this point, and a special CLEAR code is generated for the decompressor.  Late
' addition:  construct the table according to file size for noticeable speed improvement on small files.

Private g_Init_Bits   As Integer
Private ClearCode     As Integer
Private EOFCode       As Integer

'variables for positioning and control
Private CurX          As Integer      'current xpos
Private CurY          As Integer      'current ypos
Private GIFWidth      As Long         'the Width
Private GIFHeight     As Long         'the Height
Private GIFRowMod     As Long         'the rowModulo - BMPs can have extension bits for padding <=GIFWidth
Private GIFPixSize    As Long         'the nr of bits for a given pixel

Private Countdown     As Long         'pixels left to do
Private Pass          As Integer      'which pass in interlaced mode
Private GIFInterlace  As Boolean      'use interlace mode
Private FileCount     As Long         'bytes output so far

Private Const EOF     As Integer = -1 'END of input

'variables for the code accumulator (OutputCode)
Private CurAccum      As Long
Private CurBits       As Integer
Private Masks(0 To 16) As Long      'powers of 2 -1

'variables for the outputbyte accumulator

Private AccumCnt      As Integer      'Number of characters so far in this 'packet'
Private Accum()       As Byte         'will be max 256 bytes long, first byte is length

Private FileNr            As Long         'FileNumber to write to
'======================================================================================================
'======================================================================================================
Private Sub CompressAndWriteBits(Init_Bits As Integer, ByRef PixBits() As Byte)

 Dim fcode As Long
 Dim i As Long, c As Long, ent As Long, disp As Long
 Dim hshift As Long, zm() As Variant

 'set up where we are starting

  i = 0
  FileCount = 0
  Pass = 0
  CurX = 0
  CurY = 0
  Countdown = GIFWidth * GIFHeight

  'set up the code accumulator
  CurAccum = 0
  CurBits = 0
  zm = Array(&H0&, &H1&, &H3&, &H7&, &HF&, _
             &H1F&, &H3F&, &H7F&, &HFF&, _
             &H1FF&, &H3FF&, &H7FF&, &HFFF&, _
             &H1FFF&, &H3FFF&, &H7FFF&, &HFFFF&)   'Array values of 2^N-1  N=0,1,2,,..16
  For i = 0 To 16
    Masks(i) = CLng(zm(i))
  Next i

  '  Set up the globals:  g_init_bits - initial number of bits
  g_Init_Bits = Init_Bits

  '  Set up the necessary values
  Offset = 0
  Out_Count = 0
  isCleared = False

  nBits = g_Init_Bits
  MaxCode = Masks(nBits)                'MAXCODE(nBits);

  ClearCode = 2 ^ (Init_Bits - 1)
  EOFCode = ClearCode + 1
  FirstFree = ClearCode + 2

  Call Char_Init                        'set up output buffers

  hshift = 0
  fcode = HASHTABSIZE
  Do While fcode < 65536
    hshift = hshift + 1
    fcode = fcode + fcode
  Loop
  hshift = 1 + Masks(8 - hshift)        'set hash code range bound for shifting

  Call ClearHash                        'clear hash table
  Call OutputCode(ClearCode)            'get ready to go
  Out_Count = 1

  ent = GIFNextPixel(PixBits)
  In_Count = 1

  c = GIFNextPixel(PixBits)
  Do While c <> EOF
    In_Count = In_Count + 1

    fcode = c * MAXBITSHIFT + ent
    i = (c * hshift) Xor ent            'xor hashing

    If HashTab(i) = fcode Then
      ent = CodeTab(i)
      GoTo NextPixel
     ElseIf HashTab(i) < 0 Then         'empty slot
      GoTo NoMatch
    End If

    disp = HASHTABSIZE - i              ' secondary hash (after G. Knott)
    If i = 0 Then disp = 1

Probe:
    i = i - disp
    If i < 0 Then i = i + HASHTABSIZE

    If HashTab(i) = fcode Then
      ent = CodeTab(i)
      GoTo NextPixel
    End If

    If HashTab(i) > 0 Then GoTo Probe

NoMatch:
    Call OutputCode(ent)
    Out_Count = Out_Count + 1
    ent = c

    If FirstFree < MAXMAXCODE Then
      CodeTab(i) = FirstFree
      FirstFree = FirstFree + 1         'code -> hashtable
      HashTab(i) = fcode
     Else 'NOT FirstFree...
      Call ClearBlock
    End If
NextPixel:
    c = GIFNextPixel(PixBits)
  Loop

  '  Put out the final code.

  Call OutputCode(ent)
  Out_Count = Out_Count + 1
  Call OutputCode(EOFCode)
  Out_Count = Out_Count + 1

End Sub

'Return the next pixel from the image and increment positions
Private Function GIFNextPixel(ByRef PixBits() As Byte) As Integer

 Dim RowOffset As Long, Mask As Long

  If (Countdown = 0) Then
    GIFNextPixel = EOF
   Else
    Countdown = Countdown - 1
    RowOffset = LBound(PixBits) + GIFRowMod * CurY

    Select Case GIFPixSize         '1,4,8 from a bitmap
     Case 8:                                  'every byte is a pixel
      GIFNextPixel = PixBits(RowOffset + CurX)

     Case 4:                                  'every nybble is a pixel
      If (CurX And 1) = 1 Then
        GIFNextPixel = CLng(PixBits(RowOffset + CurX \ 2)) And &HF&         'odd
       Else
        GIFNextPixel = (CLng(PixBits(RowOffset + CurX \ 2)) And &HF0&) \ 16 'even
      End If

     Case 1:                                  'every bit is a pixel
      Mask = 2& ^ (7 - CurX Mod 8)
      GIFNextPixel = (CLng(PixBits(RowOffset + CurX \ 8)) And Mask) \ Mask
    End Select

    'Bump the current X position
    CurX = CurX + 1

    'If we are at the end of a scan line, set curx back to the beginning
    'If we are interlaced, bump the CurY to the appropriate spot, otherwise, just increment it.

    If CurX = GIFWidth Then
      CurX = 0
      If GIFInterlace = False Then
        CurY = CurY + 1
       Else 'NOT INTERLACE...
        Select Case Pass
         Case 0:
          CurY = CurY + 8
          If CurY >= GIFHeight Then
            Pass = Pass + 1
            CurY = 4
          End If

         Case 1:
          CurY = CurY + 8
          If CurY >= GIFHeight Then
            Pass = Pass + 1
            CurY = 2
          End If

         Case 2:
          CurY = CurY + 4
          If CurY >= GIFHeight Then
            Pass = Pass + 1
            CurY = 1
          End If

         Case 3:
          CurY = CurY + 2
        End Select
      End If
    End If
  End If

End Function

' TAG( OutputCode )
' Output the given code.
'  Inputs:
'    code: A nBits-bit integer.  If == -1, then EOF.  This assumes that nBits =< (long)wordsize - 1.
'  Outputs:
'    Outputs code to the file.
'  Assumptions:
'    Chars are 8 bits long.
'  Algorithm:
'    Maintain a MAXBITS character long buffer (so that 8 codes will fit in it exactly).
'    When the buffer fills up empty it and start over.

Private Sub OutputCode(ByVal Code As Long)

  CurAccum = CurAccum And Masks(CurBits)

  If (CurBits > 0) Then
    CurAccum = CurAccum Or (Code * (1 + Masks(CurBits)))
   Else 'NOT (CurBits...
    CurAccum = Code
  End If

  CurBits = CurBits + nBits

  Do While (CurBits >= 8)
    Call Char_Out(CurAccum And &HFF&)
    CurAccum = CurAccum \ 256&
    CurBits = CurBits - 8
  Loop

  ' If the next entry is going to be too big for the code size, then increase it, if possible.

  If FirstFree > MaxCode Or isCleared Then
    If isCleared Then
      nBits = g_Init_Bits
      MaxCode = Masks(nBits)       'MAXCODE(nBits);
      isCleared = False
     Else 'NOT (isCleared...
      nBits = nBits + 1
      If (nBits = MAXBITS) Then
        MaxCode = MAXMAXCODE
       Else 'NOT (nBits...
        MaxCode = Masks(nBits)     'MAXCODE(nBits);
      End If
    End If
  End If

  If (Code = EOFCode) Then         'At EOF, write the rest of the buffer.
    Do While (CurBits > 0)
      Call Char_Out(CurAccum And &HFF&)
      CurAccum = CurAccum \ 256&
      CurBits = CurBits - 8
    Loop
    Call Flush_Char
  End If

End Sub

' Clear out the hash table
Private Sub ClearBlock()                          ' table clear for block compress

  Call ClearHash
  FirstFree = ClearCode + 2
  isCleared = True
  Call OutputCode(ClearCode)
  Out_Count = Out_Count + 1

End Sub

Private Sub ClearHash()                           ' reset code table

 Dim i As Long

  For i = 0 To HASHTABSIZE - 1
    HashTab(i) = -1
  Next i

End Sub

' Set up the 'byte output' routine and Define the storage for the packet accumulator
Private Sub Char_Init()

  AccumCnt = 0
  ReDim Accum(0 To 255) As Byte

End Sub

' Add a character to the end of the current packet, and if it is 254 characters, flush the packet to disk.
Private Sub Char_Out(ByVal c As Integer)

  Accum(AccumCnt + 1) = c              '0,1,2,3 ....mapped to 1,2,3,4...255
  AccumCnt = AccumCnt + 1
  If AccumCnt >= 254 Then Call Flush_Char      'in the original this was >=254, the std allows 255
  '(most art programs Ive got seem to use this 254 code)

End Sub

' Flush the current packet to disk, and reset the accumulator
Private Sub Flush_Char()

  If AccumCnt > 0 Then
    Accum(0) = AccumCnt                                'set block length
    ReDim Preserve Accum(0 To AccumCnt) As Byte        'and redimension to this length
    Put #FileNr, , Accum                                  'write it to disk
    FileCount = FileCount + AccumCnt + 1               'track bytes written
    Call Char_Init
  End If

End Sub

'============================ THE REAL ROUTINES PUBLICLY VISIBLE =========================================

Public Function SaveGIF(ByVal Path As String, _
                        ByVal Width As Long, ByVal Height As Long, ByVal BitsPerPixel As Long, _
                        ByRef PixBits() As Byte, ByVal PixelWidth As Long, ByRef CMap() As RGBA, _
                        Optional ByVal Interlaced As Boolean = False, _
                        Optional ByVal TransparentRGBColor As Long = -1) As Long   '<=0=failure,1=success

 'Path:         Where you will store the file should end .gif
 'Width,Height: Pic size in pixels
 'BitsPerPixel: in planes 1=BW, 4=16 colours, 8=256 colours
 'PixBits:      the bits of the picture from top-left in BitsPerPixel=1 1 pixel=1 bit, BitsPerPixel=4 1 pixel=4bits, 8 1pixel=8bits
 'PixelWidth:   how wide is a pixel packed in bits should be 1,4,8 for MS bitmaps

 'When calling this routine independently make sure the image is the right way up, and that a colour map exists
 'CMap: the three byte tuples r,g,b for each colour in the image. (should be 4*2^n bytes n=1,,8) NOT CHECKED
 'Interlaced: make the GIF interlaced (see doco)

 Dim ID As String
 Dim GSH As GifScreenHdr
 Dim GIH As GifImageHdr
 Dim RGB As RGBA, i As Long, j As Long  'for Transparent color support

 ' attempt to save a gif file of the bitmap data with colormap cmap
 ' CMAP contains 2^BitsPerPixel colours of 3 bytes each r,g,b
 ' Bits contains the colour mapped data as 1 byte per pixel as mapped by colormap

  On Error GoTo BadPath
  FileNr = FreeFile()                                 'A GLOBAL
  Open Path For Binary Access Write As #FileNr
  On Error GoTo GIFSaveFailed

  'File identifier
  ID = "GIF89a"                 'A later Filetype with Additions like Comment and Transparency
  Put #FileNr, , ID

  'ScreenDescriptor
  '              Bits
  '         7 6 5 4 3 2 1 0  Byte -
  '        +---------------+
  '        |               |  1
  '        +-Screen Width -+      Raster width in pixels (LSB first)
  '        |               |  2
  '        +---------------+
  '        |               |  3
  '        +-Screen Height-+      Raster height in pixels (LSB first)
  '        |               |  4
  '        +-+-----+-+-----+      M = 1, Global color map follows Descriptor
  '        |M|  cr |0|pixel|  5   cr+1 = - bits of color resolution
  '        +-+-----+-+-----+      pixel+1 = - bits/pixel in image
  '        |   background  |  6   background=Color index of screen background
  '        +---------------+          (color is defined from the Global color
  '        |0 0 0 0 0 0 0 0|  7        map or default map if none specified)
  '        +---------------+

  '        The logical screen width and height can both  be  larger  than  the
  '   physical  display.   How  images  larger  than  the physical display are
  '   handled is implementation dependent and can take advantage  of  hardware
  '   characteristics  (e.g.   Macintosh scrolling windows).  Otherwise images
  '   can be clipped to the edges of the display.

  '        The value of 'pixel' also defines  the  maximum  number  of  colors
  '   within  an  image.   The  range  of  values  for 'pixel' is 0 to 7 which
  '   represents 1 to 8 bits.  This translates to a range of 2 (B & W) to  256
  '   colors.   Bit  3 of word 5 is reserved for future definition and must be
  '   zero.

  With GSH
    .Width = Width
    .Height = Height
    .MCR0Pix = &HF0 Or (BitsPerPixel - 1)
    .BC = 0
    .Aspect = 0
  End With

  'done this way to make sure LoHi storage in file
  Put #FileNr, , GSH.Width
  Put #FileNr, , GSH.Height
  Put #FileNr, , GSH.MCR0Pix
  Put #FileNr, , GSH.BC
  Put #FileNr, , GSH.Aspect

  'Global ColorMap
  '        The Global Color Map is optional but recommended for  images  where
  '   accurate color rendition is desired.  The existence of this color map is
  '   indicated in the 'M' field of byte 5 of the Screen Descriptor.  A  color
  '   map  can  also  be associated with each image in a GIF file as described
  '   later.  However this  global  map  will  normally  be  used  because  of
  '   hardware  restrictions  in equipment available today.  In the individual
  '   Image Descriptors the 'M' flag will normally be  zero.   If  the  Global
  '   color Map Is present, it    's definition immediately follows the Screen
  '   Descriptor.   The  number  of  color  map  entries  following  a  Screen
  '   Descriptor  is equal to 2**(- bits per pixel), where each entry consists
  '   of three byte values representing the relative intensities of red, green
  '   and blue respectively.  The structure of the Color Map block is:

  '              Bits
  '         7 6 5 4 3 2 1 0  Byte -
  '        +---------------+
  '        | red intensity |  1    Red value for color index 0
  '        +---------------+
  '        |green intensity|  2    Green value for color index 0
  '        +---------------+
  '        | blue intensity|  3    Blue value for color index 0
  '        +---------------+
  '        | red intensity |  4    Red value for color index 1
  '        +---------------+
  '        |green intensity|  5    Green value for color index 1
  '        +---------------+
  '        | blue intensity|  6    Blue value for color index 1
  '        +---------------+
  '        :               :       (Continues for remaining colors)

  '        Each image pixel value received will be displayed according to  its
  '   closest match with an available color of the display based on this color
  '   map.  The color components represent a fractional intensity  value  from
  '   none  (0)  to  full (255).  White would be represented as (255,255,255),
  '   black as (0,0,0) and medium yellow as (180,180,0).  For display, if  the
  '   device  supports fewer than 8 bits per color component, the higher order
  '   bits of each component are used.  In the creation of  a  GIF  color  map
  '   entry  with  hardware  supporting  fewer  than 8 bits per component, the
  '   component values for the hardware  should  be  converted  to  the  8-bit
  '   format with the following calculation:

  '        map_value> = component_value>*255/(2**nbits> -1)

  '        This assures accurate translation of colors for all  displays.   In
  '   the  cases  of  creating  GIF images from hardware without color palette
  '   capability, a fixed palette should be created  based  on  the  available
  '   display  colors for that hardware.  If no Global Color Map is indicated,
  '   a default color map is generated internally  which  maps  each  possible
  '   incoming  color  index to the same hardware color index modulo where
  '   is the number of available hardware colors.
  For i = 0 To UBound(CMap)
    Put #FileNr, , CMap(i).Red
    Put #FileNr, , CMap(i).Green
    Put #FileNr, , CMap(i).Blue
  Next i
  '  Put #FileNr, , CMap     'NOTE THIS IS NOT CHECKED should be of size 3*2^k, k=1..8
  '   Graphic Control Extension [TOC]
  '
  '   Description
  '   The Graphic Control Extension contains parameters used when processing a graphic rendering block. The scope of this extension is the first graphic rendering block to follow. The extension contains only one data sub-block.
  '   This block is OPTIONAL; at most one Graphic Control Extension may precede a graphic rendering block. This is the only limit to the number of Graphic Control Extensions that may be contained in a Data Stream.
  '
  '
  '   Required Version
  '   89a.
  '
  '   Syntax
  '       7 6 5 4 3 2 1 0        Field Name                    Type
  '      +---------------+
  '   0  |               |       Extension Introducer          Byte    x21
  '      +---------------+
  '   1  |               |       Graphic Control Label         Byte  xF9
  '      +---------------+
  '
  '      +---------------+
  '   0  |               |       Block Size                    Byte  x04
  '      +---------------+
  '   1  |     |     | | |       <Packed Fields'   See below    000|000|0|1
  '      +---------------+
  '   2  |               |       Delay Time                    Unsigned   0x00,0x00
  '      +-             -+
  '   3  |               |
  '      +---------------+
  '   4  |               |       Transparent Color Index       Byte     0x00 to 0xFF
  '      +---------------+
  '
  '      +---------------+
  '   0  |               |       Block Terminator              Byte x00
  '      +---------------+
  '
  '
  '   <Packed Fields>
  '   Reserved                      3 Bits
  '   Disposal Method               3 Bits
  '   User Input Flag               1 Bit
  '   Transparent Color Flag        1 Bit
  '
  '   Extension Introducer
  '   Identifies the beginning of an extension
  '
  '   Graphic Control Label
  '   Identifies the current block as a Graphic Control Extension. This field contains the fixed value 0xF9.
  '
  '   Block Size
  '   Number of bytes in the block, after the Block Size field and up to but not including the Block Terminator. This field contains the fixed value 4.
  '
  '   Disposal Method
  '   Indicates the way in which the graphic is to be treated after being displayed.
  '
  '   Values:
  '   0 - No disposal specified. The decoder is not required to take any action.
  '   1 - Do not dispose. The graphic is to be left in place.
  '   2 - Restore to background color. The area used by the graphic must be restored to the background color.
  '   3 - Restore to previous. The decoder is required to restore the area overwritten by the graphic with what was there prior to rendering the graphic.
  '   4-7 - To be defined.
  '
  '   User Input Flag
  '   Indicates whether or not user input is expected before continuing. If the flag is set, processing will continue when user input is entered. The nature of the User input is determined by the application (Carriage Return, Mouse Button Click, etc.).
  '
  '   Values:
  '   0 - User input is not expected.
  '   1 - User input is expected.
  '   When a Delay Time is used and the User Input Flag is set, processing will continue when user input is received or when the delay time expires, whichever occurs first.
  '
  '
  '   Transparency Flag
  '   Indicates whether a transparency index is given in the Transparent Index field. (This field is the least significant bit of the byte.)
  '
  '   Values:
  '   0 - Transparent Index is not given.
  '   1 - Transparent Index is given.
  '
  '   Delay Time
  '   If not 0, this field specifies the number of hundredths (1/100) of a second to wait before continuing with the processing of the Data Stream. The clock starts ticking immediately after the graphic is rendered. This field may be used in conjunction with the User Input Flag field.
  '
  '   Transparency Index
  '   The Transparency Index is such that when encountered, the corresponding pixel of the display device is not modified and processing goes on to the next pixel. The index is present if and only if the Transparency Flag is set to 1.
  '
  '   Block Terminator
  '   This zero-length data block marks the end of the Graphic Control Extension.
  '
  '   Extensions and Scope
  '   The scope of this Extension is the graphic rendering block that follows it; it is possible for other extensions to be present between this block and its target. This block can modify the Image Descriptor Block and the Plain Text Extension.
  '
  '   Recommendations
  '   Disposal Method
  '   The mode Restore To Previous is intended to be used in small sections of the graphic; the use of this mode imposes severe demands on the decoder to store the section of the graphic that needs to be saved. For this reason, this mode should be used sparingly. This mode is not intended to save an entire graphic or large areas of a graphic; when this is the case, the encoder should make every attempt to make the sections of the graphic to be restored be separate graphics in the data stream. In the case where a decoder is not capable of saving an area of a graphic marked as Restore To Previous, it is recommended that a decoder restore to the background color.
  '
  '   User Input Flag
  '   When the flag is set, indicating that user input is expected, the decoder may sound the bell (0x07) to alert the user that input is being expected. In the absence of a specified Delay Time, the decoder should wait for user input indefinitely. It is recommended that the encoder not set the User Input Flag without a Delay Time specified.

  'Pietro helped me with this
  If TransparentRGBColor <> -1 Then       'see if given color is in CMap and call it transparent
    RGB = GetRGBA(TransparentRGBColor)
    TransparentRGBColor = -1                           'assume we cant find it
    For i = 0 To UBound(CMap())
      If CMap(i).Red = RGB.Red Then
        If CMap(i).Green = RGB.Green Then
          If CMap(i).Blue = RGB.Blue Then Exit For
        End If
      End If
    Next i
    If i <= UBound(CMap) Then      'we found it
      TransparentRGBColor = i
    End If
  End If

  If TransparentRGBColor = -1 Then       'just leave with no transparency
    Put #FileNr, , Chr$(33) & Chr$(249) & Chr$(4) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
   Else
    Put #FileNr, , Chr$(33) & Chr$(249) & Chr$(4) & Chr$(1) & Chr$(0) & Chr$(0) & Chr$(TransparentRGBColor) & Chr$(0)
  End If

  'ImageDescriptor
  '        The Image Descriptor defines the actual placement  and  extents  of
  '   the  following  image within the space defined in the Screen Descriptor.
  '   Also defined are flags to indicate the presence of a local color  lookup
  '   map, and to define the pixel display sequence.  Each Image Descriptor is
  '   introduced by an image separator  character.   The  role  of  the  Image
  '   Separator  is simply to provide a synchronization character to introduce
  '   an Image Descriptor.  This is desirable if a GIF file happens to contain
  '   more  than  one  image.   This  character  is defined as 0x2C hex or ','
  '   (comma).  When this character is encountered between images,  the  Image
  '   Descriptor will follow immediately.

  '        Any characters encountered between the end of a previous image  and
  '   the image separator character are to be ignored.  This allows future GIF
  '   enhancements to be present in newer image formats and yet ignored safely
  '   by older software decoders.

  '              Bits
  '         7 6 5 4 3 2 1 0  Byte -
  '        +---------------+
  '        |0 0 1 0 1 1 0 0|  1    ',' - Image separator character &H2C
  '        +---------------+
  '        |               |  2    Start of image in pixels from the
  '        +-  Image Left -+       left side of the screen (LSB first)
  '        |               |  3
  '        +---------------+
  '        |               |  4
  '        +-  Image Top  -+       Start of image in pixels from the
  '        |               |  5    top of the screen (LSB first)
  '        +---------------+
  '        |               |  6
  '        +- Image Width -+       Width of the image in pixels (LSB first)
  '        |               |  7
  '        +---------------+
  '        |               |  8
  '        +- Image Height-+       Height of the image in pixels (LSB first)
  '        |               |  9
  '        +-+-+-+-+-+-----+       M=0 - Use global color map, ignore 'pixel'
  '        |M|I|0|0|0|pixel| 10    M=1 - Local color map follows, use 'pixel'
  '        +-+-+-+-+-+-----+       I=0 - Image formatted in Sequential order
  '                                I=1 - Image formatted in Interlaced order
  '                                pixel+1 - - bits per pixel for this image
  '
  '        The specifications for the image position and size must be confined
  '   to  the  dimensions defined by the Screen Descriptor.  On the other hand
  '   it is not necessary that the image fill the entire screen defined.

  With GIH
    .Left = 0
    .Top = 0
    .Width = Width
    .Height = Height
    .MIPixBits = BitsPerPixel - 1
    If Interlaced Then .MIPixBits = .MIPixBits Or &H40
    'code size is part of the raster stream but for convenience Ive added it to the ImageHeader
    If BitsPerPixel = 1 Then .CodeSize = 2 Else .CodeSize = BitsPerPixel    'see below
  End With

  'done this way to make sure LoHi storage in file
  Put #FileNr, , Chr$(&H2C)
  Put #FileNr, , GIH.Left
  Put #FileNr, , GIH.Top
  Put #FileNr, , GIH.Width
  Put #FileNr, , GIH.Height
  Put #FileNr, , GIH.MIPixBits
  Put #FileNr, , GIH.CodeSize

  'The Compressed Bits
  '        The Raster Data stream that represents the actual output image  can
  '   be represented as:

  '         7 6 5 4 3 2 1 0
  '        +---------------+
  '        |   code size   |
  '        +---------------+     ---+
  '        |blok byte count|        |
  '        +---------------+        |
  '        :               :        +-- Repeated as many times as necessary
  '        |  data bytes   |        |
  '        :               :        |
  '        +---------------+     ---+
  '        . . .       . . .
  '        +---------------+
  '        |0 0 0 0 0 0 0 0|       zero byte count (terminates data stream)
  '        +---------------+

  '        The conversion of the image from a series  of  pixel  values  to  a
  '   transmitted or stored character stream involves several steps.  In brief
  '   these steps are:

  '    Establish the Code Size -
  '       Define  the  number  of  bits  needed  to
  '       represent the actual data.

  '   Compress the Data -
  '       Compress the series of image pixels to a  series
  '       of compression codes.

  '   Build a Series of Bytes -
  '       Take the  set  of  compression  codes  and
  '       convert to a string of 8-bit bytes.

  '   Package the Bytes -
  '       Package sets of bytes into blocks  preceeded  by
  '       character counts and output.

  '   Establish Code Size
  '        The first byte of the GIF Raster Data stream is a value  indicating
  '   the minimum number of bits required to represent the set of actual pixel
  '   values.  Normally this will be the same as the  number  of  color  bits.
  '   Because  of  some  algorithmic constraints however, black & white images
  '   which have one color bit must be indicated as having a code size  of  2.
  '   This  code size value also implies that the compression codes must start
  '   out one bit longer.

  'set some external globals
  GIFWidth = Width
  GIFHeight = Height
  GIFPixSize = PixelWidth
  GIFRowMod = (UBound(PixBits) - LBound(PixBits) + 1) \ GIFHeight   'BMPS can have hanging bits at the end of rows
  GIFInterlace = Interlaced
  Call CompressAndWriteBits(GIH.CodeSize + 1, PixBits)

  'write the trailer, terminator
  Put #FileNr, , Chr$(&H3B)

  Close #FileNr
  SaveGIF = 1

Exit Function

BadPath:
  SaveGIF = 0
  On Error GoTo 0

Exit Function

GIFSaveFailed:
  Close #FileNr
  Call Kill(Path)     'no idea if file is any good so kill it (could fail here if open failed)
  SaveGIF = 0
  On Error GoTo 0

End Function

':) Ulli's VB Code Formatter V2.16.6 (2003-Jun-13 22:28) 134 + 698 = 832 Lines
