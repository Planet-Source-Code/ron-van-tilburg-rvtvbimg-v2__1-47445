VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRVTVBIMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'cRVTVBIMG.cls    - Image saving Pipeline Processing  - saves BMP,GIF and PNM

'- ©2001/2003 Ron van Tilburg - All rights reserved  1.01.2001/Jun-Jul 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'Jun/July 2003  Major Additions made
'=============================================================================================================
'================================== TYPES AND ENUMS ==========================================================
'=============================================================================================================

Public Enum IMG_PICTYPE 'PIC Save Type
  PIC_BMP = 0               'save as BMP
  PIC_GIF = 1               'save as GIF
  PIC_GIF_LACED = 2         'save as interlaced gif
  PIC_PNM = 3               'save as the most appropriate of PBM, PGM or PPM
  'others should be added here
End Enum

Public Enum IMG_COLORMODES
  PIC_COLOR = 0             'Color style  default
  PIC_BW = 1                'Just Black and White
  PIC_GREY = 2              'colours are converted eventually into greys (not necessarily unique)
End Enum

Public Enum IMG_DEPTHMODES       'Depth in Bits and the synonymous Number of colors
  PIC_1BPP = 1                    'GIF & BMP implies Black and White
  PIC_2BPP = 2                    'GIF 4 colors
  PIC_3BPP = 3                    'GIF 8 colors
  PIC_4BPP = 4                    'GIF 16 or BMP implies VGA 16 color palette
  PIC_5BPP = 5                    'GIF 32 colors
  PIC_6BPP = 6                    'GIF 64 colors
  PIC_7BPP = 7                    'GIF 128 colors
  PIC_8BPP = 8                    'GIF 256 or BMP
  PIC_9BPP = 9                    '  8 * 8 * 8 =     512
  PIC_12BPP = 12                  ' 16* 16* 16 =    4K
  PIC_15BPP = 15                  ' 32* 32* 32 =   32K
  PIC_16BPP = 16                  'BMP 32K 5b*5b*5b colour implicit in pixeldata  (no Grey Mode)
  PIC_24BPP = 24                  '256*256*256 =   16M BMP default  colour implicit in pixeldata  (no Grey Mode)
  PIC_32BPP = 32                  'BMP 16M 8b*8b*8b colour implicit in pixeldata  (no Grey Mode)

  'NColors in Bits
  PIC_C2 = 1                      'GIF & BMP implies Black and White
  PIC_C4 = 2                      'GIF 4 colors
  PIC_C8 = 3                      'GIF 8 colors
  PIC_C16 = 4                     'GIF 16 or BMP implies VGA 16 color palette
  PIC_C32 = 5                     'GIF 32 colors
  PIC_C64 = 6                     'GIF 64 colors
  PIC_C128 = 7                    'GIF 128 colors
  PIC_C256 = 8                    'GIF 256 or BMP implies Internet palette default
  
  PIC_C512 = 9                    'BMP          colour implicit in pixeldata  (no Grey Mode)
  PIC_C4K = 12                    'BMP          colour implicit in pixeldata  (no Grey Mode)
  PIC_C32K = 15                   'BMP          colour implicit in pixeldata  (no Grey Mode)
  PIC_C64K = 16                   'BMP          colour implicit in pixeldata  (no Grey Mode)
  PIC_C16M = 24                   'BMP default  colour implicit in pixeldata  (no Grey Mode)
  PIC_C16M32 = 32                 'BMP          colour implicit in pixeldata  (no Grey Mode)
End Enum

Public Enum IMG_CMAPMODES         'CMAP Actions *see also mValidate.bas for the Internal version of these
           PIC_MS_CMAP = &H2100&  'Just use the Cmap passed on from GDI or REMAP using MS API
      PIC_OPTIMAL_CMAP = &H1000&  'Make an optimal palette to correspond with BPP (use 24bits to n bits)
        PIC_FIXED_CMAP = &H2000&  'Use a nColor Fixed Color map
    PIC_FIXED_CMAP_VGA = &H200A&  'Force the VGA Colormap
   PIC_FIXED_CMAP_INET = &H200B&  'Force the 216 InternetSafe colormap (6*6*6)
  PIC_FIXED_CMAP_MS256 = &H200D&  'Force the MS  256 colormap (8,8,4)
   PIC_FIXED_CMAP_GREY = &H2200&  'Use a nColor Grey map              (FIXED + GREY)
     PIC_MODIFIED_CMAP = &H3000&  'Use a Modified nColor Color map  (OPTIMAL + FIXED)
   PIC_FIXED_CMAP_USER = &H6000&  'Use a user supplied nColorMap      (FIXED + USER)  (see SetCustomCmap())
End Enum

Public Enum IMG_DITHERMODES 'DITHER Actions
  PIC_DITHER_NONE = 0
  PIC_DITHER_BIN = &H1&          'Dither in a binary method
  PIC_DITHER_ORD = &H2&          'Dither to an ordered dithermatrix (default)
  PIC_DITHER_HTC = &H4&          'Dither to a printer Halftone
  PIC_DITHER_FDIAG = &H8&        'Dither to an forward diagonal dithermatrix
  PIC_DITHER_BDIAG = &H10&       'Dither to an backward diagonal dithermatrix
  PIC_DITHER_HORZ = &H20&        'Dither to an horizontal dithermatrix
  PIC_DITHER_VERT = &H40&        'Dither to an vertical dithermatrix
                                 'FS1,2,3 are retained for continuity - the matrices though have changed
  PIC_DITHER_FS1 = &H100&        'Serpentine Error Diffusion dithering Ulichney-Floyd-Steinberg Matrix
  PIC_DITHER_FS2 = &H200&        'Serpentine Error Diffusion dithering RVT 4 fixed coefficients
  PIC_DITHER_FS3 = &H400&        'Serpentine Error Diffusion dithering Ostromoukhov 3 Variable Coefficients
  PIC_DITHER_SED1 = &H100&       'Serpentine Error Diffusion dithering Floyd-Steinberg Matrix
  PIC_DITHER_SED2 = &H200&       'Serpentine Error Diffusion dithering RVT 4 fixed coefficients
  PIC_DITHER_SED3 = &H400&       'Serpentine Error Diffusion dithering Ostromoukhov 3 Variable Coefficients
  
  PIC_DITHER_BNM = &H800&        'Blue Noise Mask ( my algorithm, reasonable approximation so far)
End Enum

Public Enum IMG_PICPROC_COMMANDS  'These Codes are not actually used by external calls
  'The constants here have the form
  'BYTE 3 special control codes, BYTE 2 API OpCodes, BYTE 1 Command Nr, BYTE 0 The number of parameters it uses
  '00 = Ignored
  '01 = APIAssisted
  '02 = Requires      24/32bpp (so unmap if necessary)
  '04 = Requires   16/24/32bpp (so unmap if necessary)
  '08 = Requires 8/16/24/32bpp (so unmap if necessary)
  '10 = IMAGE OPERATION
  '20 = COLOR OPERATION
  '40 = TWO IMAGE OPERATION
  '80 DO NOT USE

  'Image Operations
          PIC_FLIP_VERT = &H10000100 'flip image vertically
          PIC_FLIP_HORZ = &H10000200 'flip image horizontally
        PIC_ROTATE_LEFT = &H18000300 'Rotate 90deg CCW
       PIC_ROTATE_RIGHT = &H18000400 'Rotate 90deg CW
         PIC_ROTATE_180 = &H10000500 'Rotate 180deg :-)
       PIC_CLIP_CENTRED = &H18000602 'Clip image to Centre ie to parms Long New Width,Height
            PIC_EXTRACT = &H18000704 'Extract a new Image from an Image - at parm1=x,parm2=y,parm3=Width,p4=Height
               PIC_TRIM = &H18000804 'Clip Image by removing Top,Left,Bottom,Right pixels
              PIC_FRAME = &H18000903 'Frame Image by replacing LeftRight,TopBottom pixels with FrameRGBColor
             PIC_RESIZE = &H18000A02 'Integer Resize to parm1=NewWidth, parm2=NewHeight
               PIC_ZOOM = &H18000B02 'Integer Zoom image by parm1 Single=XZoom times,parm2 Single=YZoom Times
                                     'if 0<xzoom,yzoom<1 then shrink,>1 enlarge
    PIC_BILINEAR_RESIZE = &H12000C02 'Resize by Bilinear Resampling to parm1=NewWidth, parm2=NewHeight
      PIC_BILINEAR_ZOOM = &H12000D02 'Zoom image Bilinear Resampling by parm1single =XZoom times,parm2 single=YZoom Times
                                     'if 0<xzoom,yzoom<1 then shrink,>1 enlarge
            PIC_SHARPEN = &H18000E01 'Sharpen Image by Factor 1-9
               PIC_BLUR = &H18000F01 'Blur Image    by Factor 1-9
             PIC_DEJPEG = &H18001001 'Remove JPEG block artifacts by factor 1-4
             PIC_EMBOSS = &H18001101 'Emboss Image  by Factor 1-9
               PIC_EDGE = &H18001201 'Show Edges    by Factor 1-9
           PIC_HARDEDGE = &H18001302 'Edge and Threshold Edges Factor 1-9 and Threshold 0-255
  
  'Color Operations
        PIC_UNMAP_COLOR = &H20000101 'Unmap image (to 24BPP) if mapped, (and remap afterwards)
        PIC_MSMAP_COLOR = &H20000202 'Parm1 = BPP required only 1,4,8,16,24,32 allowed
       PIC_INVERT_COLOR = &H20000300 'Invert all colors 00=FF etc
         PIC_MASK_COLOR = &H20000401 'Mask all Values to Parm1=RGBColor components
     PIC_SEPARATE_COLOR = &H20000501 'Leave a separation of colours like parm1=RGBColor
     PIC_EQUALIZE_COLOR = &H28000600 'Attempt to Equalise Colors
    PIC_THRESHOLD_COLOR = &H20000701 'Threshold Image to RGBColor Value if > then value
      PIC_GAMMA_CORRECT = &H20000803 'Gamma Correct Image three Parameters All Single RGamma,GGamma,BGamma
 PIC_BRIGHTNESS_CORRECT = &H20000903 'Brightness Correct Image three Parameters All Single RFactor,GFactor,BFactor
   PIC_ALPHABLEND_COLOR = &H20000A02 'AlphaBlend a parm1=RGBcolor by factor Alpha Parm1 single 0.0-1.0
      PIC_COMBINE_COLOR = &H20000B01 'Combine With parm1=RGBColor,given Opcode=Parm2
      PIC_REPLACE_COLOR = &H20000C03 'Replace parm1=TargetRGBColor,parm2=radius,parm3=NewRGBColor
  
  'Two Image operations
             PIC_SPLICE = &H40000102 'Parm1 = 2nd Image, Parm2=WhichEdge
            PIC_OVERLAY = &H40000303 'Overlay Parm1=2nd Image onto Image - at parm2=x,parm3=y
   PIC_ALPHABLEND_IMAGE = &H40000402 'AlphaBlend two Images, parm2 Single=Factor Alpha
      PIC_COMBINE_IMAGE = &H40000502 'Combine parm1=Image - With given Opcode=Parm2
          
  'new commands (and their VB code)
End Enum

'When combining Color Components which range 0..255, according to Result = Pixel1 Opcode Pixel2
Public Enum PIC_COMBINATION_OPCODES
    
   PCC_COPY = 0     'r = P2
  PCC_BLACK = 1     'r = &H00
  PCC_WHITE = 2     'r = &HFF
    PCC_ALT = 3     'r = if P1>127 then P1 else P2
    
    PCC_ADD = 4     'r = MIN( P1+P2, 255)
    PCC_SUB = 5     'r = MAX( P1-P2,   0)
    
    PCC_MOD = 6     'r = MIN(   (P1.P2)\255, 255)
   PCC_MOD2 = 7     'r = MIN( (2.P1.P2)\255, 255)
   PCC_ADDM = 8     'r = MAX(MIN( P1 + P2 - P1.P2\255, 255), 0) 'Add Smooth
    
    PCC_AND = 9     'r = P1 AND P2
     PCC_OR = 10    'r = P1  OR P2
    PCC_XOR = 11    'r = P1 XOR P2
  
    PCC_MAX = 12    'r = MAX(P1,P2)
    PCC_MIN = 13    'r = MIN(P1,P2)
    PCC_AVE = 14    'r = (P1+P2)\2  === alphablend with factor 1/2
    
    PCC_DOT = 15    'r = SQR((P1*P1 + P2*P2)\2)
  
  'Parameters and result may be affected by ORing in these codes
   PCC_NOT1 = 16    'P1 = NOT P1
   PCC_NOT2 = 32    'P2 = NOT P2
   PCC_NOTR = 64    'r  = NOT r
End Enum

Public Enum PIC_SEPARATION_OPCODES
  PSC_BLACK = 128 + 0  'Max(not r,not g,not b)
  PSC_RED = 128 + 1    'G=B=0
  PSC_GREEN = 128 + 2  'R=B=0
  PSC_BLUE = 128 + 4   'R=G=0
  PSC_YELLOW = 128 + 3 'B=0
  PSC_MAGENTA = 128 + 5 'G=0
  PSC_CYAN = 128 + 6   'R=0
End Enum

Public Enum PIC_JOIN_EDGES
  PJE_LEFT = 1
  PJE_TOP = 2
  PJE_RIGHT = 3
  PJE_BOTTOM = 4
End Enum

Public Enum IMG_RENDEROPTIONS  '(additive bitstates 4,8,16)
  IRO_ASIS = 0                   'Put at CurrXY, and Without scaling
  IRO_STRETCH = 1                'Stretch to Fit Clip Area (ie AutoCentred)
  IRO_TILE = 2                   'Tile Image to Fit Clip Area (image is not Scaled)

  IRO_KEEPASPECT = 4             'Image Aspect Ratio is Kept during Stretch (wont necessarily fit to edges)
  IRO_INTSCALE = 8               'Image is scaled by nearest integer to StretchFit (wont necessarily fit to edges)
  IRO_CENTRE = 16                'Image is put in the centre of the clip
End Enum

Public Type RGBA    'all values 0..255  'This is a synonym of PALETTENTRY
  Red    As Byte                 'the red value
  Green  As Byte                 'the green value
  Blue   As Byte                 'the blue value
  Alpha  As Byte                 'the grey value of the pixel (aka ALPHA)
End Type

'Internally Used
Public Enum IMG_PICSTATES      'additive bitstates (read only)
  IS_VALID_PIPELINE = &H1&
  IS_VALID_CLIP = &H2&           'the clipping rectangle is inside the image
  IS_CMAPPED = &H4&              'Colormapping has taken place (all pixels are 1,4 or 8bit indices to CMap)
  IS_TOP_TO_BOTTOM = &H8&        'the image is topdown oriented, not bottom up as a BMP usually is
  IS_MS_FORMAT = &H10&           'the current result image is Mapped 1,4,8 or 16,24,32 bits
End Enum

'=============================================================================================================
'============================= THE BASIC ELEMENTS OF THE IMAGE ================================================
'=============================================================================================================

Private zPicState     As Long         'see enum above
Private zWidth        As Long         'in pixels
Private zHeight       As Long         'in pixels
Private zRowModulo    As Long         'in Bytes
Private zPixBits()    As Byte         'the array of pixel bits, may be indices into the CMap
Private zIntBPP       As Long         'The Internal BitsPerPixel (1,4,8,16,24 for BMPS, =24 for GIF,JPG,PNM)
Private zCMap()       As RGBA         'the colormap in A,B,G,R tuples of 0..255  (MUST BE NCMapColors)
Private zNCMapColors  As Long         'The number of Colors in the CMap, 2^BitsPerPixel

Private zUserCMAP()   As RGBA         'will be used if CMAPMode=PIC_FIXED_CMAP_USER
                                      'your problem is to make it the right size

Private zPicType      As Long         'BMP or GIF
Private zColorMode    As Long         'COLOR, BW, GREY
Private zReqBPP       As Long         'The requested format in Bits Per Pixel
Private zCMAPMode     As Long         'See Enum above
Private zDitherMode   As Long         'See Enum Above
Private zTransparentC As Long         'The Transparent color in GIFs (if in final GIF)
                
'for DemoPurposes Only
Private zTimer        As cTimingPC

'Elapsed times in seconds
Public etLoad         As Single    'Time To Load/Capture
Attribute etLoad.VB_VarDescription = "Elapsed time (secs) to load an Image"
Public etOps          As Single    'Time Operations
Attribute etOps.VB_VarDescription = "Elapsed Time for Operations"
Public etQuant        As Single    'Time Quantize colors
Attribute etQuant.VB_VarDescription = "Elapsed Time for Quantization"
Public etRemap        As Single    'Time Remap/Dither
Attribute etRemap.VB_VarDescription = "Elaped tme to Remap and Image to a given ColorMap"
Public etSave         As Single    'Time To Save Files
Attribute etSave.VB_VarDescription = "Elapsed TIme (secs) to save an Image"
Public etElapsed      As Single    'Total Time (for Major Functions)
Attribute etElapsed.VB_VarDescription = "Elapsed time (secs) for the entire saving process"

'=============================================================================================================
'=================================== INIT/TERMINATE ==========================================================
'=============================================================================================================

Private Sub Class_Initialize()
  
  Set zTimer = New cTimingPC
  zTimer.Reset
  Call InitValidationArrays     'mValidate.bas
  Call InitBMPScales            'mBBMFileIO.bas
  zIntBPP = PIC_24BPP
  zReqBPP = PIC_24BPP
  zPicState = IS_VALID_PIPELINE
  
End Sub

Private Sub Class_Terminate()
  
  Erase zCMap(), zPixBits()
  Set zTimer = Nothing
  
End Sub

'=============================================================================================================
'=============================== PROPERTIES ===================================================================
'=============================================================================================================

Public Property Get Width() As Long
Attribute Width.VB_Description = "return Width of Image in Pixels"
  
  Width = zWidth

End Property

Public Property Get Height() As Long
Attribute Height.VB_Description = "Return Height ogfImage in Pixels"
  
  Height = zHeight

End Property

Public Property Get BitsPerPixel() As Long
Attribute BitsPerPixel.VB_Description = "Depth of Image in Bits"
  
  BitsPerPixel = zIntBPP
  
End Property

Public Property Get ReqBPP() As Long
Attribute ReqBPP.VB_Description = "The Required Bitmap Depth in BitsPerPixel"
  
  ReqBPP = zReqBPP
  
End Property

'DstPixBits should be dynamically allocated
Public Function GetPixBits(ByRef DstPixBits() As Byte) As Boolean
Attribute GetPixBits.VB_Description = "Return the Pixel Bitdata"

  On Error Resume Next
  DstPixBits() = zPixBits()
  If Err.Number = 0 Then GetPixBits = True
  On Error GoTo 0

End Function

Friend Property Get PixBitsBasePtr() As Long    'return the BaseAddress of the zPixBits() Array

  PixBitsBasePtr = VarPtr(zPixBits(0))

End Property

Public Property Get RowModulo() As Long
Attribute RowModulo.VB_Description = "The Real width in bytes of a bitmap (ie Nr of Pixels in Bytes plus padding "
  
  RowModulo = BMPRowModulo(zWidth, zIntBPP)

End Property

'DstCMap should be dynamically allocated
Public Function GetCMap(ByRef DstCMap() As RGBA) As Boolean
Attribute GetCMap.VB_Description = "Return current ColorMap"

  On Error Resume Next
  DstCMap() = zCMap()
  If Err.Number = 0 Then GetCMap = True
  On Error GoTo 0

End Function

Public Property Get NCMapColors() As Long
Attribute NCMapColors.VB_Description = "Return the Nr of colors in the current ColorMap"
  
  NCMapColors = zNCMapColors

End Property

'set aside a user prepared colormap - it should match NCMapColors (users responsibility to check)
Public Property Let UserCMap(ByRef UserCMap() As RGBA)
Attribute UserCMap.VB_Description = "A user defined ColorMap"

  On Error Resume Next
  zUserCMAP() = UserCMap()
  On Error GoTo 0

End Property

'These are internally maintained
Public Property Get PicState() As IMG_PICSTATES
Attribute PicState.VB_Description = "Internal State Flags"
  
  PicState = zPicState

End Property

Public Property Get PicType() As IMG_PICTYPE
Attribute PicType.VB_Description = "BMP, GIF or PNM"
  
  PicType = zPicType

End Property

Public Property Get ColorMode() As IMG_COLORMODES
Attribute ColorMode.VB_Description = "The Color Type of the Picture - Color, Black and White or Grey"
  
  ColorMode = zColorMode

End Property

Public Property Get CMAPMode() As IMG_CMAPMODES
Attribute CMAPMode.VB_Description = "The Method of Color Mapping that is in effect"
  
  CMAPMode = zCMAPMode

End Property

Public Property Get DitherMode() As IMG_DITHERMODES
Attribute DitherMode.VB_Description = "Which Method is used to Dither colours"
  
  DitherMode = zDitherMode

End Property

'Set this property to make a GIF transparent , set it with -1 to stop this happening
Public Property Let TransparentGIFColor(ByVal vdata As Long)
Attribute TransparentGIFColor.VB_Description = "Which shade will be considered transparent in the gif (colorDEF)"

  If vdata = -1 Then
    zTransparentC = -1
   Else
    zTransparentC = (vdata And &HFFFFFF)
  End If

End Property

Public Property Get TransparentGIFColor() As Long

  TransparentGIFColor = zTransparentC

End Property

'=============================================================================================================
'======================================= M E T H O D S =======================================================
'=============================================================================================================
'Make a complete copy of an image
Public Function Clone() As cRVTVBIMG    'Nothing if Failed to Copy
Attribute Clone.VB_Description = "Create an Exact Copy of the Image as a New Image"
  
  Set Clone = New cRVTVBIMG
  If Not Clone.CloneFrom(Me) Then Set Clone = Nothing
  
End Function

'As an Image is strictly ReadOnly apart from sanctioned methods so we clone From the Original here
Friend Function CloneFrom(ByVal RefImage As cRVTVBIMG) As Boolean
  
  With RefImage
    zPicState = .PicState
    zWidth = .Width
    zHeight = .Height
    zIntBPP = .BitsPerPixel
    CloneFrom = .GetPixBits(zPixBits())
    If CloneFrom Then CloneFrom = .GetCMap(zCMap())
    zNCMapColors = .NCMapColors
    zPicType = .PicType
    zColorMode = .ColorMode
    zReqBPP = .ReqBPP
    zCMAPMode = .CMAPMode
    zDitherMode = .DitherMode
    zTransparentC = .TransparentGIFColor
  End With
  
End Function

'Make a cleared Image, optionally of a different size, always cleared to Black
Public Function EraseImage(Optional ByVal NewWidth As Long = 0, _
                           Optional ByVal NewHeight As Long = 0, _
                           Optional ByVal NewBPP As Long = 0) As Boolean
Attribute EraseImage.VB_Description = "Clear and Optionally redefine Image Dimensions"

  Call SetF(zPicState, IS_VALID_PIPELINE Or IS_TOP_TO_BOTTOM)
  NewWidth = Abs(NewWidth)
  NewHeight = Abs(NewHeight)
  If NewWidth = 0 Then NewWidth = zWidth
  If NewHeight = 0 Then NewHeight = zHeight
  If NewWidth = 0 Then NewWidth = 256
  If NewHeight = 0 Then NewHeight = 256
  zWidth = NewWidth
  zHeight = NewHeight
  If NewBPP <= PIC_24BPP Then
    zIntBPP = PIC_24BPP
  Else
    zIntBPP = NewBPP
  End If
  ReDim zPixBits(0 To zHeight * BMPRowModulo(zWidth, zIntBPP) - 1)
  If zReqBPP = 0 Then zReqBPP = PIC_24BPP
  Erase zCMap()
  zNCMapColors = 0
  EraseImage = True
  
End Function

'Return a Pixel at Offset from PixBits (NOTE: NO ERROR CHECKING HERE AT ALL)
Friend Function GetPixelz(ByVal Offset As Long) As RGBA
  
  Dim w As Long, Pixel As RGBA
  
  Select Case zIntBPP
    Case PIC_8BPP:
      GetPixelz = zCMap(zPixBits(Offset))
      
    Case PIC_24BPP:
      With GetPixelz
        .Blue = zPixBits(Offset)
        .Green = zPixBits(Offset + 1)
        .Red = zPixBits(Offset + 2)
      End With
      
    Case PIC_32BPP:
      With GetPixelz
        .Blue = zPixBits(Offset)
        .Green = zPixBits(Offset + 1)
        .Red = zPixBits(Offset + 2)
        .Alpha = zPixBits(Offset + 3)
      End With
      
    Case PIC_16BPP:   '5,5,5
      w = zPixBits(Offset) Or 256& * zPixBits(Offset + 1)
      With GetPixelz
        .Blue = BMP31Scale((w And &H1F&))
        .Green = BMP31Scale((w And &H3E0&) \ 32&)
        .Red = BMP31Scale((w And &H7C00&) \ 1024&)
      End With
  End Select
    
End Function

'get the value of a pixel as a VBColor ie Long as &HAABBGGRR
Public Function GetPixel(ByVal x As Long, ByVal y As Long) As Long    '-1 is outside image
Attribute GetPixel.VB_Description = "Return the ColorValue  at given Pixel"
  
  Dim p As Long, cc As RGBA
  
  GetPixel = -1
  If zIntBPP >= PIC_16BPP Then                       'is not mapped
    If x >= 0 And x < zWidth Then                                 'x in range
      If y >= 0 And y < zHeight Then                              'y in range
        p = y * BMPRowModulo(zWidth, zIntBPP) + x * (zIntBPP \ 8)
        Select Case zIntBPP
          Case PIC_24BPP:
            cc.Blue = zPixBits(p)
            cc.Green = zPixBits(p + 1)
            cc.Red = zPixBits(p + 2)
          
          Case PIC_32BPP:
            cc.Blue = zPixBits(p)
            cc.Green = zPixBits(p + 1)
            cc.Red = zPixBits(p + 2)
            cc.Alpha = zPixBits(p + 3)
            
          Case PIC_16BPP:
            cc = GetPixel16(zPixBits, (p))
        
        End Select
        GetPixel = PutRGBA(cc)
      End If
    End If
  Else        'its a mapped image
    p = GetPixelIndex(x, y)         'get index
    If p <> -1 Then
      GetPixel = PutRGBA(zCMap(p))  'return its color
    End If
  End If
  
End Function

'set the value of a pixel if image is not mapped -1 is returned
Public Function SetPixel(ByVal x As Long, ByVal y As Long, ByVal RGBColor As Long) As Long '-1 is outside image
Attribute SetPixel.VB_Description = "Set the color of a given Pixel if it is not colormapped 16,24,32 BPP"
  
  Dim p As Long, cc As RGBA
  
  SetPixel = -1
  If zIntBPP >= PIC_16BPP Then                       'is not mapped
    If x >= 0 And x < zWidth Then                             'x in range
      If y >= 0 And y < zHeight Then                          'y in range
        cc = GetRGBA(RGBColor)
        p = y * BMPRowModulo(zWidth, zIntBPP) + x * (zIntBPP \ 8)
        Select Case zIntBPP
          Case PIC_24BPP:
            zPixBits(p) = cc.Blue
            zPixBits(p + 1) = cc.Green
            zPixBits(p + 2) = cc.Red
          
          Case PIC_32BPP:
            zPixBits(p) = cc.Blue
            zPixBits(p + 1) = cc.Green
            zPixBits(p + 2) = cc.Red
            zPixBits(p + 3) = cc.Alpha
            
          Case PIC_16BPP:
            Call PutPixel16(zPixBits, (p), cc)
        
        End Select
        SetPixel = RGBColor
      End If
    End If
  End If
  
End Function

'get the value of a pixel index for Mapped Images
Public Function GetPixelIndex(ByVal x As Long, ByVal y As Long) As Long    '-1 is outside image
Attribute GetPixelIndex.VB_Description = "Get ColorIndex of a given Pixel if it is ColorMapped"
  
  Dim p As Long, q As Long
  
  GetPixelIndex = -1
  If zIntBPP <= PIC_8BPP Then                       'is mapped
    If x >= 0 And x < zWidth Then                             'x in range
      If y >= 0 And y < zHeight Then                          'y in range
        p = y * BMPRowModulo(zWidth, zIntBPP)
        Select Case zIntBPP
          Case PIC_8BPP:
            GetPixelIndex = zPixBits(p + x)
            
          Case PIC_4BPP:
            If (x And 1) = 1 Then
              GetPixelIndex = (zPixBits(p + x \ 2) And &HF)
            Else
              GetPixelIndex = (zPixBits(p + x \ 2) And &HF0) \ 16
            End If
          
          Case PIC_1BPP:
            If (CLng(zPixBits(p + x \ 8)) And (Not 2 ^ CLng(7 - x Mod 8))) = 0 Then
              GetPixelIndex = 0
            Else
              GetPixelIndex = 1
            End If
        End Select
      End If
    End If
  End If
  
End Function

'set the value of a pixel by its color Index if Image is not mapped -1 is returned
Public Function SetPixelIndex(ByVal x As Long, ByVal y As Long, ByVal ColorIndex As Long) As Long '-1 is outside image
Attribute SetPixelIndex.VB_Description = "Set the colorindex of a pixel if it is ColorMapped"
  
  Dim p As Long, q As Long
  
  SetPixelIndex = -1
  If zIntBPP <= PIC_8BPP Then                       'is mapped
    If x >= 0 And x < zWidth Then                             'x in range
      If y >= 0 And y < zHeight Then                          'y in range
        If ColorIndex >= 0 And ColorIndex < NCMapColors Then  'colorindex in range
          p = y * BMPRowModulo(zWidth, zIntBPP)
          Select Case zIntBPP
            Case PIC_8BPP:
              zPixBits(p + x) = ColorIndex
              
            Case PIC_4BPP:
              If (x And 1) = 1 Then
                zPixBits(p + x \ 2) = (zPixBits(p + x \ 2) And &HF0) Or ColorIndex
              Else
                zPixBits(p + x \ 2) = (zPixBits(p + x \ 2) And &HF) Or (ColorIndex * 16)
              End If
            
            Case PIC_1BPP:
              q = 2 ^ (7 - x Mod 8)
              If ColorIndex = 1 Then
                q = (zPixBits(p + x \ 8) And (Not q)) Or q
              Else
                q = (zPixBits(p + x \ 8) And (Not q))
              End If
              zPixBits(p + x \ 8) = (q And &HFF)
          End Select
          SetPixelIndex = ColorIndex
        End If
      End If
    End If
  End If
  
End Function

'----------------------------- Passing Complete Image data between You and Me --------------------------------

'Takes a filled in BM and Captures it Internally
'It is the callers responsibility to divest itself of hUserBM
'This may fail for lack of memory - always captured as 24BPP
Public Function PutImageBitMap(ByVal hUserBM As Long) As Boolean     'TRUE if Filled in correctly
Attribute PutImageBitMap.VB_Description = "Create and return a bitmap from the current Image"
  
  
  If hUserBM <> 0 Then
    zIntBPP = 24
    PutImageBitMap = CaptureBitmap(hUserBM, zPicState, zWidth, zHeight, _
                                            zPixBits(), zIntBPP, zCMap(), zNCMapColors)
  End If
  
End Function

'Clones the Internal Image (if it is a valid MS Format) and returns it in a pointer
'It is the callers responsibility to divest itself of hUserBM
'This may fail for lack of memory - a 5,6,5 16bit bitmap cannot be made, all others OK
Public Function GetImageBitMap(ByRef hUserBM As Long) As Boolean
Attribute GetImageBitMap.VB_Description = "Construct an image from a Passed in Bitmap"
  
  Dim tmpDC As Long
  
  GetImageBitMap = DCBitMapFromImage(zWidth, zHeight, _
                                     zPicState, zPixBits(), zIntBPP, zCMap(), zNCMapColors, _
                                     tmpDC, hUserBM)
  If tmpDC <> 0 Then Call DeleteObject(tmpDC)

End Function

'------------------------------ Data to and From StdPicture --------------------------------------------------

'Creates an Image from the Passed in StdPicture
Public Function PutStdPicture(ByRef Pic As StdPicture) As Boolean
Attribute PutStdPicture.VB_Description = "Create and return a StdPicture of the current image"
  
  PutStdPicture = CaptureBitmap(Pic.Handle, zPicState, zWidth, zHeight, _
                                            zPixBits(), zIntBPP, zCMap(), zNCMapColors)

End Function

'Returns a StdPicture from the Internal Image  (based on Original Code by Vlad Vissoultchev.)
Public Function GetStdPicture() As StdPicture   'Nothing if it Failed
Attribute GetStdPicture.VB_Description = "Create Image from  passed in STDpicture"
 
 Dim sp As StdPicture, PicDesc As PictDesc, IID_IDispatch As GUID
 Dim tmpDC As Long, tmpBM As Long, OldSMode As Long

 If DCBitMapFromImage(zPicState, zWidth, zHeight, zPixBits(), zIntBPP, zCMap(), zNCMapColors, tmpDC, tmpBM) Then
    '-- Fill image info
    With PicDesc
      .size = Len(PicDesc)
      .Type = vbPicTypeBitmap
      .hBmp = tmpBM
      .hPal = 0
    End With
    
    '-- Fill GUID info
    With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
    End With
    
    '-- Create picture from bitmap handle
    Call OleCreatePictureIndirect(PicDesc, IID_IDispatch, 1, sp)
    Set GetStdPicture = sp
      
    '-- Cleanup
    Call FreeDCandBitmap(tmpDC, tmpBM)
  End If
  
End Function

'------------------------------ Data to and From an Obj.hDC --------------------------------------------------

'Capture an Image From an Obj.hDC

Public Function GetImageObjhDC(ByVal Obj As Object, _
                               Optional ByVal ClipLeft As Long = -1, _
                               Optional ByVal ClipTop As Long = -1, _
                               Optional ByVal ClipRight As Long = -1, _
                               Optional ByVal ClipBottom As Long = -1) As Boolean
Attribute GetImageObjhDC.VB_Description = "Get the Image data from a given hDC"


  If FSet(zPicState, IS_VALID_PIPELINE) Then       'must have a valid pipeline to proceed
    GetImageObjhDC = ImagefromObjhDC(Obj, zPicType, zPicState, zWidth, zHeight, _
                                          zPixBits(), zIntBPP, zCMap(), zNCMapColors, _
                                          ClipLeft, ClipTop, ClipRight, ClipBottom)
  End If

End Function

'Render the Image Stored to a given Obj.hDC [clipped]
Public Function PutImageObjhDC(ByVal Obj As Object, _
                               Optional ByVal ImageRenderOptions As IMG_RENDEROPTIONS = IRO_ASIS, _
                               Optional ByVal ClipLeft As Long = -1, _
                               Optional ByVal ClipTop As Long = -1, _
                               Optional ByVal ClipRight As Long = -1, _
                               Optional ByVal ClipBottom As Long = -1) As Boolean
Attribute PutImageObjhDC.VB_Description = "Put the image onto a hDC ie. Render"

  PutImageObjhDC = RenderImage(Obj, zPicState, zWidth, zHeight, _
                                    zPixBits(), zIntBPP, zCMap(), zNCMapColors, _
                                    ImageRenderOptions, ClipLeft, ClipTop, ClipRight, ClipBottom)
End Function

'AUTOPILOT Render the Image Stored to a given DC (clipped and centred)
Public Function DisplayImage(ByVal Obj As Object) As Boolean
Attribute DisplayImage.VB_Description = "Render Image to an hDC"

  DisplayImage = PutImageObjhDC(Obj)

End Function

'AUTOPILOT OBJ.hDC SAVE
' The whole or part of an Object hDC will be saved according to your passed-in preferences
' The simplest possible call is
'    SaveObjDCCLip(Obj,Path) which saves the whole Obj.hDC as a 256Color Bitmap to Path
'  a return code of false indicates something went wrong.

Public Function SaveObjDCClip(ByVal Obj As Object, _
                              ByRef DestPath As String, _
                              Optional ByVal PicType As IMG_PICTYPE = PIC_BMP, _
                              Optional ByVal ColorMode As IMG_COLORMODES = PIC_COLOR, _
                              Optional ByVal BitsPerPixel As IMG_DEPTHMODES = PIC_8BPP, _
                              Optional ByVal CMAPMode As IMG_CMAPMODES = PIC_OPTIMAL_CMAP, _
                              Optional ByVal DitherMode As IMG_DITHERMODES = PIC_DITHER_NONE, _
                              Optional ByVal ClipLeft As Long = -1, _
                              Optional ByVal ClipTop As Long = -1, _
                              Optional ByVal ClipRight As Long = -1, _
                              Optional ByVal ClipBottom As Long = -1) As Boolean
Attribute SaveObjDCClip.VB_Description = "Capture and Save the Bitmap attached to Object hDC and Save according to Specification"

  etElapsed = zTimer.Elapsed
  If Not Obj Is Nothing Then
    If Len(DestPath) <> 0 Then
      If SetPipeline(PicType, ColorMode, BitsPerPixel, CMAPMode, DitherMode) Then  'set the pipeline
        If GetImageObjhDC(Obj, ClipLeft, ClipTop, ClipRight, ClipBottom) Then      'get the basic image data
          If DoColorMapping() Then                                                 'Post Processing
            SaveObjDCClip = CommitToDisk(DestPath, PicType)                        'File IO
          End If
        End If
      End If
    End If
  End If
  etElapsed = zTimer.Elapsed - etElapsed
  
End Function

'----------------------------- Data to and From Disk Files -----------------------------------------------------

Public Function LoadFromFile(ByRef FilePath As String) As Boolean
Attribute LoadFromFile.VB_Description = "Load an Image from a file - BMP, JPG,GIF or PNM (PBM,PGM,PPM)"

  Dim tmpPic As StdPicture
  
  etLoad = zTimer.Elapsed
  If Len(FilePath) <> 0 Then
  
    'Try BMP
    LoadFromFile = LoadBMP(FilePath, zPicState, zWidth, zHeight, zPixBits, zIntBPP, zCMap(), zNCMapColors)
    
    If Not LoadFromFile Then  'Try PNM
      LoadFromFile = LoadPNM(FilePath, zPicState, zWidth, zHeight, zPixBits, zIntBPP, zCMap(), zNCMapColors)
    End If
        
    If Not LoadFromFile Then  'use emergency measures  (StdPic) GIF,JPG,JPEG
      On Error Resume Next
      Set tmpPic = LoadPicture(FilePath)
      On Error GoTo 0
      If Not tmpPic Is Nothing Then LoadFromFile = PutStdPicture(tmpPic)
    End If
  End If
    
  etLoad = zTimer.Elapsed - etLoad

End Function

'This function is a one step save routine, Current Image must be already loaded
Public Function SavetoFile(ByRef DestPath As String, _
                           Optional ByVal PicType As IMG_PICTYPE = PIC_BMP, _
                           Optional ByVal ColorMode As IMG_COLORMODES = PIC_COLOR, _
                           Optional ByVal BitsPerPixel As IMG_DEPTHMODES = PIC_8BPP, _
                           Optional ByVal CMAPMode As IMG_CMAPMODES = PIC_OPTIMAL_CMAP, _
                           Optional ByVal DitherMode As IMG_DITHERMODES = PIC_DITHER_NONE) As Boolean '(true is Good)
Attribute SavetoFile.VB_Description = "Save Image to File"

  etLoad = 0
  etElapsed = zTimer.Elapsed
  If Len(DestPath) <> 0 Then
    If SetPipeline(PicType, ColorMode, BitsPerPixel, CMAPMode, DitherMode) Then ' Set the pipeline
      If DoColorMapping() Then                                                  ' Post Processing
        SavetoFile = CommitToDisk(DestPath, PicType)                                  ' File IO
      End If
    End If
  End If
  etElapsed = zTimer.Elapsed - etElapsed

End Function

'This function is a BMP,GIF,JPG,PNM one step conversion routine
'Load the Image and Save it again (with requested parms)
Public Function ConvertImageFile(ByRef SourcePath As String, _
                                 ByRef DestPath As String, _
                                 Optional ByVal PicType As IMG_PICTYPE = PIC_BMP, _
                                 Optional ByVal ColorMode As IMG_COLORMODES = PIC_COLOR, _
                                 Optional ByVal BitsPerPixel As IMG_DEPTHMODES = PIC_8BPP, _
                                 Optional ByVal CMAPMode As IMG_CMAPMODES = PIC_OPTIMAL_CMAP, _
                                 Optional ByVal DitherMode As IMG_DITHERMODES = PIC_DITHER_NONE) As Boolean
Attribute ConvertImageFile.VB_Description = "Take an Imagefile of One Format and Convert it to a second"

  etElapsed = zTimer.Elapsed
  If Len(SourcePath) <> 0 And Len(DestPath) <> 0 Then
    If SetPipeline(PicType, ColorMode, BitsPerPixel, CMAPMode, DitherMode) Then    'Set the pipeline
      If LoadFromFile(SourcePath) Then                  ' Get The Image we want To Convert Must be a valid image
        If DoColorMapping() Then                        ' Post Processing
          ConvertImageFile = CommitToDisk(DestPath, PicType)  ' File IO
        End If
      End If
    End If
  End If
  etElapsed = zTimer.Elapsed - etElapsed
End Function

'==================================================================================================================
'HERE FOLLOW the individual steps of the Pipeline as applied by the SaveOBJDCCLip
'SetPipeline should always be called BEFORE capture of individual items
'==================================================================================================================
'Just calling this will give you a pipeline of PIC_BMP,PIC_COLOR,PIC_8BPP,PIC_OPTIMAL_CMAP,PIC_DITHER_NONE
Public Function SetPipeline(Optional ByVal PicType As IMG_PICTYPE = PIC_BMP, _
                            Optional ByVal ColorMode As IMG_COLORMODES = PIC_COLOR, _
                            Optional ByVal BitsPerPixel As IMG_DEPTHMODES = PIC_8BPP, _
                            Optional ByVal CMAPMode As IMG_CMAPMODES = PIC_OPTIMAL_CMAP, _
                            Optional ByVal DitherMode As IMG_DITHERMODES = PIC_DITHER_NONE) As Boolean
Attribute SetPipeline.VB_Description = "Declare the methods to be used when processing an Image"

  SetPipeline = QualifyPipeline(PicType, ColorMode, BitsPerPixel, CMAPMode, DitherMode)

End Function

'The pipeline can also be set from the Ordinals of the options within each Enum (-1) uses a default
'Just calling this will give you a pipeline of PIC_BMP,PIC_COLOR,PIC_8BPP,PIC_OPTIMAL_CMAP,PIC_DITHER_NONE
Public Function SetPipelineOrdinal(Optional ByVal PicTypeOrdinal As Long = -1, _
                                   Optional ByVal ColorModeOrdinal As Long = -1, _
                                   Optional ByVal BitsPerPixelOrdinal As Long = -1, _
                                   Optional ByVal CMAPModeOrdinal As Long = -1, _
                                   Optional ByVal DitherModeOrdinal As Long = -1) As Boolean
Attribute SetPipelineOrdinal.VB_Description = "Declare the methods to be used when processing an Image by passing in th eordinal of the option required"

  Call PicEnumFromOrdinals(PicTypeOrdinal, ColorModeOrdinal, BitsPerPixelOrdinal, CMAPModeOrdinal, DitherModeOrdinal)
  SetPipelineOrdinal = QualifyPipeline(PicTypeOrdinal, ColorModeOrdinal, BitsPerPixelOrdinal, CMAPModeOrdinal, DitherModeOrdinal)

End Function

'========= ELEMENTERY IMAGE PROCESSING ===== New Options should be added Here =================================

'They should all follow the convention that rc=False indicates failure, True=success

Public Function FlipVert() As Boolean
Attribute FlipVert.VB_Description = "Flip an image from top to bottom"

  FlipVert = FlipImageVert(zPicState, zWidth, zHeight, zIntBPP, zPixBits())   'deals with all MS formats

End Function

Public Function FlipHorz() As Boolean
Attribute FlipHorz.VB_Description = "Flip Image Right to Left"

  FlipHorz = FlipImageHorz(zWidth, zHeight, zIntBPP, zPixBits())              'deals with all MS formats

End Function

Public Function RotateL() As Boolean
Attribute RotateL.VB_Description = "Rotate Image 90 degrees Left"

  RotateL = RotateImageLeft(zPicState, zWidth, zHeight, zIntBPP, zPixBits())  '24 and 32BPP

End Function

Public Function RotateR() As Boolean
Attribute RotateR.VB_Description = "Rotate Image 90 degrees Right"

  RotateR = RotateImageRight(zPicState, zWidth, zHeight, zIntBPP, zPixBits()) '24 and 32BPP

End Function

Public Function Rotate180() As Boolean
Attribute Rotate180.VB_Description = "Rotate the image through 180 degrees"

  Rotate180 = RotateImage180(zPicState, zWidth, zHeight, zIntBPP, zPixBits()) 'all Formats

End Function

'Clip about centre
Public Function Clip(ByVal NewWidth As Long, ByVal NewHeight As Long) As Boolean
Attribute Clip.VB_Description = "Clip an image by selecting subimage from its centre"

  Dim LeftClip As Long, TopClip As Long
  
  LeftClip = (zWidth - NewWidth) \ 2
  TopClip = (zHeight - NewHeight) \ 2
  Clip = Trim(LeftClip, TopClip, zWidth - (NewWidth + LeftClip), zHeight - (NewHeight + TopClip))

End Function

'NOTE::: This cannot be called through DoOperations as it Returns a new Image
Public Function Extract(ByVal atX As Long, ByVal atY As Long, _
                        ByVal Width As Long, ByVal Height As Long) As cRVTVBIMG
Attribute Extract.VB_Description = "Extract a sub Image from an Image. Will always be 24BPP"
                        
  Set Extract = ExtractImage(zWidth, zHeight, zIntBPP, zPixBits(), atX, atY, Width, Height)

End Function

'Trim strips off left,top,right and bottom
Public Function Trim(ByVal LeftClip As Long, ByVal TopClip As Long, _
                     ByVal RightClip As Long, ByVal BottomClip As Long) As Boolean
Attribute Trim.VB_Description = "Trim boundaries off an image"
  
  Trim = TrimImage(zWidth, zHeight, zIntBPP, zPixBits(), LeftClip, TopClip, RightClip, BottomClip)

End Function

'Frame
Public Function Frame(ByVal LeftRight As Long, TopBottom As Long, FrameRGBColor As Long) As Boolean
Attribute Frame.VB_Description = "Draw a Frame around the image with given color"
  
  Frame = FrameImage(zWidth, zHeight, zIntBPP, zPixBits(), LeftRight, TopBottom, FrameRGBColor)

End Function

'Integer Rescale
Public Function Resize(ByVal NewWidth As Long, ByVal NewHeight As Long) As Boolean
Attribute Resize.VB_Description = "Resize the image using nearest neigbour method"

  Resize = ResizeImage(zWidth, zHeight, zIntBPP, zPixBits(), NewWidth, NewHeight)

End Function

'Integer Rescale
Public Function Zoom(ByVal XZoom As Single, Optional ByVal YZoom As Single = 0) As Boolean
Attribute Zoom.VB_Description = "Zoom (rescale) an image using nearest neighbour method"

  If YZoom = 0 Then YZoom = XZoom
  Zoom = ResizeImage(zWidth, zHeight, zIntBPP, zPixBits(), zWidth * XZoom, zHeight * YZoom)

End Function

Public Function BilinearResize(ByVal NewWidth As Long, ByVal NewHeight As Long) As Boolean
Attribute BilinearResize.VB_Description = "Use Bilinear Interpolation to Resize an Image"

  BilinearResize = BilinearResizeImage(zWidth, zHeight, zIntBPP, zPixBits(), NewWidth, NewHeight)

End Function

Public Function BilinearZoom(ByVal XZoom As Single, Optional ByVal YZoom As Single = 0) As Boolean
Attribute BilinearZoom.VB_Description = "Use bilinear resampling to resize an image"

  If YZoom = 0 Then YZoom = XZoom
  BilinearZoom = BilinearResizeImage(zWidth, zHeight, zIntBPP, zPixBits(), zWidth * XZoom, zHeight * YZoom)

End Function

Public Function Sharpen(Optional ByVal Factor As Long = 1) As Boolean
Attribute Sharpen.VB_Description = "Enhance image by sharpening it"

  Sharpen = SharpenImage(zWidth, zHeight, zIntBPP, zPixBits(), Factor)

End Function

Public Function Blur(Optional ByVal Factor As Long = 1) As Boolean
Attribute Blur.VB_Description = "Blur an image"

  Blur = BlurImage(zWidth, zHeight, zIntBPP, zPixBits(), Factor)

End Function

'Correct JPG Quantization Errors (somewhat)
Public Function DeJPEG(Optional ByVal Factor As Long = 1) As Boolean
Attribute DeJPEG.VB_Description = "Reduce JPEG quantization errors (use Edge to Demonstrate effect)"

  If Factor < 1 Or Factor > 4 Then Factor = 1
  DeJPEG = Blur(Factor)
  If DeJPEG Then DeJPEG = Sharpen(2 * Factor)

End Function

Public Function Emboss(Optional ByVal Factor As Long = 1) As Boolean
Attribute Emboss.VB_Description = "Give a 3D effect to an Image"

  Emboss = EmbossImage(zWidth, zHeight, zIntBPP, zPixBits(), Factor)

End Function

Public Function Edge(Optional ByVal Factor As Long = 1) As Boolean
Attribute Edge.VB_Description = "Show the edges of an Image"

  Edge = EdgeImage(zWidth, zHeight, zIntBPP, zPixBits(), Factor)

End Function

Public Function HardEdge(Optional ByVal Factor As Long = 1, Optional ByVal Threshold As Long = 240) As Boolean
Attribute HardEdge.VB_Description = "Edge and Threshold"

  HardEdge = HardEdgeImage(zWidth, zHeight, zIntBPP, zPixBits(), Factor, Threshold)

End Function

'===================== TWO Image Procs ======================================================================

Public Function Splice(ByVal Image2 As cRVTVBIMG, Optional ByVal WhichEdge As PIC_JOIN_EDGES) As Boolean
Attribute Splice.VB_Description = "Join two Images according to given edge. Images MUST BE same Depth"
  
  Splice = SpliceImage(zWidth, zHeight, zIntBPP, zPixBits(), Image2, WhichEdge)
  
End Function

Public Function AlphaBlend2(ByVal Image2 As cRVTVBIMG, ByVal RGBAlphaMask As Long, _
                            Optional ByVal AutoTile As Boolean = False) As Boolean
Attribute AlphaBlend2.VB_Description = "AlphaBlend two Images together using a passed in Mask"
  
  AlphaBlend2 = Combine2Images(zWidth, zHeight, zIntBPP, zPixBits(), _
                               Image2, RGBAlphaMask, AutoTile, AlphaBlending:=True)
  
End Function

Public Function Combine2(ByVal Image2 As cRVTVBIMG, ByVal OpCode As PIC_COMBINATION_OPCODES, _
                         Optional ByVal AutoTile As Boolean = False) As Boolean
Attribute Combine2.VB_Description = "Combine 2 Images according to Passed in Opcode"
  
  Combine2 = Combine2Images(zWidth, zHeight, zIntBPP, zPixBits(), _
                            Image2, OpCode, AutoTile, AlphaBlending:=False)
  
End Function

'===================== THREE Image Procs ======================================================================

Public Function AlphaBlend3(ByVal Image2 As cRVTVBIMG, ByVal RGBAlphaImage As cRVTVBIMG) As Boolean
Attribute AlphaBlend3.VB_Description = "Blend 3 Images where Image1 is blended with Image2 according to Image3 as a mask"
  
  AlphaBlend3 = AlphaBlend3Images(zWidth, zHeight, zIntBPP, zPixBits(), Image2, RGBAlphaImage)
  
End Function


'================================== COLOR OPERATIONS ======================================================

Public Function UnMapColor() As Boolean
Attribute UnMapColor.VB_Description = "reduce colormap into a 24BPP image"

  UnMapColor = APIOperations(zPicState, zWidth, zHeight, zPixBits(), zIntBPP, zCMap(), zNCMapColors, PIC_UNMAP_COLOR)

End Function

'This one cannot be used in conjunction with Other Opcodes passing parameters or UnMapColor
Public Function MSMapColor(ByVal NewBPP As Long) As Boolean
Attribute MSMapColor.VB_Description = "Use MS mapping to determine ColorMap"

  MSMapColor = APIOperations(zPicState, zWidth, zHeight, zPixBits(), zIntBPP, zCMap(), zNCMapColors, PIC_MSMAP_COLOR, NewBPP)

End Function

Public Function InvertColor() As Boolean
Attribute InvertColor.VB_Description = "Replace all colors with their inverse"

  InvertColor = InvertImageColor(zWidth, zHeight, zIntBPP, zPixBits(), zCMap(), zNCMapColors)

End Function

Public Function MaskColor(ByVal RGBMask As Long) As Boolean
Attribute MaskColor.VB_Description = "Return only color components with mask values (AND)"

  MaskColor = CombineColor(PCC_AND, RGBMask)

End Function

'Replace TargetColor in region with ReplacementColor
Public Function ReplaceColor(ByVal TargetRGBColor As Long, ByVal SearchRadius As Long, ReplacementRGBColor As Long) As Boolean
Attribute ReplaceColor.VB_Description = "Replace a given color (or those within a range of it) with a new color"
  
  ReplaceColor = ReplaceImageColor(zWidth, zHeight, zIntBPP, zPixBits(), zCMap(), zNCMapColors, _
                                   TargetRGBColor, SearchRadius, ReplacementRGBColor)

End Function

Public Function SeparateColor(ByVal Which As PIC_SEPARATION_OPCODES) As Boolean
Attribute SeparateColor.VB_Description = "Make a color separation"

  SeparateColor = CombineColor(Which, &HFFFFFF)

End Function

Public Function EqualizeColor() As Boolean
Attribute EqualizeColor.VB_Description = "Attempt to color balance the image"

  EqualizeColor = EqualizeImageColor(zWidth, zHeight, zIntBPP, zPixBits())

End Function

Public Function ThresholdColor(ByVal RGBMask As Long) As Boolean
Attribute ThresholdColor.VB_Description = "Return Threshold if color>Threshold else color component"

  ThresholdColor = CombineColor(PCC_MIN, RGBMask)

End Function

'sRGB Gamma color correction, Abs(Gammas) must be >=1 <=2.4   - makes darker, + makes lighter
Public Function GammaCorrectColor(Optional ByVal RGamma As Single = 2.4, _
                                  Optional ByVal GGamma As Single = 2.4, _
                                  Optional ByVal BGamma As Single = 2.4) As Boolean
Attribute GammaCorrectColor.VB_Description = "Apply Gamma correction to an Image. Inverse also possible"
 
  GammaCorrectColor = GammaImageColor(zWidth, zHeight, zIntBPP, zPixBits(), zCMap(), zNCMapColors, _
                                      RGamma, GGamma, BGamma)
End Function

'Brightness color correction, Factor must be -1<=Factor<=1   - less Brightness, + more Brightness
Public Function BrightnessCorrectColor(Optional ByVal RFactor As Single = 0, _
                                     Optional ByVal GFactor As Single = 0, _
                                     Optional ByVal BFactor As Single = 0) As Boolean
Attribute BrightnessCorrectColor.VB_Description = "According to factors passed in adjust colour brightness, Factor affects the midpoint of a parabolic curve"
 
  BrightnessCorrectColor = BrightnessImageColor(zWidth, zHeight, zIntBPP, zPixBits(), zCMap(), zNCMapColors, _
                                                RFactor, GFactor, BFactor)
End Function

'Each component of AlphaMask contributes, for flat alpgha pass a grey value ie r=g=b
Public Function AlphaBlendColor(ByVal RGBColor As Long, ByVal RGBAlphaMask As Long) As Boolean

  AlphaBlendColor = CombineImageColor(zWidth, zHeight, zIntBPP, zPixBits(), zCMap(), zNCMapColors, _
                                      RGBAlphaMask, RGBColor, AlphaBlending:=1)

End Function

'Each component of AlphaMask contributes, for flat alpgha pass a grey value ie r=g=b
Public Function InverseAlphaBlendColor(ByVal RGBColor As Long, ByVal RGBAlphaMask As Long) As Boolean
Attribute InverseAlphaBlendColor.VB_Description = "Remove color from an image by Inverse Alpha blending"

  InverseAlphaBlendColor = CombineImageColor(zWidth, zHeight, zIntBPP, zPixBits(), zCMap(), zNCMapColors, _
                                             RGBAlphaMask, RGBColor, AlphaBlending:=-1)

End Function

'values are automatically clamped to 0..255
Public Function CombineColor(ByVal OpCode As PIC_COMBINATION_OPCODES, ByVal RGBColor As Long) As Boolean
Attribute CombineColor.VB_Description = "Combine Image and a color in manner defined by Opcode"

  CombineColor = CombineImageColor(zWidth, zHeight, zIntBPP, zPixBits(), zCMap(), zNCMapColors, _
                                   OpCode, RGBColor, AlphaBlending:=0)

End Function

'===== COLORMAPPING IN PREPARATION FOR IMAGE SAVING ==========================================================

Public Function DoColorMapping() As Boolean
Attribute DoColorMapping.VB_Description = "Obvious"
  
  DoColorMapping = XEQRemapping()
  
  If Not DoColorMapping Then
    MsgBox "RVTVBImg:: Color Mapping has Failed:" & vbCrLf _
         & "Either Pipeline is Invalid" & vbCrLf _
         & "or there is a shortage of Memory", vbCritical
  End If

End Function

'================================================================================================================
'Actually commit to disk (this could (probably will) go awry if called at the wrong time)
Public Function CommitToDisk(FilePath As String, PicType As IMG_PICTYPE) As Boolean
Attribute CommitToDisk.VB_Description = "Write the current Image to disk according to Pipeline parameters"
  
  etSave = zTimer.Elapsed
  
  If FSet(zPicState, IS_VALID_PIPELINE) Then
    On Error Resume Next
    If Len(FilePath) <> 0 Then Call Kill(FilePath)
    On Error GoTo 0
    
    Select Case PicType
     Case PIC_BMP:
       CommitToDisk = SaveBMP(FilePath, zPicState, zWidth, zHeight, _
                                        zPixBits(), zIntBPP, zReqBPP, zCMap(), zNCMapColors)
       
     Case PIC_GIF, PIC_GIF_LACED:
       If FSet(zPicState, IS_CMAPPED) Then
         CommitToDisk = SaveGIF(FilePath, zWidth, zHeight, _
                                          zReqBPP, zPixBits(), zIntBPP, zCMap(), _
                                          (PicType = PIC_GIF_LACED), zTransparentC)
       End If
       
     Case PIC_PNM:
       CommitToDisk = SavePNM(FilePath, zWidth, zHeight, _
                                        zPixBits(), zReqBPP, zCMap(), zNCMapColors, zCMAPMode)
    End Select
  End If
  
  etSave = zTimer.Elapsed - etSave
  
  If Not CommitToDisk Then
    MsgBox "RVTVBImg:: An Error has occurred while Saving Image" & vbCrLf _
          & "Is Filepath< " & FilePath & " > valid?", vbCritical
  End If

End Function

'To see all of the valid modes available call this sequentially from right to left
'If an option is given as -1 then the valid values for the combinations up to here is returned
'as a bitflagset. If returned as zero, the combination so far is actually invalid
Public Function ValidOptionFlags(Optional ByVal PicType As IMG_PICTYPE = -1, _
                                 Optional ByVal ColorMode As IMG_COLORMODES = -1, _
                                 Optional ByVal NColorMode As IMG_DEPTHMODES = -1, _
                                 Optional ByVal CMAPMode As IMG_CMAPMODES = -1, _
                                 Optional ByVal DitherMode As IMG_DITHERMODES = -1) As Long
Attribute ValidOptionFlags.VB_Description = "For a given combination of Options return a flagset show which are valid"
End Function

'==============================================================================================================
'=== PRIVATE BELOW HERE =======================================================================================
'==============================================================================================================

'==============================================================================================================
'================================ PIPELINE VALIDATION AND EXECUTION ===========================================
'==============================================================================================================
Private Function QualifyPipeline(ByVal PicType As IMG_PICTYPE, _
                                 ByVal ColorMode As IMG_COLORMODES, _
                                 ByVal BitsPerPixel As IMG_DEPTHMODES, _
                                 ByVal CMAPMode As IMG_CMAPMODES, _
                                 ByVal DitherMode As IMG_DITHERMODES) As Boolean

  Call ClrF(zPicState, IS_VALID_PIPELINE)
  Call ValidatePipeline(PicType, ColorMode, BitsPerPixel, CMAPMode, DitherMode) 'fully extends to a valid pipeline
  
  zPicType = PicType
  zColorMode = ColorMode
  zCMAPMode = CMAPMode
  zDitherMode = DitherMode
 
  zReqBPP = BitsPerPixel          'required representation
    
'  If FSet(zCMAPMode, PIC_MSMAP) Then
'    zIntBPP = RestrictToMSModes(zCMAPMode, zReqBPP)  'we are limited to 1,4,8,16 24 or 32 planes of quantizing/mapping
'  Else
    If zReqBPP = PIC_1BPP Then Call SetF(zCMAPMode, PIC_GMAP)     'force greymap for BW
'    zIntBPP = PIC_24BPP                                           'internal representation (assumed)
'  End If

  'figure out the Nr of Colors
  If zReqBPP <= PIC_8BPP Then zNCMapColors = 2 ^ zReqBPP Else zNCMapColors = 0
  
  Call SetF(zPicState, IS_VALID_PIPELINE)
  QualifyPipeline = True
  
End Function

'===== COLORMAPPING IN PREPARATION FOR IMAGE SAVING ==========================================================

Private Function XEQRemapping() As Boolean
  
  If FClr(zPicState, IS_VALID_PIPELINE) Then GoTo ErrorFound       'must have a validated pipeline
   
  etRemap = zTimer.Elapsed
  
  If FSet(zCMAPMode, PIC_MSMAP) Then     'can be true for BMP image types for 1,4,8,15,32 BPP
    'There are cases where Operations may have unmapped the Image and it is now 24BPP
    'In this case we use the API to do MS style of remapping  'this is in mBitmap.bas
    
    etQuant = zTimer.Elapsed
    If zReqBPP = PIC_15BPP Then             'Special case where MS format is actually 1 bit wider than used
      If zIntBPP <> PIC_16BPP Then
        If Not MSMapColor(PIC_16BPP) Then GoTo ErrorFound
      End If
    ElseIf zReqBPP <> zIntBPP Then
      If Not MSMapColor(zReqBPP) Then GoTo ErrorFound
    End If
    etQuant = zTimer.Elapsed - etQuant
    
   Else                               '1-8,15,16,24,32 from a MS Bitmap
    etQuant = zTimer.Elapsed
      
    If FClr(zCMAPMode, PIC_UNMAPPED) Then   'Not going for 24 or 32 bit output
        
      If zIntBPP > PIC_8BPP Then      'we havent got a colormap so need to make one or use a fixed one
  
        '-------------------- COLORMAP DETERMINATION ---------------------------------------------------
        
        If FSet(zCMAPMode, PIC_VMAP) Then    'it will get picked up in AnalyseGamut
          '
        ElseIf FSet(zCMAPMode, PIC_FIXED_CMAP) Then                     'we will use a fixed map
        
          If FSet(zCMAPMode, PIC_MODIFIED_CMAP) Then                    'we will modify a fixed color map
            zCMAPMode = PIC_FIXED_CMAP Or zReqBPP
            Call GenFixedMap(zCMAPMode, zCMap(), zNCMapColors, zReqBPP)
            zCMAPMode = PIC_OPTIMAL_CMAP Or zReqBPP
            Call OptimiseColors(zWidth, zHeight, zPixBits(), zIntBPP, zCMap(), zNCMapColors, zCMAPMode)
            
          ElseIf FSet(zCMAPMode, PIC_FIXED_CMAP_USER) Then              'we will use the user's map
            Call GenUserCMap(zCMap(), zNCMapColors, zUserCMAP())
          Else
            Call GenFixedMap(zCMAPMode, zCMap(), zNCMapColors, zReqBPP) 'we will use a fixed color/grey map
          End If
          
        Else   '------------------------------------- we are going to make one ourselves (hooray) -------
          
          'we are going to remap the BMP onto a best fit palette. We will use the 5 most significant bits
          'so we are effectively going to select from a gamut of 32x32x32 or 32768 shades
          'if you've got the memory you could make this 6 deep or 64^3=262144 shades, you'll also have to up HistCmap
          'NOTE:::: zIntBPP is the format given, zReqBPP=BitsPerPixel is what we want to achieve
          
          If zReqBPP > PIC_8BPP Then
            zNCMapColors = 256
          Else
            zNCMapColors = 2 ^ zReqBPP
          End If
          
          ReDim zCMap(0 To zNCMapColors - 1)
  
          'Initialize Remapping variables  - Using OctTree Quantization    'PixelWidth=significant bits BPP
          'also the colormap will be resized to the minimum power of 2 greater or equal to the number of colors used
          zReqBPP = CreateOctTreeCMap(zPicType, zWidth, zHeight, zPixBits(), zIntBPP, _
                                      zCMap(), zNCMapColors, OctTreeDepthBits:=6&)
  
          'Cleanup after ourselves to minimise storage
          Call DeleteOctTreeCMap
        End If
        etQuant = zTimer.Elapsed - etQuant
  
        '-------------------------- PIXEL REMAPPING  -----------------------------------------------------
        etRemap = zTimer.Elapsed
        If zReqBPP <> 0 Then
            
          ' Remap the new Palette onto the Image by dithering - PixBits is resized where necessary
          Select Case zDitherMode
           
           Case PIC_DITHER_SED1, PIC_DITHER_SED2, PIC_DITHER_SED3:
             Call SEDDitherMapColors(zWidth, zHeight, zPixBits(), zIntBPP, _
                                     zCMap(), zNCMapColors, zCMAPMode, zDitherMode)
           Case PIC_DITHER_NONE:
             Call SimpleMapColors(zWidth, zHeight, zPixBits(), zIntBPP, zCMap(), zNCMapColors, zCMAPMode)
    
           Case Else:
             Call MatDitherMapColors(zWidth, zHeight, zPixBits(), zIntBPP, _
                                     zCMap(), zNCMapColors, zCMAPMode, zDitherMode)
          End Select
        End If
        If zReqBPP = 0 Then GoTo ErrorFound
        
        'after colormapping has happened we are usually now an 8 bit pixelmap resized properly
        If FSet(zCMAPMode, PIC_VMAP) Then
          zIntBPP = PIC_16BPP
        Else
          zIntBPP = PIC_8BPP         ' we have made this into a newly colormapped 8BPP Bitmap
        End If
      End If
      
      'we now have a colormap in most cases, and the internal format is now PixelWidth 1,4 or 8 bits
      If NCMapColors <> 0 Then Call SetF(zPicState, IS_CMAPPED)
    End If
  End If

  '----------------------------- GREY REMAPPING  -------------------------------------------------------
  'we may have asked for a grey mapping - this is a cheap way to do it (but only if there is a colormap)

  If zColorMode = PIC_GREY And FClr(zCMAPMode, PIC_GMAP) Then     'but dont bother if we mapped to grey
    If FSet(zPicState, IS_CMAPPED) Then
      Call CMaptoGrey(zCMap())                                    'change the RGB palette to YUV greys
     Else
      Call MakePixelsGrey(zWidth, zHeight, zPixBits(), zIntBPP)   'Grey requires individual pixels to be changed
    End If
  End If

  '-------------------------- PALETTE SHRINKING (GIFS) -------------------------------------------------
  ' attempt to reduce the number of colors needed. This is done by counting their uses
  ' then shifting used ones to the front of the CMAP palette while at the same time fixing their
  ' references in the PixBits array

  If zPicType = PIC_GIF Or zPicType = PIC_GIF_LACED Then
    If FSet(zPicState, IS_CMAPPED) Then
      If FSet(zCMAPMode, PIC_SMALLEST_CMAP) Then
        If zIntBPP > PIC_1BPP Then
          zReqBPP = ShrinkCMap(zPixBits(), zIntBPP, zCMap(), zReqBPP)  'no point if only 2 colours
        End If
      End If
    End If
  End If

  XEQRemapping = True

ErrorFound:
  etRemap = zTimer.Elapsed - etRemap

End Function
'===========================================================================================================

':) Ulli's VB Code Formatter V2.16.6 (2003-Jun-03 21:35) 166 + 543 = 709 Lines
