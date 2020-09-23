Attribute VB_Name = "mValidate"
Option Explicit

'mValidate.bas - A General Purpose Validation Module for Valid RVTVBIMG pipeline combinations
'This can be called through Valid... options in main class for various purposes

'- ©2003 Ron van Tilburg - All rights reserved  15.07.2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please


'These CMAP Modes are fully builtup versions of the externally available parts

Public Enum IMG_CMAPMODES_INTERNAL
     PIC_SMALLEST_CMAP = &H80&    'Use the smallest palette possible (of the colors given by MS)
             PIC_MSMAP = &H100&   'Using a MS CMAP
              PIC_GMAP = &H200&   'Using a Grey CMAP
              PIC_VMAP = &H400&   'Using a Virtual CMAP
          PIC_UNMAPPED = &H800&   'Do not use a Map at all
         
     PIC_FIXED_CMAP_C4 = &H2002&  'Use a 4 tone color KCMY
     PIC_FIXED_CMAP_C8 = &H2003&  'Use a   8 tone color KRGBWCMY (8=2*2*2)
    PIC_FIXED_CMAP_C16 = &H2004&  'Use a  16 tone color map (16=2*4*2)
    PIC_FIXED_CMAP_C32 = &H2005&  'Use a  32 tone color map (27=3*3*3)
    PIC_FIXED_CMAP_C64 = &H2006&  'Use a  64 tone color map (64=4*4*4)
   PIC_FIXED_CMAP_C128 = &H2007&  'Use a 128 tone dithermap (125)
   PIC_FIXED_CMAP_C256 = &H2008&  'Use a 256 colormap (252 = 6,7,6)
   
     PIC_FIXED_CMAP_BW = &H2201&  'Use a   2 tone grey = Black and White
     PIC_FIXED_CMAP_G4 = &H2202&  'Use a   4 tone grey map
     PIC_FIXED_CMAP_G8 = &H2203&  'Use a   8 tone grey map
    PIC_FIXED_CMAP_G16 = &H2204&  'Use a  16 tone grey map
    PIC_FIXED_CMAP_G32 = &H2205&  'Use a  32 tone grey map
    PIC_FIXED_CMAP_G64 = &H2206&  'Use a  64 tone grey map
   PIC_FIXED_CMAP_G128 = &H2207&  'Use a 128 tone grey map
   PIC_FIXED_CMAP_G256 = &H2208&  'Use a 256 tone grey map
   
   PIC_FIXED_VMAP_C512 = &H2409&  'Use a  8* 8* 8    (Virtual) Fixed Regular Color map  9 bit 3,3,3 16bit save
    PIC_FIXED_VMAP_C4K = &H240C&  'Use a 16*16*16    (Virtual) Fixed Regular Color map 12 bit 4,4,4 16bit save
   PIC_FIXED_VMAP_C32K = &H240F&  'Use a 32*32*32    (Virtual) Fixed Regular Color map 15 bit 5,5,5 16bit save
   PIC_FIXED_VMAP_C64K = &H2410&  'Use a 32*64*32    (Virtual) Fixed Regular Color map 16 bit 5,6,5 16bit save
     
     PIC_UNMAPPED_C16M = &H2818&  '24 bit 8,8,8
   PIC_UNMAPPED_C16M32 = &H2820&  '32 bit 8,8,8,8a
End Enum

Private Const VALID_PIC_TYPES As Long = &HF&                   '          1111
Private Const VALID_COLOR_OPTIONS As Long = &H7&               '           111
                                                               '  24b  9b   4b   1b
Private Const VALID_DEPTH_OPTIONS As Long = &H3FFF&            '  11,1111,1111,1111
Private Const VALID_BMP_DEPTH_OPTIONS As Long = &H3FFF&        '  11,1111,1111,1111
Private Const VALID_GIF_DEPTH_OPTIONS As Long = &HFF&          '  00,0000,1111,1111
Private Const VALID_PNM_DEPTH_OPTIONS As Long = &H10FF&        '  01,0000,1111,1111

Private Const VALID_BW_DEPTH_OPTIONS As Long = &H1&            '  00,0000,0000,0001
Private Const VALID_GREY_DEPTH_OPTIONS As Long = &HFF&         '  00,0000,1111,1111

Private Const VALID_CMAP_OPTIONS As Long = &H1FF&              '   1,1111,1111
Private Const VALID_BMP_CMAP_OPTIONS As Long = &H1FF&          '   1,1111,1111
Private Const VALID_GIF_CMAP_OPTIONS As Long = &H1FE&          '   1,1111,1110
Private Const VALID_PNM_CMAP_OPTIONS As Long = &H1FE&          '   1,1111,1110

Private Const VALID_BW_CMAP_OPTIONS As Long = &H145&           '   1,0100,0101
Private Const VALID_WW_CMAP_OPTIONS As Long = &H1C4&           '   1,1100,0100   '4,8
Private Const VALID_XX_CMAP_OPTIONS As Long = &H1CF&           '   1,1100,1111   '16
Private Const VALID_YY_CMAP_OPTIONS As Long = &H1C6&           '   1,1100,0110   '32,64,128
Private Const VALID_ZZ_CMAP_OPTIONS As Long = &H1F7&           '   1,1111,0111   '256
Private Const VALID_32K_CMAP_OPTIONS As Long = &H5&            '   0,0000,0101   '32K
Private Const VALID_64K_CMAP_OPTIONS As Long = &H4&            '   0,0000,0100   '512,4K,64K
Private Const VALID_16M_CMAP_OPTIONS As Long = &H1&            '   0,0000,0001   '16M

Private Const VALID_MSMAP_DITHER_OPTIONS As Long = &H1&        '0000,0000,0001
Private Const VALID_DITHER_OPTIONS As Long = &HFFF&            '1111,1111,1111

Private zPicTypes()     As Variant
Private zColorModes()   As Variant
Private zDepthModes()   As Variant
Private zCMAPModes()    As Variant
Private zDitherModes()  As Variant

Public Sub InitValidationArrays()

  zPicTypes() = Array(PIC_BMP, PIC_GIF, PIC_GIF_LACED, PIC_PNM)
  
  zColorModes() = Array(PIC_COLOR, PIC_BW, PIC_GREY)
  
  zDepthModes() = Array(PIC_1BPP, PIC_2BPP, PIC_3BPP, PIC_4BPP, PIC_5BPP, PIC_6BPP, PIC_7BPP, PIC_8BPP, _
                        PIC_9BPP, PIC_12BPP, PIC_15BPP, PIC_16BPP, PIC_24BPP, PIC_32BPP)
  
  zCMAPModes() = Array(PIC_MS_CMAP, PIC_OPTIMAL_CMAP, PIC_FIXED_CMAP, PIC_FIXED_CMAP_VGA, PIC_FIXED_CMAP_INET, _
                       PIC_FIXED_CMAP_MS256, PIC_FIXED_CMAP_GREY, PIC_MODIFIED_CMAP, PIC_FIXED_CMAP_USER)
  
  zDitherModes() = Array(PIC_DITHER_NONE, PIC_DITHER_BIN, PIC_DITHER_ORD, PIC_DITHER_HTC, _
                         PIC_DITHER_FDIAG, PIC_DITHER_BDIAG, PIC_DITHER_HORZ, PIC_DITHER_VERT, _
                         PIC_DITHER_SED1, PIC_DITHER_SED2, PIC_DITHER_SED3, PIC_DITHER_BNM)
End Sub

'If the ordinals from the Enums are passed in the Values of the Ordinals Are Returned
Public Sub PicEnumFromOrdinals(ByRef PicType As Long, _
                               ByRef ColorMode As Long, _
                               ByRef RequiredBPP As Long, _
                               ByRef CMAPMode As Long, _
                               ByRef DitherMode As Long)

  If LBound(zPicTypes) <= PicType And PicType <= UBound(zPicTypes) Then
    PicType = zPicTypes(PicType)
  Else
    PicType = PIC_BMP
  End If
  
  If LBound(zColorModes) <= ColorMode And ColorMode <= UBound(zColorModes) Then
    ColorMode = zColorModes(ColorMode)
  Else
    ColorMode = PIC_COLOR
  End If
  
  If LBound(zDepthModes) <= RequiredBPP And RequiredBPP <= UBound(zDepthModes) Then
    RequiredBPP = zDepthModes(RequiredBPP)
  Else
    RequiredBPP = PIC_8BPP
  End If
  
  If LBound(zCMAPModes) < CMAPMode And CMAPMode < UBound(zCMAPModes) Then
    CMAPMode = zCMAPModes(CMAPMode)
  Else
    CMAPMode = PIC_OPTIMAL_CMAP
  End If
  
  If LBound(zDitherModes) < DitherMode And DitherMode <= UBound(zDitherModes) Then
    DitherMode = zDitherModes(DitherMode)
  Else
    DitherMode = PIC_DITHER_NONE
  End If
  
End Sub

'From a passed In Pipeline Parameter Set build up the fully qualified pipeline
Public Sub ValidatePipeline(ByRef PicType As Long, _
                            ByRef ColorMode As Long, _
                            ByRef RequiredBPP As Long, _
                            ByRef CMAPMode As Long, _
                            ByRef DitherMode As Long)
  
  Dim i As Long, j As Long
  
  PicType = ValidatePicType(PicType)
  ColorMode = ValidateColorMode(PicType, ColorMode)
  RequiredBPP = ValidateDepthMode(PicType, ColorMode, RequiredBPP)
  CMAPMode = ValidateCMAPMode(PicType, ColorMode, RequiredBPP, CMAPMode)
  DitherMode = ValidateDitherMode(PicType, ColorMode, RequiredBPP, CMAPMode, DitherMode)
End Sub


Private Function ValidatePicType(ByVal PicType As Long) As Long
  Dim i As Long, j As Long
  
  j = 1
  For i = 0 To UBound(zPicTypes)
    If (j And VALID_PIC_TYPES) <> 0 Then
      If PicType = zPicTypes(i) Then
        ValidatePicType = zPicTypes(i)
        Exit Function
      End If
    End If
    j = j + j
  Next i
  
  ValidatePicType = PIC_BMP
End Function

'ColorMode is independent of Pictype
Private Function ValidateColorMode(ByVal PicType As Long, _
                                   ByVal ColorMode As Long) As Long
  Dim i As Long, j As Long
  
  j = 1
  For i = 0 To UBound(zColorModes)
    If (j And VALID_COLOR_OPTIONS) <> 0 Then
      If ColorMode = zColorModes(i) Then
        ValidateColorMode = zColorModes(i)
        Exit Function
      End If
    End If
    j = j + j
  Next i
  
  ValidateColorMode = PIC_COLOR

End Function

'DepthMode depends on PicType and ColorMode
Private Function ValidateDepthMode(ByVal PicType As Long, _
                                   ByVal ColorMode As Long, _
                                   ByVal DepthBPP As Long) As Long
  
  Dim i As Long, j As Long, ValidValues As Long, Default As Long
  
  If ColorMode = PIC_BW Then
    ValidateDepthMode = PIC_1BPP
    Exit Function
  Else
    Select Case PicType
      Case PIC_BMP:                 ValidValues = VALID_BMP_DEPTH_OPTIONS
      Case PIC_GIF, PIC_GIF_LACED:  ValidValues = VALID_GIF_DEPTH_OPTIONS
      Case PIC_PNM:                 ValidValues = VALID_PNM_DEPTH_OPTIONS
    End Select
    If ColorMode = PIC_GREY Then ValidValues = ValidValues And VALID_GREY_DEPTH_OPTIONS
    Default = PIC_8BPP
    
    j = 1
    For i = 0 To UBound(zDepthModes)
      If (j And ValidValues) <> 0 Then
        If DepthBPP = zDepthModes(i) Then
          ValidateDepthMode = zDepthModes(i)
          Exit Function
        End If
      End If
      j = j + j
    Next i
  End If
  ValidateDepthMode = Default

End Function
    
'  If IntBPP = PIC_1BPP And FClr(CMAPMode, PIC_MS_CMAP) Then CMAPMode = PIC_FIXED_CMAP_GREY
'  If IntBPP <= PIC_3BPP Then Call SetF(CMAPMode, PIC_FIXED_CMAP)

'CMAPMode depends on PicType,ColorMode and DepthBPP
Private Function ValidateCMAPMode(ByVal PicType As Long, _
                                  ByVal ColorMode As Long, _
                                  ByVal DepthBPP As Long, _
                                  ByVal CMAPMode As Long) As Long
  
  Dim i As Long, j As Long, ValidValues As Long, Default As Long
  
  'check correspondences
  Select Case DepthBPP
     Case PIC_1BPP:
       ValidValues = VALID_BW_CMAP_OPTIONS:      Default = PIC_FIXED_CMAP_GREY  '2 Fixed Grey

     Case PIC_2BPP, PIC_3BPP:
       ValidValues = VALID_WW_CMAP_OPTIONS:      Default = PIC_FIXED_CMAP  '4,8

     Case PIC_4BPP:
       ValidValues = VALID_XX_CMAP_OPTIONS:      Default = PIC_FIXED_CMAP  '16

     Case PIC_5BPP, PIC_6BPP, PIC_7BPP:
       ValidValues = VALID_YY_CMAP_OPTIONS:      Default = PIC_FIXED_CMAP  '32,64,128

     Case PIC_8BPP:
       ValidValues = VALID_ZZ_CMAP_OPTIONS:      Default = PIC_OPTIMAL_CMAP  '256 'New

     Case PIC_15BPP:
       ValidValues = VALID_32K_CMAP_OPTIONS:     Default = PIC_FIXED_CMAP  '32768

     Case PIC_9BPP, PIC_12BPP, PIC_16BPP:
       ValidValues = VALID_64K_CMAP_OPTIONS:     Default = PIC_FIXED_CMAP  '512,4096,65536

     Case PIC_24BPP, PIC_32BPP:
       ValidValues = 0:                          Default = PIC_UNMAPPED  '16M
  End Select

  Select Case PicType
    Case 0:
      ValidValues = VALID_BMP_CMAP_OPTIONS And ValidValues
    Case 1, 2:
      ValidValues = VALID_GIF_CMAP_OPTIONS And ValidValues
    Case 3:
      ValidValues = VALID_PNM_CMAP_OPTIONS And ValidValues
  End Select

  j = 1
  For i = 0 To UBound(zCMAPModes)
    If (j And ValidValues) <> 0 Then
      If CMAPMode = zCMAPModes(i) Then
        CMAPMode = zCMAPModes(i)
        Exit For
      End If
    End If
    j = j + j
  Next i

  If i > UBound(zCMAPModes) Then CMAPMode = Default
  
  'Having got this far fix the BPP into relevant codes
  If CMAPMode <> PIC_FIXED_CMAP_VGA _
  And CMAPMode <> PIC_FIXED_CMAP_INET _
  And CMAPMode <> PIC_FIXED_CMAP_MS256 Then
    CMAPMode = CMAPMode Or DepthBPP
    If DepthBPP >= PIC_9BPP And DepthBPP <= PIC_16BPP Then
      Call SetF(CMAPMode, PIC_VMAP)
    End If
  End If
  
  ValidateCMAPMode = CMAPMode 'at last

End Function

'DitherMode depends only on CMAPMode
Private Function ValidateDitherMode(ByVal PicType As Long, _
                                    ByVal ColorMode As Long, _
                                    ByVal DepthBPP As Long, _
                                    ByVal CMAPMode As Long, _
                                    ByVal DitherMode As Long) As Long
  
  Dim i As Long, j As Long
  
  If FSet(CMAPMode, PIC_MSMAP) Then  'cannot dither
    ValidateDitherMode = PIC_DITHER_NONE
  Else
    j = 1
    For i = 0 To UBound(zDitherModes)
      If (j And VALID_DITHER_OPTIONS) <> 0 Then
        If DitherMode = zDitherModes(i) Then
          ValidateDitherMode = zDitherModes(i)
          Exit Function
        End If
      End If
      j = j + j
    Next i
    ValidateDitherMode = PIC_DITHER_NONE
  End If

End Function

'==============================================================================================================
'If MS Mapping is to be used only certain BitDepths are supported
Public Function RestrictToMSModes(ByRef CMAPMode As IMG_CMAPMODES, _
                                   ByVal ReqBPP As IMG_DEPTHMODES) As IMG_DEPTHMODES
  
  Dim IntBPP As Long
  
  If ReqBPP <= PIC_1BPP Then
    IntBPP = PIC_1BPP
  
  ElseIf ReqBPP <= PIC_4BPP Then
    IntBPP = PIC_4BPP
    Call SetF(CMAPMode, PIC_SMALLEST_CMAP)
  
  ElseIf ReqBPP <= PIC_8BPP Then
    IntBPP = PIC_8BPP
    Call SetF(CMAPMode, PIC_SMALLEST_CMAP)
  
  ElseIf ReqBPP = PIC_15BPP Then
    IntBPP = PIC_16BPP
  
  ElseIf ReqBPP = PIC_32BPP Then
    IntBPP = PIC_32BPP
    
  ElseIf ReqBPP > PIC_8BPP Then
    IntBPP = PIC_24BPP                           ' and will need quantizing,mapping and dithering of colours!!
    Call ClrF(CMAPMode, PIC_MSMAP)
  End If
  
  RestrictToMSModes = IntBPP

End Function
