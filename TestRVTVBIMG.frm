VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TestRVTVBIMG - Image Processing Pipeline Example - Display is Clipped for Oversize Images"
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBrightnessDown 
      Caption         =   "Darken"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   4600
      TabIndex        =   89
      ToolTipText     =   "Decrease Contrast Factor=-0.2"
      Top             =   9050
      Width           =   1100
   End
   Begin VB.CheckBox chkBrightnessUp 
      Caption         =   "Brighten"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   3500
      TabIndex        =   88
      ToolTipText     =   "Increase Contrast Factor=+0.2"
      Top             =   9050
      Width           =   1120
   End
   Begin VB.CheckBox chkInvBlend 
      Caption         =   "-ABlend"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   1300
      TabIndex        =   87
      ToolTipText     =   "eg. Inverse Blend Light Blue at 70%"
      Top             =   8700
      Width           =   1100
   End
   Begin VB.CheckBox chkTransparent 
      Caption         =   "Transp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   2400
      TabIndex        =   86
      ToolTipText     =   "make Black Transparent"
      Top             =   9050
      Width           =   1100
   End
   Begin VB.CheckBox chkCombine 
      Caption         =   "AddMod"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   200
      TabIndex        =   85
      ToolTipText     =   "eg. AddSmooth &H80C040"
      Top             =   9050
      Width           =   1100
   End
   Begin VB.CheckBox chkBlend3 
      Caption         =   "ABlend3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   1300
      TabIndex        =   84
      ToolTipText     =   "Blend Test Image, ColorTest by Grey Image"
      Top             =   9050
      Width           =   1100
   End
   Begin VB.CheckBox chkExtract 
      Caption         =   "Extract"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   4600
      TabIndex        =   83
      ToolTipText     =   "Extract a sub Image"
      Top             =   8700
      Width           =   1100
   End
   Begin VB.CheckBox chkFrame 
      Caption         =   "Frame"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   4600
      TabIndex        =   82
      ToolTipText     =   "Frame 10 pixels of all sides of image in Magenta"
      Top             =   8350
      Width           =   1100
   End
   Begin VB.CheckBox chkThresh 
      Caption         =   "Thresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   1300
      TabIndex        =   81
      ToolTipText     =   "Threshold above &HC0C0C0"
      Top             =   8000
      Width           =   1100
   End
   Begin VB.CheckBox chkReplace 
      Caption         =   "Replace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   2400
      TabIndex        =   80
      ToolTipText     =   "eg. Replace MidTones with Red"
      Top             =   8700
      Width           =   1100
   End
   Begin VB.CheckBox chkUnGamma 
      Caption         =   "-Gamma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   200
      TabIndex        =   79
      ToolTipText     =   "UnGamma Correct gRGB=1.8"
      Top             =   8700
      Width           =   1100
   End
   Begin VB.CheckBox chkGamma 
      Caption         =   "+Gamma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   200
      TabIndex        =   78
      ToolTipText     =   "Gamma Correct gRGB=1.8"
      Top             =   8350
      Width           =   1100
   End
   Begin VB.CheckBox chkDeJPG 
      Caption         =   "dJPG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   3500
      TabIndex        =   77
      ToolTipText     =   "Fix JPEG Quantize Errors"
      Top             =   8350
      Width           =   1100
   End
   Begin VB.CheckBox chkTrim 
      Caption         =   "Trim"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   4600
      TabIndex        =   76
      ToolTipText     =   "Trim 10 pixels of all sides of image"
      Top             =   8000
      Width           =   1100
   End
   Begin VB.CheckBox chkEqualize 
      Caption         =   "Equal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   1300
      TabIndex        =   75
      ToolTipText     =   "Equalize Image Balance"
      Top             =   7650
      Width           =   1100
   End
   Begin VB.CheckBox chkEmboss 
      Caption         =   "Emboss"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   3500
      TabIndex        =   74
      ToolTipText     =   "Emboss the Image"
      Top             =   8700
      Width           =   1100
   End
   Begin VB.CheckBox chkEdge 
      Caption         =   "Edge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   3500
      TabIndex        =   73
      ToolTipText     =   "Show Edges of Image"
      Top             =   8000
      Width           =   1100
   End
   Begin VB.CheckBox chkBlend 
      Caption         =   "+ABlend"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   1300
      TabIndex        =   72
      ToolTipText     =   "eg. Blend Light Blue at 70%"
      Top             =   8350
      Width           =   1100
   End
   Begin VB.CheckBox chkMaskY 
      Caption         =   "Mask"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   200
      TabIndex        =   71
      ToolTipText     =   "eg. Show Yellow Shades"
      Top             =   8000
      Width           =   1100
   End
   Begin VB.CheckBox chkBlur 
      Caption         =   "Blur"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   2400
      TabIndex        =   70
      ToolTipText     =   "Blur Image"
      Top             =   8350
      Width           =   1100
   End
   Begin VB.CheckBox chkSharpen 
      Caption         =   "Sharp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   2400
      TabIndex        =   69
      ToolTipText     =   "Sharpen Image"
      Top             =   8000
      Width           =   1100
   End
   Begin VB.CheckBox chkSize53 
      Caption         =   " 5/3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   3500
      TabIndex        =   68
      ToolTipText     =   "Enlarge to 5/3 size"
      Top             =   7650
      Width           =   1100
   End
   Begin VB.CheckBox chkRot180 
      Caption         =   "Rot180"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   4600
      TabIndex        =   67
      ToolTipText     =   "Rotate 180 deg"
      Top             =   7300
      Width           =   1100
   End
   Begin VB.CheckBox chkRotR 
      Caption         =   "Rot R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   3500
      TabIndex        =   66
      ToolTipText     =   "Rotate to Right 90 deg"
      Top             =   7300
      Width           =   1100
   End
   Begin VB.CheckBox chkRotL 
      Caption         =   "Rot L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   2400
      TabIndex        =   65
      ToolTipText     =   "Rotate to LEft 90 deg"
      Top             =   7300
      Width           =   1100
   End
   Begin VB.CommandButton cmdImageTest 
      Caption         =   "Image Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      ToolTipText     =   "A standard Test Image"
      Top             =   10980
      Width           =   1305
   End
   Begin VB.CommandButton cmdColorTest 
      Caption         =   "Color Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Color Ramps"
      Top             =   10560
      Width           =   1305
   End
   Begin VB.CommandButton cmdGreyTest 
      Caption         =   "Grey Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Grey Ramps"
      Top             =   10140
      Width           =   1305
   End
   Begin VB.CheckBox chkInvert 
      Caption         =   "Invert"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   200
      TabIndex        =   48
      ToolTipText     =   "Invert Colours ie Negative"
      Top             =   7650
      Width           =   1100
   End
   Begin VB.CheckBox chkSize23 
      Caption         =   " 2/3 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   2400
      TabIndex        =   47
      ToolTipText     =   "Reduce to 2/3 size"
      Top             =   7650
      Width           =   1100
   End
   Begin VB.CheckBox chkFlipH 
      Caption         =   "FlipH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   1300
      TabIndex        =   46
      ToolTipText     =   "Flip Image Horizontally"
      Top             =   7300
      Width           =   1100
   End
   Begin VB.CheckBox chkFlipV 
      Caption         =   "FlipV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   200
      TabIndex        =   45
      ToolTipText     =   "Flip Image vertically"
      Top             =   7300
      Width           =   1100
   End
   Begin VB.CheckBox chkClip 
      Caption         =   "Clip?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   350
      Left            =   4600
      TabIndex        =   49
      ToolTipText     =   "Clip to Size"
      Top             =   7650
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlgFileLoad 
      Left            =   2460
      Top             =   10800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Picture"
      Filter          =   "*gif|*jpg|*jpeg|*.bmp"
      InitDir         =   "AppDir"
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   14010
      TabIndex        =   44
      ToolTipText     =   "Complete Processing"
      Top             =   3900
      Width           =   1330
   End
   Begin VB.Frame frDither 
      Caption         =   "DitherMode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   3825
      Left            =   14010
      TabIndex        =   53
      Top             =   60
      Width           =   1330
      Begin VB.OptionButton OptDither 
         Caption         =   "Bwd Diag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   60
         TabIndex        =   37
         ToolTipText     =   "Emphasises \\\\"
         Top             =   1740
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Blue Noise"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   60
         TabIndex        =   43
         ToolTipText     =   "My own Blue Noise Mask - 32x32"
         Top             =   3480
         Width           =   1230
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "SED Equal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   60
         TabIndex        =   40
         ToolTipText     =   "Serpentine Error Diffusion - Ostromoukhov Variable Kernel"
         Top             =   3120
         Width           =   1230
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "SED RVT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   60
         TabIndex        =   42
         ToolTipText     =   "Serpentine Error Diffusion - My Kernel"
         Top             =   2880
         Width           =   1230
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "SED F-S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   60
         TabIndex        =   41
         ToolTipText     =   "Serpentine Error DIffusion - Floyd and Steinberg Kernel"
         Top             =   2640
         Width           =   1230
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Vertical"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   60
         TabIndex        =   39
         ToolTipText     =   "Emphasises |||||"
         Top             =   2280
         Width           =   1230
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Horizontal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   60
         TabIndex        =   38
         ToolTipText     =   "Emphasises ------"
         Top             =   2040
         Width           =   1230
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Fwd Diag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   60
         TabIndex        =   36
         ToolTipText     =   "Emphasises /////"
         Top             =   1500
         Width           =   1230
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Halftone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   60
         TabIndex        =   35
         ToolTipText     =   "A printer friendly dither"
         Top             =   1140
         Width           =   1230
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Ordered"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   34
         ToolTipText     =   "Classic Bayer Ordered Dither"
         Top             =   900
         Width           =   1230
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Binary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   33
         ToolTipText     =   "A regular Blue Noise Mask"
         Top             =   660
         Width           =   1230
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   32
         ToolTipText     =   "Simple Color Replacement"
         Top             =   300
         Width           =   1230
      End
   End
   Begin VB.Frame frCMAP 
      Caption         =   "CMAPMode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   4155
      Left            =   12900
      TabIndex        =   52
      Top             =   60
      Width           =   1090
      Begin VB.OptionButton OptCMAP 
         Caption         =   "User Map"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   8
         Left            =   75
         TabIndex        =   64
         ToolTipText     =   "Using a Colormap that the user has defined"
         Top             =   3540
         Width           =   960
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "Altered Fixed Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   7
         Left            =   75
         TabIndex        =   31
         ToolTipText     =   "Modifies A fixed color map for better color matching"
         Top             =   2940
         Width           =   960
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "Fixed Grey"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   75
         TabIndex        =   30
         ToolTipText     =   "N Shades of Grey"
         Top             =   2400
         Width           =   885
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "MS256"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   75
         TabIndex        =   29
         ToolTipText     =   "The Standard MS 256 color Palette"
         Top             =   1980
         Width           =   945
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "iNet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   75
         TabIndex        =   28
         ToolTipText     =   "An Internet Safe ColorMap"
         Top             =   1680
         Width           =   885
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "VGA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   75
         TabIndex        =   27
         ToolTipText     =   "The Standard VGA (ugh) ColorMap"
         Top             =   1380
         Width           =   885
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "Fixed Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   75
         TabIndex        =   26
         ToolTipText     =   "Uses a predefined ColorMap of NColors"
         Top             =   960
         Width           =   885
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   75
         TabIndex        =   25
         ToolTipText     =   "Creates a new ColorMap of NColors"
         Top             =   540
         Width           =   885
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "MS Map"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   75
         TabIndex        =   24
         ToolTipText     =   "Uses MS mapping methods (blitting)"
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame frNColors 
      Caption         =   "Colors/Bits"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   4155
      Left            =   11760
      TabIndex        =   51
      Top             =   60
      Width           =   1125
      Begin VB.OptionButton OptNColors 
         Caption         =   "16M 32"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   63
         Top             =   3720
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "16M 24"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   62
         Top             =   3445
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "64K 16"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   61
         Top             =   3160
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   " 4K  12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   22
         Top             =   2580
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "32K 15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   23
         Top             =   2860
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "512   9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   2315
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "256   8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   1980
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "128   7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   1740
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   " 64    6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   1500
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   " 32    5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1260
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   " 16    4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1020
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "  8     3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   780
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "  4     2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   540
         Width           =   930
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "  2     1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   930
      End
   End
   Begin VB.Frame frColorMode 
      Caption         =   "ColorMode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   4155
      Left            =   10650
      TabIndex        =   50
      Top             =   60
      Width           =   1095
      Begin VB.OptionButton optColorMode 
         Caption         =   "Grey"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   900
      End
      Begin VB.OptionButton optColorMode 
         Caption         =   "B + W"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1740
         Width           =   900
      End
      Begin VB.OptionButton optColorMode 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   780
         Width           =   900
      End
   End
   Begin VB.Frame frPicType 
      Caption         =   "PicType"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A3010&
      Height          =   4155
      Left            =   9690
      TabIndex        =   6
      Top             =   60
      Width           =   945
      Begin VB.OptionButton optPicType 
         Caption         =   "iGIF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   60
         ToolTipText     =   "Interlaced GIF"
         Top             =   1920
         Width           =   765
      End
      Begin VB.OptionButton optPicType 
         Caption         =   "PNM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "PBM, PGM, PPM Pixmap Files"
         Top             =   2700
         Width           =   705
      End
      Begin VB.OptionButton optPicType 
         Caption         =   "GIF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1620
         Width           =   645
      End
      Begin VB.OptionButton optPicType 
         Caption         =   "BMP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   675
      End
   End
   Begin VB.PictureBox picResult 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   5700
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   1
      ToolTipText     =   "Display of the Result"
      Top             =   4260
      Width           =   9660
   End
   Begin VB.CommandButton cmdLoadPic 
      Caption         =   "Load Pic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Load a BMP,GIF,JPG or PNM file"
      Top             =   9660
      Width           =   1305
   End
   Begin VB.PictureBox picOriginal 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   15
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      ToolTipText     =   "Display of the Original"
      Top             =   0
      Width           =   9660
   End
   Begin VB.Label lblLoadTime 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1980
      TabIndex        =   59
      Top             =   10020
      Width           =   1725
   End
   Begin VB.Label lblElapsed 
      Caption         =   "Elapsed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   58
      Top             =   10620
      Width           =   1845
   End
   Begin VB.Label lblSaveTime 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1980
      TabIndex        =   57
      Top             =   10320
      Width           =   1725
   End
   Begin VB.Label lblDitherTime 
      Caption         =   "Remap"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   56
      Top             =   10320
      Width           =   1845
   End
   Begin VB.Label lblCMAPTime 
      Caption         =   "Quantize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   55
      Top             =   10020
      Width           =   1845
   End
   Begin VB.Label Label3 
      Caption         =   "Timings (sec)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   54
      Top             =   9660
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'- ©2001/2003 Ron van Tilburg - All rights reserved  1.01.2001/Jun-Jul 2003
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'TESTRVTVBIMG.frm
'This program is just a tester for the DLL which is where the real work gets done -
'however it may suffice as a starting point for a Image processing application - serious submissions welcome

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

' Require References to RVTVBGDI , and Components to COMMDLG32.OCX
' Jul 2003    - Updated for all of the New Options Available
'             - New Grey and Color Tests  - no longer need separate Image Files

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

Private Const NPICTYPE_OPTIONS As Long = 4
Private Const NCOLOR_OPTIONS   As Long = 3
Private Const NNCOLOR_OPTIONS As Long = 14
Private Const NCMAP_OPTIONS As Long = 9
Private Const NDITHER_OPTIONS As Long = 12

Private Const TEMPFILE As String = "C:\TEMP\zzz.zzz"
Private Const TEMPFILE2 As String = "C:\TEMP\yyy.yyy"

Private CurPicType As Long
Private CurColorMode As Long
Private CurNColors As Long
Private CurCMAPMode As Long
Private CurDitherMode As Long

Private FileDir As String
Private FileName As String

Private ImgCB As cRVTVBIMG
Private UserCMap() As RGBA

Private Sub Form_Load()

  CurPicType = -1
  CurColorMode = -1
  CurNColors = -1
  CurCMAPMode = -1
  CurDitherMode = -1
  Call picOriginal.ZOrder(0)
  Call optPicType_Click(0)

  Set ImgCB = New cRVTVBIMG '("RVTVBIMG.cRVTVBIMG")
  '  Set ImgCB = CreateObject("RVTVBIMG.cRVTVBIMG")

End Sub

Private Sub cmdLoadPic_Click()

 Dim q() As String

  Call Form1.picOriginal.ZOrder(0)
  With dlgFileLoad
    .DialogTitle = "Select Picture File"
    .CancelError = False
    If FileDir = "" Then .InitDir = App.Path Else .InitDir = FileDir
    .FileName = FileName
    'ToDo: set the flags and attributes of the common dialog control
    .Filter = "GIF Files (*.gif)|*.gif|JPG Files (*.jp*)|*.jp*|BMP Files (*.bmp)|*.bmp|All Files (*.*)|*.*"
    .FilterIndex = 4
    .ShowOpen
    If Len(.FileName) <> 0 Then
      q = Split(.FileName, "\")
      FileName = q(UBound(q))
      q(UBound(q)) = ""
      FileDir = Join$(q, "\")
      If ImgCB.LoadFromFile(.FileName) Then
        If Not ImgCB.PutImageObjhDC(Form1.picOriginal, IRO_TILE) Then
          MsgBox "Error during Display (image too large (>16M)?)", vbInformation
        End If
       Else
        MsgBox "Error during Load (image invalid/unsupported type?)", vbInformation
      End If
    End If
  End With

End Sub

Private Sub cmdGreyTest_Click()

 Dim i As Long, j As Long, c As Long

  With Form1.picOriginal
    Call .ZOrder(0)
    Call GradientFillRectDC(.hDC, 0, 0, 480, 480, &H808080, &H808080, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 480, 0, 640, 480, vbWhite, vbBlack, GF_RECTVERT)
    .Refresh
  End With
  FileName = ""
End Sub

Private Sub cmdColorTest_Click()

  With Form1.picOriginal
    Call .ZOrder(0)
    Call GradientFillRectDC(.hDC, 0, 0, 128, 160, vbRed, vbGreen, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 128, 0, 256, 160, vbCyan, vbRed, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 256, 0, 384, 160, vbMagenta, vbRed, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 384, 0, 512, 160, vbYellow, vbRed, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 512, 0, 640, 160, vbCyan, vbMagenta, GF_TRI4WAYG)

    Call GradientFillRectDC(.hDC, 0, 160, 128, 320, vbGreen, vbBlue, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 128, 160, 256, 320, vbCyan, vbGreen, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 256, 160, 384, 320, vbMagenta, vbGreen, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 384, 160, 512, 320, vbYellow, vbGreen, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 512, 160, 640, 320, vbMagenta, vbYellow, GF_TRI4WAYG)

    Call GradientFillRectDC(.hDC, 0, 320, 128, 480, vbBlue, vbRed, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 128, 320, 256, 480, vbCyan, vbBlue, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 256, 320, 384, 480, vbMagenta, vbBlue, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 384, 320, 512, 480, vbYellow, vbBlue, GF_TRI4WAYG)
    Call GradientFillRectDC(.hDC, 512, 320, 640, 480, vbYellow, vbCyan, GF_TRI4WAYG)
    .Refresh
  End With
  FileName = ""
End Sub

Private Sub cmdImageTest_Click()

  Call Form1.picOriginal.ZOrder(0)
  FileName = App.Path & "\ImageTest.jpg"
  If ImgCB.LoadFromFile(FileName) Then
    Call ImgCB.DisplayImage(Form1.picOriginal)
  End If
End Sub

'This is the Guts Of it - we call the various Dll functions in turn
Private Sub cmdGo_Click()

 Dim PicType As Long, ColorMode As Long, RequiredBPP As Long, CMAPMode As Long, DitherMode As Long
 Dim rcOK As Boolean, Parm1 As Long, Parm2 As Long, ImgCB2 As cRVTVBIMG, ImgCB3 As cRVTVBIMG, Color As Long

  Call picOriginal.ZOrder(0)
  DoEvents
  Me.MousePointer = vbHourglass

  'Lets Work out all of the parameters of the main (and only) call

  Call GetPipelineParms(PicType, ColorMode, RequiredBPP, CMAPMode, DitherMode)
  
  'We dummy up a Colormap just to demonstrate the process
  If CMAPMode = PIC_FIXED_CMAP_USER Then
    Call GenUserCMap(2 ^ RequiredBPP)     'This should always match
    ImgCB.UserCMap() = UserCMap()         'Assign it
  End If
  
  ImgCB.TransparentGIFColor = vbWhite       'we will make white backgrounds transparent in GIFs

  If chkFlipV.Value <> 1 And chkFlipH.Value <> 1 _
  And chkRotL.Value <> 1 And chkRotR.Value <> 1 And chkRot180.Value <> 1 _
  And chkSize53.Value <> 1 And chkSize23.Value <> 1 _
  And chkSharpen.Value <> 1 And chkBlur.Value <> 1 And chkDeJPG.Value <> 1 _
  And chkEmboss.Value <> 1 And chkEdge.Value <> 1 _
  And chkTrim.Value <> 1 And chkFrame.Value <> 1 _
  And chkMaskY.Value <> 1 And chkInvert.Value <> 1 _
  And chkEqualize.Value <> 1 And chkGamma.Value <> 1 _
  And chkUnGamma.Value <> 1 And chkReplace.Value <> 1 And chkTransparent.Value <> 1 _
  And chkThresh.Value <> 1 And chkExtract.Value <> 1 _
  And chkBlend3.Value <> 1 And chkCombine.Value <> 1 _
  And chkBrightnessUp.Value <> 1 And chkBrightnessDown.Value <> 1 _
  And chkBlend.Value <> 1 And chkInvBlend.Value <> 1 Then              'we do it the easy way
    If chkClip.Value = False Then
      rcOK = ImgCB.SaveObjDCClip(picOriginal, TEMPFILE, PicType, ColorMode, RequiredBPP, CMAPMode, DitherMode)
      
      '02.06.2003 --- Also Demonstrate the Direct Conversion Method for a Loaded File (elapsed times will double)
'      If Len(FileName) > 0 Then
'        Call ImgCB.ConvertImageFile(FileName, TEMPFILE2, PicType, ColorMode, RequiredBPP, CMAPMode, DitherMode)
'      End If
      '---------------
     Else
      rcOK = ImgCB.SaveObjDCClip(picOriginal, TEMPFILE, PicType, ColorMode, RequiredBPP, CMAPMode, DitherMode, _
             200, 200, 639, 479)
    End If

   Else                              'we do it a step at a time
    rcOK = ImgCB.SetPipeline(PicType, ColorMode, RequiredBPP, CMAPMode, DitherMode)
    
    If rcOK Then
      If chkClip.Value = False Then
        rcOK = ImgCB.GetImageObjhDC(picOriginal)
       Else
        rcOK = ImgCB.GetImageObjhDC(picOriginal, 200, 200, 639, 479)
      End If
    End If
    
    If rcOK And chkBlend3 = 1 Then  'Do a Threeway alphaBlend
      Call cmdColorTest_Click
      Set ImgCB2 = New cRVTVBIMG
      rcOK = ImgCB2.GetImageObjhDC(picOriginal)                   'the blended image
      If rcOK Then
        Call cmdGreyTest_Click
        Set ImgCB3 = New cRVTVBIMG
        rcOK = ImgCB3.GetImageObjhDC(picOriginal)                 'the blend mask
        If rcOK Then rcOK = ImgCB.AlphaBlend3(ImgCB2, ImgCB3)
      End If
      Set ImgCB2 = Nothing
      Set ImgCB3 = Nothing
    End If
    
    If rcOK And chkInvert.Value = 1 Then
      rcOK = ImgCB.InvertColor()
    End If
    
    If rcOK And chkEqualize.Value = 1 Then
      rcOK = ImgCB.EqualizeColor()
    End If
    
    If rcOK And chkGamma.Value = 1 Then
      rcOK = ImgCB.GammaCorrectColor(RGamma:=1.8, GGamma:=1.8, BGamma:=1.8)
    End If
    
    If rcOK And chkUnGamma.Value = 1 Then
      rcOK = ImgCB.GammaCorrectColor(RGamma:=-1.8, GGamma:=-1.8, BGamma:=-1.8)
    End If
    
    If rcOK And chkBrightnessUp.Value = 1 Then
      rcOK = ImgCB.BrightnessCorrectColor(RFactor:=0.5, GFactor:=0.5, BFactor:=0.5)
    End If
    
    If rcOK And chkBrightnessDown.Value = 1 Then
      rcOK = ImgCB.BrightnessCorrectColor(RFactor:=-0.5, GFactor:=-0.5, BFactor:=-0.5)
    End If
    
    If rcOK And chkFlipV.Value = 1 Then
      rcOK = ImgCB.FlipVert()
    End If
    
    If rcOK And chkFlipH.Value = 1 Then
      rcOK = ImgCB.FlipHorz()
    End If
    
    If rcOK And chkRotL.Value = 1 Then
      rcOK = ImgCB.RotateL()
    End If
    
    If rcOK And chkRotR.Value = 1 Then
      rcOK = ImgCB.RotateR()
    End If
    
    If rcOK And chkRot180.Value = 1 Then
      rcOK = ImgCB.Rotate180()
    End If
    
    If rcOK And chkSize53.Value = 1 Then
      rcOK = ImgCB.BilinearResize(NewWidth:=1024, NewHeight:=768)
    End If
    
    If rcOK And chkSize23.Value = 1 Then
      rcOK = ImgCB.BilinearResize(NewWidth:=523, NewHeight:=392)
    End If
    
    If rcOK And chkDeJPG.Value = 1 Then
      rcOK = ImgCB.DeJPEG(Factor:=1)                  '1-4
    End If
    
    If rcOK And chkBlur.Value = 1 Then
      rcOK = ImgCB.Blur()
    End If
    
    If rcOK And chkSharpen.Value = 1 Then
      rcOK = ImgCB.Sharpen(Factor:=2)
    End If
    
    If rcOK And chkEmboss.Value = 1 Then
      rcOK = ImgCB.Emboss()
    End If
    
    If rcOK And chkEdge.Value = 1 Then
      rcOK = ImgCB.Edge(Factor:=4)
    End If
    
    If rcOK And chkTrim.Value = 1 Then
      rcOK = ImgCB.Trim(10, 10, 10, 10)
    End If
    
    If rcOK And chkFrame.Value = 1 Then
      rcOK = ImgCB.Frame(LeftRight:=10, TopBottom:=10, FrameRGBColor:=vbMagenta)
    End If
    
    If rcOK And chkMaskY.Value = 1 Then
      rcOK = ImgCB.MaskColor(vbYellow)
    End If
    
    If rcOK And chkBlend.Value = 1 Then
      rcOK = ImgCB.AlphaBlendColor(RGBColor:=&HF09078, RGBAlphaMask:=&H203040)
    End If
    
    If rcOK And chkInvBlend.Value = 1 Then
      rcOK = ImgCB.InverseAlphaBlendColor(RGBColor:=&HF09078, RGBAlphaMask:=&H203040)
    End If
    
    If rcOK And chkCombine.Value = 1 Then
      rcOK = ImgCB.CombineColor(Opcode:=PCC_ADDM, RGBColor:=&H80C040)
    End If
    
    If rcOK And chkReplace.Value = 1 Then
      rcOK = ImgCB.ReplaceColor(TargetRGBColor:=&H808080, SearchRadius:=32, ReplacementRGBColor:=vbRed)
    End If
    
    If rcOK And chkTransparent.Value = 1 Then
      Call OleTranslateColor(picResult.BackColor, 0, Color)
      rcOK = ImgCB.ReplaceColor(TargetRGBColor:=&H0, SearchRadius:=16, ReplacementRGBColor:=Color)
    End If
    
    If rcOK And chkThresh.Value = 1 Then
      rcOK = ImgCB.ThresholdColor(RGBMask:=&HC0C0C0)
    End If
    
    If rcOK And chkExtract.Value = 1 Then
      Set ImgCB = ImgCB.Extract(atX:=10, atY:=270, Width:=300, Height:=200) 'always 24bpp
    End If
    
    If rcOK Then rcOK = ImgCB.DoColorMapping()
    If rcOK Then rcOK = ImgCB.CommitToDisk(TEMPFILE, PicType)
  End If
  
  
  Me.MousePointer = vbDefault

  If rcOK Then
    lblLoadTime.Caption = "  Load        " & Format$(ImgCB.etLoad, "0.000s")
    lblSaveTime.Caption = "  Save        " & Format$(ImgCB.etSave, "0.000s")
'        lblOpTime.Caption = "Operations   " & Format$(ImgCB.etOps, "0.000s")
      lblCMAPTime.Caption = "Quantize     " & Format$(ImgCB.etQuant, "0.000s")
    lblDitherTime.Caption = "Remap        " & Format$(ImgCB.etRemap, "0.000s")
       lblElapsed.Caption = "Elapsed      " & Format$(ImgCB.etElapsed + ImgCB.etLoad, "0.000s")

    Call picResult.ZOrder(0)
    Call picResult.Cls
    Call ImgCB.PutImageObjhDC(picResult)    'Display It
    If ImgCB.LoadFromFile(TEMPFILE) Then    'And Reload It so that we can see if its alright
      Call ImgCB.DisplayImage(Form1.picResult)
    End If
    '    Kill TEMPFILE
   Else
    MsgBox "Something has gone wrong, rcOK was not True" & vbCrLf _
         & "Check that your most recent code change didnt break something"
  End If
  
End Sub

'The user map MUST BE DYNAMICALLY ASSIGNED to make the Assignment work into rvtVBImg
Private Sub GenUserCMap(ByVal NColors As Long)  'We make a Sepia (sortof) ColorMap as a demo

  Dim i As Long, dv As Long, v As Long
  
  ReDim UserCMap(0 To NColors - 1)
  
  For i = 0 To NColors - 1
    With UserCMap(i)
      v = (i * 256) \ NColors: If v > 255 Then v = 255
      dv = (64 * v * (255 - v)) \ 255 \ 255
      .Red = v + dv
      .Green = v
      .Blue = v - dv
    End With
  Next
  
End Sub

Private Sub picOriginal_Click()

  Call Form1.picOriginal.ZOrder(0)

End Sub

Private Sub picResult_Click()

  Call Form1.picResult.ZOrder(0)

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set ImgCB = Nothing

End Sub

'==================================== FORM VALIDATION =========================================================

Private Sub GetPipelineParms(ByRef PicType As Long, _
                             ByRef ColorMode As Long, _
                             ByRef RequiredBPP As Long, _
                             ByRef CMAPMode As Long, _
                             ByRef DitherMode As Long)

 Dim z() As Variant

  z = Array(PIC_BMP, PIC_GIF, PIC_GIF_LACED, PIC_PNM)
  PicType = z(CurPicType)

  z = Array(PIC_COLOR, PIC_BW, PIC_GREY)
  ColorMode = z(CurColorMode)

  z = Array(PIC_1BPP, PIC_2BPP, PIC_3BPP, PIC_4BPP, PIC_5BPP, PIC_6BPP, PIC_7BPP, PIC_8BPP, _
            PIC_9BPP, PIC_12BPP, PIC_15BPP, PIC_16BPP, PIC_24BPP, PIC_32BPP)
  RequiredBPP = z(CurNColors)

  z = Array(PIC_MS_CMAP, PIC_OPTIMAL_CMAP, PIC_FIXED_CMAP, PIC_FIXED_CMAP_VGA, PIC_FIXED_CMAP_INET, _
            PIC_FIXED_CMAP_MS256, PIC_FIXED_CMAP_GREY, PIC_MODIFIED_CMAP, PIC_FIXED_CMAP_USER)
  CMAPMode = z(CurCMAPMode)

  z = Array(PIC_DITHER_NONE, PIC_DITHER_BIN, PIC_DITHER_ORD, PIC_DITHER_HTC, _
            PIC_DITHER_FDIAG, PIC_DITHER_BDIAG, PIC_DITHER_HORZ, PIC_DITHER_VERT, _
            PIC_DITHER_FS1, PIC_DITHER_FS2, PIC_DITHER_FS3, PIC_DITHER_BNM)
  DitherMode = z(CurDitherMode)

End Sub

'------------------------------------------------- PICTURE TYPE --------------------------------------------
Private Sub optPicType_Click(Index As Integer)

  If Index <> CurPicType Then
    If CurPicType <> -1 Then optPicType(CurPicType).Value = False
    CurPicType = Index
    optPicType(CurPicType).Value = True

    CurColorMode = -1
    CurNColors = -1
    CurCMAPMode = -1
    CurDitherMode = -1
    Call SetColorDefaults(VALID_COLOR_OPTIONS, 0)   'Default=COLOR
  End If

End Sub

'------------------------------------------------- COLOR MODE ------------------------------------------------

Private Sub SetColorDefaults(ByVal ColorOptions As Long, ByVal DefaultColorMode As Long)

 Dim i As Long, j As Long

  If ColorOptions <> -1 Then
    j = 1
    For i = 0 To NCOLOR_OPTIONS - 1
      optColorMode(i).Value = False
      If (ColorOptions And j) <> 0 Then
        optColorMode(i).Enabled = True
       Else
        optColorMode(i).Enabled = False
      End If
      j = j + j
    Next i
    If DefaultColorMode <> -1 Then CurColorMode = -1
  End If
  If DefaultColorMode <> -1 Then Call optColorMode_Click(CInt(DefaultColorMode))

End Sub

Private Sub optColorMode_Click(Index As Integer)

 Dim ValidNColors As Long, DefaultNColors As Long

  If Index <> CurColorMode Then
    If CurColorMode <> -1 Then optColorMode(CurColorMode).Value = False
    CurColorMode = Index
    optColorMode(CurColorMode).Value = True

    Select Case CurColorMode
     Case 0:
      Select Case CurPicType
       Case 0:
        ValidNColors = VALID_BMP_DEPTH_OPTIONS: DefaultNColors = 7
       Case 1, 2:
        ValidNColors = VALID_GIF_DEPTH_OPTIONS: DefaultNColors = 7
       Case 3:
        ValidNColors = VALID_PNM_DEPTH_OPTIONS: DefaultNColors = 7
      End Select

     Case 1:
      ValidNColors = VALID_BW_DEPTH_OPTIONS: DefaultNColors = 0

     Case 2:
      Select Case CurPicType
       Case 0:
        ValidNColors = VALID_BMP_DEPTH_OPTIONS And VALID_GREY_DEPTH_OPTIONS: DefaultNColors = 7
       Case 1, 2:
        ValidNColors = VALID_GIF_DEPTH_OPTIONS And VALID_GREY_DEPTH_OPTIONS: DefaultNColors = 7
       Case 3:
        ValidNColors = VALID_PNM_DEPTH_OPTIONS And VALID_GREY_DEPTH_OPTIONS: DefaultNColors = 7
      End Select
    End Select
    Call SetNColorDefaults(VALID_DEPTH_OPTIONS And ValidNColors, DefaultNColors)
  End If

End Sub

'------------------------------------------------- N COLORS --------------------------------------------

Private Sub SetNColorDefaults(ByVal NColorOptions As Long, ByVal DefaultNColor As Long)

 Dim i As Long, j As Long

  If NColorOptions <> -1 Then
    j = 1
    For i = 0 To NNCOLOR_OPTIONS - 1
      OptNColors(i).Value = False
      If (NColorOptions And j) <> 0 Then
        OptNColors(i).Enabled = True
       Else
        OptNColors(i).Enabled = False
      End If
      j = j + j
    Next i
    If DefaultNColor <> -1 Then CurNColors = -1
  End If
  If DefaultNColor <> -1 Then Call optNColors_Click(CInt(DefaultNColor))

End Sub

Private Sub optNColors_Click(Index As Integer)

 Dim ValidCMaps As Long, DefaultCMap As Long

  If Index <> CurNColors Then
    If CurNColors <> -1 Then OptNColors(CurNColors).Value = False
    CurNColors = Index
    OptNColors(CurNColors).Value = True

    Select Case CurNColors
     Case 0:
      ValidCMaps = VALID_BW_CMAP_OPTIONS:      DefaultCMap = 6  '2 Fixed Grey

     Case 1, 2:
      ValidCMaps = VALID_WW_CMAP_OPTIONS:      DefaultCMap = 2  '4,8

     Case 3:
      ValidCMaps = VALID_XX_CMAP_OPTIONS:      DefaultCMap = 2  '16

     Case 4, 5, 6:
      ValidCMaps = VALID_YY_CMAP_OPTIONS:      DefaultCMap = 2  '32,64,128

     Case 7:
      ValidCMaps = VALID_ZZ_CMAP_OPTIONS:      DefaultCMap = 1  '256 'New

     Case 8, 9, 11:
      ValidCMaps = VALID_64K_CMAP_OPTIONS:     DefaultCMap = 2  '32768

     Case 10:
      ValidCMaps = VALID_32K_CMAP_OPTIONS:     DefaultCMap = 2  '65536

     Case 11, 12:
      ValidCMaps = VALID_16M_CMAP_OPTIONS:     DefaultCMap = 0  '16M
    End Select

    Select Case CurPicType
      Case 0:
        Call SetCMAPDefaults(VALID_BMP_CMAP_OPTIONS And ValidCMaps, DefaultCMap)
      Case 1, 2:
        Call SetCMAPDefaults(VALID_GIF_CMAP_OPTIONS And ValidCMaps, DefaultCMap)
      Case 3:
        Call SetCMAPDefaults(VALID_PNM_CMAP_OPTIONS And ValidCMaps, DefaultCMap)
    End Select
  End If

End Sub

'------------------------------------------------- CMAP MODE --------------------------------------------

Private Sub SetCMAPDefaults(ByVal CMapOptions As Long, ByVal DefaultCMap As Long)

 Dim i As Long, j As Long

  If CMapOptions <> -1 Then
    j = 1
    For i = 0 To NCMAP_OPTIONS - 1
      OptCMAP(i).Value = False
      If (CMapOptions And j) <> 0 Then
        OptCMAP(i).Enabled = True
       Else
        OptCMAP(i).Enabled = False
      End If
      j = j + j
    Next i
    If DefaultCMap <> -1 Then CurCMAPMode = -1
  End If
  If DefaultCMap <> -1 Then Call OptCMAP_Click(CInt(DefaultCMap))

End Sub

Private Sub OptCMAP_Click(Index As Integer)

  If Index <> CurCMAPMode Then
    If CurCMAPMode <> -1 Then
      OptCMAP(CurCMAPMode).Value = False
    End If
    CurCMAPMode = Index
    OptCMAP(CurCMAPMode).Value = True

    If CurCMAPMode = 0 Then
      Call SetDitherDefaults(VALID_MSMAP_DITHER_OPTIONS, 0)
     Else
      Call SetDitherDefaults(VALID_DITHER_OPTIONS, 0)
    End If
  End If

End Sub

'------------------------------------------------- DITHER MODE --------------------------------------------
Private Sub SetDitherDefaults(ByVal DitherOptions As Long, ByVal DefaultDither As Long)

 Dim i As Long, j As Long

  If DitherOptions <> -1 Then
    j = 1
    For i = 0 To NDITHER_OPTIONS - 1
      OptDither(i).Value = False
      If (DitherOptions And j) <> 0 Then
        OptDither(i).Enabled = True
       Else
        OptDither(i).Enabled = False
      End If
      j = j + j
    Next i
    If DefaultDither <> -1 Then CurDitherMode = -1
  End If
  If DefaultDither <> -1 Then Call OptDither_Click(CInt(DefaultDither))

End Sub

Private Sub OptDither_Click(Index As Integer)

  If Index <> CurDitherMode Then
    If CurDitherMode <> -1 Then OptDither(CurDitherMode).Value = False
    CurDitherMode = Index
    OptDither(CurDitherMode).Value = True
  End If

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jul-05 11:39) 58 + 450 = 508 Lines
