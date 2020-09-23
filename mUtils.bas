Attribute VB_Name = "mUtils"
Option Explicit

'©2001/2 Ron van Tilburg  - rivit@f1.net.au

Public Declare Sub CopyMemoryRR Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Sub CopyMemoryRV Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Long, ByVal Length As Long)
Public Declare Sub CopyMemoryVR Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByRef Source As Any, ByVal Length As Long)
Public Declare Sub CopyMemoryVV Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

'Recast a RGBAColor as RGBA Type
Public Function GetRGBA(ByVal RGBAColor As Long) As RGBA

  Call CopyMemoryRR(GetRGBA, RGBAColor, 4&)
  
End Function

'Recast a RGBA type as RGBAColor
Public Function PutRGBA(ByRef Pixel As RGBA) As Long

  Call CopyMemoryRR(PutRGBA, Pixel, 4&)
  
End Function

'Some fundamental Flag Manipulation Functions

'Are the requested Flags Set?
Public Function FSet(ByVal FlagGroup As Long, ByVal ReqFlags As Long) As Boolean

  FSet = ((FlagGroup And ReqFlags) = ReqFlags)

End Function

'Are the requested Flags Clr?
Public Function FClr(ByVal FlagGroup As Long, ByVal ReqFlags As Long) As Boolean

  FClr = ((FlagGroup And ReqFlags) = 0)

End Function

'Are the requested Flags Set? and if so Clear Them
Public Function FSetClr(ByRef FlagGroup As Long, ByVal ReqFlags As Long) As Boolean

  FSetClr = ((FlagGroup And ReqFlags) = ReqFlags)
  If FSetClr Then FlagGroup = (FlagGroup And Not ReqFlags)

End Function

'Set the requested Flags
Public Sub SetF(ByRef FlagGroup As Long, ByVal ReqFlags As Long)

  FlagGroup = (FlagGroup Or ReqFlags)

End Sub

'Clear the requested Flags
Public Sub ClrF(ByRef FlagGroup As Long, ByVal ReqFlags As Long)

  FlagGroup = (FlagGroup And Not ReqFlags)

End Sub

'Toggle the requested Flags
Public Sub ToggleF(ByRef FlagGroup As Long, ByVal ReqFlags As Long)

  FlagGroup = (FlagGroup Xor ReqFlags)

End Sub

'mask (select only) the requested Flags
Public Function MaskF(ByVal FlagGroup As Long, ByVal ReqFlags As Long) As Long

  MaskF = (FlagGroup And ReqFlags)

End Function

':) Ulli's VB Code Formatter V2.6.10 (20-Dec-01 15:46:30) 1 + 47 = 48 Lines
