Attribute VB_Name = "modAlpha_Blende"
Public Declare Function FoxxCreatePicture Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function FoxxDeletePicture Lib "alphablend.dll" (ByVal Bitmap As Long) As Long
Public Declare Function FoxxBlendPicture Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Bitmap As Long, ByVal alpha As Byte, Optional ByVal MaskColor As Long, Optional ByVal Flags As Long) As Long
Public Declare Function FoxxBlendPictures Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Pic1 As Long, ByVal Pic2 As Long, ByVal Buffer As Long, ByVal alpha As Byte) As Long
Public Declare Function FoxxCreateFastMask Lib "alphablend.dll" (ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal DstWidth As Long, Optional ByVal DstHeight As Long, Optional ByVal MaskColor As Long, Optional ByVal Flags As Long) As Long
Public Declare Function FoxxDeleteMask Lib "alphablend.dll" (ByVal FoxPicture As Long) As Long
Public Declare Function FoxxFastMask Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal FoxPicture As Long, Optional ByVal Flags As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function FoxCreateFastMask Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxCreateFastData Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxFastMask Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal hMaskDC As Long, ByVal xMask As Long, ByVal yMask As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxMosaic Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Level As Long, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxBrightness Lib "alphablend.dll" (ByVal hDC As Long, ByVal Handle As Long, ByVal hSrcDC As Long, ByVal SrcHandle As Long, ByVal Brightness As Long, Optional ByVal TransColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxBlur Lib "alphablend.dll" (ByVal hDC As Long, ByVal Handle As Long, ByVal hSrcDC As Long, ByVal SrcHandle As Long, Blur As Long, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxInvert Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxGreyScale Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxAlphaBlend Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal alpha As Byte, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxAlphaMask Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal hMaskDC As Long, ByVal xMask As Long, ByVal yMask As Long, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxRotate Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Angle As Double, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxOutline Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal LineColor As Long, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxFlip Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxBumpMap Lib "alphablend.dll" (ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, Optional ByVal MskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxHSL Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Hue As Single, ByVal Saturation As Single, ByVal Lightness As Single, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxHSLRGB Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Hue As Single, ByVal Saturation As Single, ByVal LightnessR As Single, ByVal LightnessG As Single, ByVal LightnessB As Single, ByVal ScaleR As Single, ByVal ScaleG As Single, ByVal ScaleB As Single, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxChrome Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Level As Byte, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxMonochrome Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Level As Byte, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxShift Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Level As Byte, ByVal Shift As Byte, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxPsycho Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Level As Byte, ByVal Shift As Byte, ByVal Effekt As Byte, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxWave Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Size As Long, ByVal Movement As Long, ByVal Shift As Single, ByVal Angle As Double, Optional ByVal MaskColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long
Public Declare Function FoxDrawPreview Lib "alphablend.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal Flags As Long) As Long
Public Declare Function FoxTime Lib "alphablend.dll" (Optional ByVal Flags As FoxTimeFlags) As Long
Public Declare Function FoxCounter Lib "alphablend.dll" (Optional ByVal Flags As FoxCounterFlags) As Long
Public Declare Function FoxTimer Lib "alphablend.dll" (Optional ByVal Time As Long, Optional ByVal Flags As FoxTimeFlags) As Long

Enum FoxEffectFlags
    FOX_USE_MASK = &H1
    FOX_ANTI_ALIAS = &H2
    FOX_CHROME_LINEAR = &H4
    FOX_SRC_INVERT = &H100
    FOX_DST_INVERT = &H200
    FOX_MASK_INVERT = &H400
    FOX_SRC_GREYSCALE = &H1000
    FOX_DST_GREYSCALE = &H2000
    FOX_FLIP_X = &H40000
    FOX_FLIP_Y = &H80000
    FOX_TURN_LEFT = &H10000
    FOX_TURN_RIGHT = FOX_FLIP_X Or FOX_FLIP_Y
    FOX_TURN_90DEG = FOX_TURN_LEFT
    FOX_TURN_180DEG = FOX_TURN_RIGHT
    FOX_TURN_270DEG = FOX_FLIP_X Or FOX_FLIP_Y Or FOX_TURN_LEFT
End Enum

Enum FoxTimeFlags
    FOX_TIME_RESET = &H1
End Enum

Enum FoxCounterFlags
    FOX_COUNTER_RESET = &H1
    FOX_COUNTER_COUNT = &H2
End Enum


