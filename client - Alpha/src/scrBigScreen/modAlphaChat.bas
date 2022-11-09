Attribute VB_Name = "modAlphaChat"
Option Base 0
Option Compare Text
Option Explicit

Public Type OSVERSIONINFO
    dwOSVersionInfoSize                     As Long   ' 32
    dwMajorVersion                          As Long   ' 32
    dwMinorVersion                          As Long   ' 32
    dwBuildNumber                           As Long   ' 32
    dwPlatformId                            As Long   ' 32
    szCSDVersion                            As String * 128
End Type

Public Const GWL_EXSTYLE                    As Long = (-20)
Public Const LWA_ALPHA                      As Long = &H2
Public Const WS_EX_LAYERED                  As Long = &H80000

Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
