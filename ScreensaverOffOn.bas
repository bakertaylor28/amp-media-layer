Attribute VB_Name = "ScreensaverOffOn"
Option Explicit

Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
                                (ByVal uAction As Long, ByVal uParam As Long, _
                                ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long

Public Sub ScreenSaverActive(Active As Boolean)
Dim Enabled As Long
Dim ret As Long

Enabled = IIf(Active, 1, 0)
ret = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, Enabled, 0, 0)
End Sub


