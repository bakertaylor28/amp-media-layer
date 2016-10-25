Attribute VB_Name = "modHotKey"
Option Explicit

Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, _
    ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
    
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, _
    ByVal ID As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = -4

Public Const MOD_CTRL = &H2
Public Const MOD_SHFT = &H4
Public Const MOD_ALT = &H1

Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_SEPARATOR = &H6C

Public Const VK_DOWN = &H28
Public Const VK_RIGHT = &H27
Public Const VK_UP = &H26
Public Const VK_LEFT = &H25

Public Const VK_RBUTTON = &H2
Public Const VK_LBUTTON = &H1

Public Const VK_DELETE = &H2E
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_INSERT = &H2D

Public Const VK_SHIFT = &H10
Public Const VK_RETURN = &HD

Public Const VK_ESCAPE = &H1B
Public Const VK_PAUSE = &H13

Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87

Public glWinRet As Long

Public Function CallbackMsgs(ByVal wHwnd As Long, ByVal wmsg As Long, ByVal wp_id As Long, ByVal lp_id As Long) As Long
    If wmsg = WM_HOTKEY Then
        Call DoFunctions(wp_id)
        CallbackMsgs = 1
        Exit Function
    End If
    CallbackMsgs = CallWindowProc(glWinRet, wHwnd, wmsg, wp_id, lp_id)
End Function

Public Sub DoFunctions(ByVal vKeyID As Byte)
    DoEvents
If blnMediaChoosen Then
    Select Case vKeyID
        Case 0 '+
            If Fullscreen = False Then
                If intSize < 300 Then intSize = intSize + 25
                Call ResizeMovie(CCur(intSize))
                frmMedia.lblSize.Caption = intSize & " %"
            End If
        Case 1 '-
            If Fullscreen = False Then
                If intSize > 50 Then intSize = intSize - 25
                Call ResizeMovie(CCur(intSize))
                frmMedia.lblSize.Caption = intSize & " %"
            End If
        Case 2 'F7
            If Fullscreen Then
                Fullscreen = False
            Else
                Fullscreen = True
            End If
            PlayMedia
            blnPause = True
        Case 3 'F5
            If blnPause Then
                PauseMedia
                blnPause = False
            Else
                PlayMedia
                blnPause = True
            End If
        Case 4 'F6
            frmMedia.Timer1.Enabled = False
            frmMedia.lblCurTime.Caption = "0:00:00"
            frmMedia.SeekSlider.Value = 0
            Call MoveMedia(0)
            Call PauseMedia
            frmMedia.cmdMedia(1).Visible = False
            frmMedia.cmdMedia(0).Visible = True
            blnPause = False
        Case 5 'Down Arrow
            If MediaVolume > 0 Then
                MediaVolume = MediaVolume - 10
                If MediaVolume < 0 Then MediaVolume = 0
                SetVolume (MediaVolume)
                frmMedia.VolumeSlider.Value = MediaVolume
            End If
        Case 6 'Up Arrow
            If MediaVolume < 100 Then
                MediaVolume = MediaVolume + 10
                If MediaVolume > 100 Then MediaVolume = 100
                SetVolume (MediaVolume)
                frmMedia.VolumeSlider.Value = MediaVolume
            End If
    End Select
End If
End Sub

