Attribute VB_Name = "ModMedia"
Option Explicit

Public MediaPath As String          'Holds the path to the mediafile, also used as alias
Public MediaHeight As Long          'The media Height in pixels
Public MediaWidth As Long           'The media Width in pixels
Public MediaLengthMS As Long        'The Duration of the media in milliseconds
Public MediaLengthFrames As Long    'The media total no of frames
Public blnMediaChoosen As Boolean        'Tells if a media has been choosen
Public intSize As Integer           'Tells the size in percent to show media
Public RunTime As Long              'Tells current media positon
Public blnPause As Boolean          'Tells to play or pause
Public Fullscreen As Boolean        'Tells to play fullscreen or window
Public MediaVolume As Integer       'Percentage of full volume

'***API used to communicate with Media Device***'
Public Declare Function mciSendString Lib "winmm.dll" Alias _
        "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
        lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
        hwndCallback As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wmsg As Long, _
        ByVal wParam As Integer, ByVal lParam As Long) As Long

        
Public Function OpenMedia() As Long
Dim CmdStr As String
Dim ret As Long

CmdStr = "Open " & MediaPath & " Type MPEGVideo alias " & MediaPath
ret = mciSendString(CmdStr, 0&, 0&, 0&)
OpenMedia = ret
End Function

Public Sub PlayMedia()
Dim CmdStr As String
Dim ret As Long
If Fullscreen Then
    ret = mciSendString("Play " & MediaPath & " fullscreen", 0&, 0&, 0&)
Else
    ret = mciSendString("Play " & MediaPath, 0&, 0&, 0&)
End If

blnMediaChoosen = True
End Sub

Public Sub PauseMedia()
Dim ret As Long
ret = mciSendString("Pause " & MediaPath, 0&, 0&, 0&)
End Sub

Public Sub StopMedia()
Dim ret As Long
ret = mciSendString("Stop " & MediaPath, 0&, 0&, 0&)
End Sub

Public Sub CloseMedia()
Dim CmdStr As String
Dim ret As Long
CmdStr = "Close all"
ret = mciSendString(CmdStr, 0&, 0&, 0&)
End Sub

Public Sub GetSize()
Dim ret As Long
Dim size As String * 128
Dim var() As String
ret = mciSendString("Where " & MediaPath & " destination", size, 128, 0&)
var = Split(size, " ", -1)
MediaWidth = CCur(var(2))
MediaHeight = CCur(var(3))
End Sub

Public Sub ResizeMovie(Optional Multiplie As Currency)
Dim ret As Long
If Multiplie <> 0 Then
    Multiplie = Multiplie / 100
    ret = mciSendString("put " & MediaPath & " window at " & 0 & " " & 0 & " " & _
                        MediaWidth * Multiplie & " " & MediaHeight * Multiplie, 0&, 0&, 0&)
Else
    ret = mciSendString("put " & MediaPath & " window at " & 0 & " " & 0 & " " & _
                        MediaWidth & " " & MediaHeight, 0&, 0&, 0&)
End If
End Sub

Public Function GetCurrentMediaPos() As Long
Dim ret As Long
Dim pos As String * 128
ret = mciSendString("set " & MediaPath & " time format ms", pos, 128, 0&)
ret = mciSendString("status " & MediaPath & " position", pos, 128, 0&)

If ret <> 0 Then
    GetCurrentMediaPos = -1
    Exit Function
End If

GetCurrentMediaPos = Val(CLng(pos))
End Function

Public Sub MoveMedia(Where As Long)
Dim ret As Long
Dim pos As String * 128
blnMediaChoosen = True
ret = mciSendString("set " & MediaPath & " time format ms", pos, 128, 0&)
ret = mciSendString("seek " & MediaPath & " to " & Where, 0&, 0&, 0&)
ret = mciSendString("Play " & MediaPath, 0&, 0&, 0&)
End Sub

Public Function MediaDuration() As String
Dim ret As Long
Dim TotalTime As String * 128

ret = mciSendString("set " & MediaPath & " time format frames", 0&, 0&, 0&)
ret = mciSendString("status " & MediaPath & " length", TotalTime, 128, 0&)
MediaLengthFrames = Val(TotalTime)

ret = mciSendString("set " & MediaPath & " time format ms", TotalTime, 128, 0&)
ret = mciSendString("status " & MediaPath & " length", TotalTime, 128, 0&)

MediaLengthMS = Val(TotalTime)
MediaDuration = FormatCount(Val(TotalTime))

End Function

Public Function FormatCount(Count As Long) As String
Dim Days As Integer, Hours As Long, Minutes As Long, Seconds As Long

Count = Count \ 1000
Days = Count \ (24& * 3600&)
If Days > 0 Then Count = Count - (24& * 3600& * Days)
Hours = Count \ 3600&
If Hours > 0 Then Count = Count - (3600& * Hours)
Minutes = Count \ 60
Seconds = Count Mod 60

FormatCount = Hours & ":" & Minutes & ":" & Seconds
End Function

Public Sub SetVolume(Volume As Long)
Dim CmdStr As String
Dim ret As Long
CmdStr = "setaudio " & MediaPath & " Volume to " & (Volume * 10)
ret = mciSendString(CmdStr, 0&, 0&, 0&)
End Sub
