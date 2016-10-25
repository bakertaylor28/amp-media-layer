VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMedia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AMP Media Player  (32-bit DivX )"
   ClientHeight    =   2205
   ClientLeft      =   5430
   ClientTop       =   9150
   ClientWidth     =   5220
   Icon            =   "frmMedia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "DivX Player"
   MaxButton       =   0   'False
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   Begin VB.CommandButton cmdMedia 
      Appearance      =   0  'Flat
      Caption         =   "Help"
      Height          =   255
      Index           =   4
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chkFullscreen 
      Caption         =   "Fullscreen"
      Height          =   195
      Left            =   2280
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1035
   End
   Begin VB.CommandButton cmdSize 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton cmdSize 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Width           =   255
   End
   Begin ComctlLib.Slider VolumeSlider 
      Height          =   315
      Left            =   3960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      _Version        =   327682
      LargeChange     =   100
      SmallChange     =   10
      Max             =   100
      SelStart        =   100
      TickStyle       =   3
      Value           =   100
   End
   Begin VB.CommandButton cmdMedia 
      Appearance      =   0  'Flat
      Caption         =   "Open"
      Height          =   255
      Index           =   3
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin ComctlLib.Slider SeekSlider 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
      _Version        =   327682
      LargeChange     =   0
      SmallChange     =   0
      TickStyle       =   3
      TickFrequency   =   1000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   -120
   End
   Begin VB.CommandButton cmdMedia 
      Appearance      =   0  'Flat
      Caption         =   "Stop"
      Height          =   375
      Index           =   2
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdMedia 
      Appearance      =   0  'Flat
      Caption         =   "Play"
      Height          =   375
      Index           =   0
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdMedia 
      Appearance      =   0  'Flat
      Caption         =   "Pause"
      Height          =   375
      Index           =   1
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   1680
      Width           =   435
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Seek"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      Height          =   195
      Left            =   4080
      TabIndex        =   8
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label lblCurTime 
      BackStyle       =   0  'Transparent
      Caption         =   "0:00:00"
      Height          =   195
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resize"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   555
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File "
      Begin VB.Menu MnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu MnuBaseHlp 
      Caption         =   "Help"
      Begin VB.Menu MnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu MnuAboutbox 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'This is a application to play .avi, .asf, .mpg, .mpeg, .wmv videos.
'You need to have all necessary codecs already installed on your computer.
'This player can increase/decrease movie size by 25 % up to 300 %
'And of course Fullscreen.
Option Explicit
Private CD As CommonDialog

'Fullscreen
Private Sub chkFullscreen_Click()
If chkFullscreen.Value = 1 Then
    chkFullscreen.Value = 0
    Fullscreen = True
    PlayMedia
Else
    Fullscreen = False
    PlayMedia
End If
blnPause = True
End Sub

'The Commandobuttons- Open and help buttons are hidden and provided for as Menu Funtions.

Private Sub cmdMedia_Click(Index As Integer)
Dim ret As Long
Dim tmp As String
Select Case Index
    Case 0 'Play
        If blnMediaChoosen Then
            PlayMedia
            blnPause = True
        End If
    Case 1 'Pause
        If blnMediaChoosen Then
            PauseMedia
            blnPause = False
        End If
    Case 2 'Stop
        If blnMediaChoosen Then
            Timer1.Enabled = False
            lblCurTime.Caption = "0:00:00"
            SeekSlider.Value = 0
            Call MoveMedia(0)
            Call PauseMedia
            blnPause = False
        End If
    Case 3 'Open
        If blnMediaChoosen Then
            Clear
            CloseMedia
        End If
        intSize = 100
        lblSize.Caption = intSize & " %"
        CD.ShowOpen
        tmp = CD.FileName
        If tmp <> "" Then
            MediaPath = """" & tmp & """"
            ret = OpenMedia
            If ret = 0 Then
                Call GetSize
               
                SeekSlider.Max = MediaLengthMS
                ResizeMovie
                Call SetVolume(VolumeSlider.Value)
                PlayMedia
                Timer1.Enabled = True
                blnPause = True
            Else
                MsgBox "Media cant be played!", vbCritical
            End If
        End If
    Case 4
        frmHelp.Show
End Select
End Sub

'Increase/Decrease Moviesize
Private Sub cmdSize_Click(Index As Integer)
If blnMediaChoosen Then
    Select Case Index
        Case 0
            If intSize < 300 Then intSize = intSize + 25
        Case 1
            If intSize > 50 Then intSize = intSize - 25
    End Select
    If intSize = 100 Then
        Call ResizeMovie
    Else
        Call ResizeMovie(CCur(intSize))
    End If
    lblSize.Caption = intSize & " %"
End If
End Sub

Private Sub Form_Load()
Dim ret As Long
'Set initial values
intSize = 100
MediaVolume = 100
Set CD = New CommonDialog
CD.Filter = "Supported Media Files|*.avi;*.asf;*.mpg;*.mpeg;*.wmv|DivX File (*.avi)|*.avi"
CD.DialogTitle = "Choose media to play"

'Disable the screensaver
Call ScreenSaverActive(False)

'Register the hotkeys
ret = RegisterHotKey(Me.hwnd, 0, MOD_CTRL, VK_ADD)
ret = RegisterHotKey(Me.hwnd, 1, MOD_CTRL, VK_SUBTRACT)
ret = RegisterHotKey(Me.hwnd, 2, MOD_CTRL, VK_F7)
ret = RegisterHotKey(Me.hwnd, 3, MOD_CTRL, VK_F5)
ret = RegisterHotKey(Me.hwnd, 4, MOD_CTRL, VK_F6)
ret = RegisterHotKey(Me.hwnd, 5, MOD_CTRL, VK_DOWN)
ret = RegisterHotKey(Me.hwnd, 6, MOD_CTRL, VK_UP)

' Subclassing the form to get the Windows callback msgs.
glWinRet = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf CallbackMsgs)

'Incase the player get associated with a movie format
If Command <> "" Then
    frmMedia.Show
    lblSize.Caption = intSize & " %"
    MediaPath = """" & Command & """"
    ret = OpenMedia
    If ret = 0 Then
        Call GetSize
        
        SeekSlider.Max = MediaLengthMS
        ResizeMovie
        Call SetVolume(VolumeSlider.Value)
        PlayMedia
        Timer1.Enabled = True
        blnPause = True
    Else
        MsgBox "Media cant be played!", vbCritical
    End If
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
CloseMedia
Unload frmHelp

'Enable the screensaver
Call ScreenSaverActive(True)

'Unregister the hotkeys
For i = 0 To 6
    UnregisterHotKey Me.hwnd, i
Next
End Sub

Private Sub MnuAboutbox_Click()
MsgBox "AMP Media Player V. 1.0    Copyright 2016, J. Kevin Pelham, All Rights Reserved.", vbOKOnly, "About AMP Media Player"
End Sub

Private Sub MnuHelp_Click()
Dim ret As Long
Dim tmp As String
 frmHelp.Show
End Sub

Private Sub MnuOpen_Click()
Dim ret As Long
Dim tmp As String
        If blnMediaChoosen Then
            Clear
            CloseMedia
        End If
        intSize = 100
        lblSize.Caption = intSize & " %"
        CD.ShowOpen
        tmp = CD.FileName
        If tmp <> "" Then
            MediaPath = """" & tmp & """"
            ret = OpenMedia
            If ret = 0 Then
                Call GetSize
                
                SeekSlider.Max = MediaLengthMS
                ResizeMovie
                Call SetVolume(VolumeSlider.Value)
                PlayMedia
                Timer1.Enabled = True
                blnPause = True
            Else
                MsgBox "Media cant be played!", vbCritical
            End If
        End If
        
End Sub

'Seek to a choosen point in the movie
Private Sub SeekSlider_Click()
If blnMediaChoosen Then Call MoveMedia(SeekSlider.Value)
End Sub

Private Sub SeekSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If blnMediaChoosen Then Timer1.Enabled = False
End Sub

Private Sub SeekSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If blnMediaChoosen Then
    Timer1.Enabled = True
    blnPause = True
End If
End Sub

'Choose Volume 0 to 100 %
Private Sub VolumeSlider_Click()
If blnMediaChoosen Then Call SetVolume(VolumeSlider.Value)
MediaVolume = VolumeSlider.Value
End Sub

'Get current position in media
'If end of movie rewind and pause
'Set On/Off Play/Pause commandbuttons
Private Sub Timer1_Timer()
RunTime = GetCurrentMediaPos
If RunTime < MediaLengthMS Then
    SeekSlider.Value = RunTime
    lblCurTime.Caption = Format(FormatCount(RunTime), "h:mm:ss")
Else
    Timer1.Enabled = False
    lblCurTime.Caption = "0:00:00"
    SeekSlider.Value = 0
    Call MoveMedia(0)
    Call PauseMedia
    blnPause = False
End If

Select Case blnPause
    Case False
        cmdMedia(1).Visible = False
        cmdMedia(0).Visible = True
    Case True
        cmdMedia(1).Visible = True
        cmdMedia(0).Visible = False
End Select
End Sub
'Disable Timer and set some initial values
Private Sub Clear()
Timer1.Enabled = False
SeekSlider.Value = 0

lblCurTime.Caption = "0:00:00"
cmdMedia(1).Visible = False
cmdMedia(0).Visible = True
End Sub

