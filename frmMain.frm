VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{5FFA8403-F5BE-4496-BEBB-E071BE4EC0AB}#1.0#0"; "AudioLed.ocx"
Object = "{8579AD78-DFAC-4DA8-B567-CAF4FF08AC1E}#1.0#0"; "XRFMK01.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Ausgefüllt
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9885
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdSlowFadeMPlay1 
      BackColor       =   &H00848284&
      Caption         =   "< Slow Fade"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1140
      Style           =   1  'Grafisch
      TabIndex        =   27
      Top             =   4920
      Width           =   1275
   End
   Begin VB.CheckBox chkFullVolMPlay2 
      Caption         =   "Full Volume"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   26
      Top             =   3420
      Width           =   1335
   End
   Begin VB.CheckBox chkFullVolMPlay1 
      Caption         =   "Full Volume"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   25
      Top             =   1980
      Width           =   1335
   End
   Begin VB.CheckBox chkContPlay2 
      Caption         =   "Continues Play"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1980
      TabIndex        =   24
      Top             =   3420
      Width           =   1635
   End
   Begin VB.CheckBox chkMuteMPlay2 
      Caption         =   "Mute"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   23
      Top             =   3420
      Width           =   735
   End
   Begin VB.CheckBox chkContPlay1 
      Caption         =   "Continues Play"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1980
      TabIndex        =   22
      Top             =   1980
      Width           =   1575
   End
   Begin VB.CheckBox chkMuteMplay1 
      Caption         =   "Mute"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   21
      Top             =   1980
      Width           =   795
   End
   Begin VB.CommandButton cmdSyncStop 
      Caption         =   "Syncronous Stop"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5100
      TabIndex        =   19
      Top             =   5400
      Width           =   4035
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4140
      TabIndex        =   20
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdSyncPlay 
      Caption         =   "Syncronous Play"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   18
      Top             =   5400
      Width           =   3975
   End
   Begin MSComctlLib.ProgressBar ProgressBar3 
      Height          =   5475
      Left            =   9360
      TabIndex        =   17
      Top             =   300
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   9657
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Max             =   30
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdFadeMPlay2 
      Caption         =   "Fade >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdCenter 
      Caption         =   "Center"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4140
      TabIndex        =   15
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdFadeMPlay1 
      Caption         =   "<< Fade"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   14
      Top             =   4920
      Width           =   975
   End
   Begin XRS_File_Managing_Kit_1.FileBrowser FileBrowser1 
      Height          =   315
      Left            =   60
      TabIndex        =   12
      Top             =   1260
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   556
      ForeColor       =   16777215
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      TextFieldEnabled=   0   'False
      ButtonEnabled   =   0   'False
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   2760
      TabIndex        =   10
      ToolTipText     =   "Player1 Master Output"
      Top             =   300
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Max             =   20
      Scrolling       =   1
   End
   Begin LED.hLED hLED1 
      Height          =   255
      Left            =   60
      TabIndex        =   6
      ToolTipText     =   "Player1 Left Output"
      Top             =   300
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   450
   End
   Begin ComctlLib.Slider BalSlide 
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   979
      _Version        =   327682
      MousePointer    =   9
      Max             =   100
      TickStyle       =   2
   End
   Begin LED.hLED hLED2 
      Height          =   255
      Left            =   60
      TabIndex        =   7
      ToolTipText     =   "Player1 Right Output"
      Top             =   660
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   450
   End
   Begin LED.hLED hLED3 
      Height          =   255
      Left            =   4740
      TabIndex        =   8
      ToolTipText     =   "Player2 Left Output"
      Top             =   300
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   450
   End
   Begin LED.hLED hLED4 
      Height          =   255
      Left            =   4740
      TabIndex        =   9
      ToolTipText     =   "Player2 Right Output"
      Top             =   660
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   450
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   615
      Left            =   7440
      TabIndex        =   11
      ToolTipText     =   "Player2 Master OutPut"
      Top             =   300
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Max             =   20
      Scrolling       =   1
   End
   Begin XRS_File_Managing_Kit_1.FileBrowser FileBrowser2 
      Height          =   315
      Left            =   60
      TabIndex        =   13
      Top             =   2700
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   556
      ForeColor       =   13684944
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      TextFieldEnabled=   0   'False
      ButtonEnabled   =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   62
      Left            =   2820
      Top             =   360
   End
   Begin VB.Timer timBalSlideMPlay2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   8340
      Top             =   360
   End
   Begin VB.Timer timBalSlideMPlay1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   7920
      Top             =   360
   End
   Begin VB.Timer timBalSlideCenter 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   7500
      Top             =   360
   End
   Begin VB.CommandButton cmdSlowFadeMPlay2 
      BackColor       =   &H00848284&
      Caption         =   "Slow Fade >"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6900
      Style           =   1  'Grafisch
      TabIndex        =   28
      Top             =   4920
      Width           =   1275
   End
   Begin VB.Label Label8 
      BackColor       =   &H00848084&
      Caption         =   " Deck 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D0D0D0&
      Height          =   675
      Left            =   0
      TabIndex        =   31
      Top             =   2400
      Width           =   9375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00848284&
      Caption         =   " Deck 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D0D0D0&
      Height          =   675
      Left            =   0
      TabIndex        =   30
      Top             =   960
      Width           =   9375
   End
   Begin MediaPlayerCtl.MediaPlayer MPlay1 
      Height          =   675
      Left            =   0
      TabIndex        =   29
      Top             =   1680
      Width           =   9315
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   -1  'True
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -20
      WindowlessVideo =   0   'False
   End
   Begin VB.Image cmdClose 
      Height          =   195
      Left            =   9600
      Top             =   60
      Width           =   255
   End
   Begin VB.Image cmdMinimize 
      Height          =   195
      Left            =   9120
      Top             =   60
      Width           =   195
   End
   Begin VB.Label Label5 
      BackColor       =   &H00848284&
      Caption         =   "0%"
      ForeColor       =   &H8000000A&
      Height          =   195
      Left            =   9060
      TabIndex        =   5
      Top             =   4680
      Width           =   315
   End
   Begin VB.Label Label4 
      BackColor       =   &H00848284&
      Caption         =   "100%"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   8940
      TabIndex        =   4
      Top             =   3840
      Width           =   435
   End
   Begin VB.Label Label3 
      BackColor       =   &H00848284&
      Caption         =   "100%"
      ForeColor       =   &H8000000A&
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Width           =   435
   End
   Begin VB.Label Label2 
      BackColor       =   &H00848284&
      Caption         =   "   0%"
      ForeColor       =   &H8000000A&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   3840
      Width           =   495
   End
   Begin MediaPlayerCtl.MediaPlayer MPlay2 
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   9315
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   65280
      DisplayMode     =   0
      DisplaySize     =   0
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Image Titlebar 
      Height          =   270
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Top             =   0
      Width           =   9870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00848284&
      Caption         =   "50%"
      ForeColor       =   &H8000000A&
      Height          =   195
      Left            =   180
      TabIndex        =   32
      Top             =   3840
      Width           =   8955
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00848284&
      Caption         =   "50%"
      ForeColor       =   &H8000000A&
      Height          =   195
      Left            =   180
      TabIndex        =   33
      Top             =   4680
      Width           =   8955
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Single
Dim SetX As Boolean
Dim y As Single
Dim SetY As Boolean
Dim FadeTo As String
Dim DifX As Single
Dim DifY As Single
Dim AllX As Single
Dim AllY As Single
Dim PosX As Single
Dim PosY As Single

Private Sub BalSlide_Change() ': On Error Resume Next
    'If chkFullVol.value >= 1 Then
    '    Exit Sub
    'ElseIf chkFullVol.value = 2 Then
    If chkFullVolMPlay1.Value = 1 Or chkFullVolMPlay2.Value = 1 Then
        If chkFullVolMPlay1.Value = 1 Then
            MPlay1.Volume = 0
            MPlay2.Volume = -((100 - BalSlide.Value) * 50)
            If BalSlide.Value <= BalSlide.Min Then MPlay2.Volume = -10000
        ElseIf chkFullVolMPlay2.Value = 1 Then
            MPlay1.Volume = -(BalSlide.Value * 50)
            MPlay2.Volume = 0
            If BalSlide.Value >= BalSlide.Max Then MPlay1.Volume = -10000
        End If
    Else
        MPlay1.Volume = -(BalSlide.Value * 50)
        MPlay2.Volume = -((100 - BalSlide.Value) * 50)
        If BalSlide.Value <= BalSlide.Min Then MPlay2.Volume = -10000
        If BalSlide.Value >= BalSlide.Max Then MPlay1.Volume = -10000
    End If
End Sub

Private Sub BalSlide_Scroll()
    Call BalSlide_Change
End Sub

Private Sub chkContPlay1_Click()
    If chkContPlay1.Value = 1 Then
        MPlay1.PlayCount = 0
    Else
        MPlay1.PlayCount = 1
    End If
End Sub

Private Sub chkContPlay2_Click()
    If chkContPlay2.Value = 1 Then
        MPlay2.PlayCount = 0
    Else
        MPlay2.PlayCount = 1
    End If
End Sub

Private Sub chkFullVolMPlay1_Click()
    If chkFullVolMPlay1.Value = 1 Then
        MPlay1.Volume = 0
    Else
        Call BalSlide_Change
    End If
End Sub

Private Sub chkFullVolMPlay2_Click()
    If chkFullVolMPlay2.Value = 1 Then
        MPlay2.Volume = 0
    Else
        Call BalSlide_Change
    End If

End Sub

Private Sub chkMuteMplay1_Click()
    If chkMuteMplay1.Value = 1 Then
        MPlay1.Mute = True
    Else
        MPlay1.Mute = False
    End If
End Sub

Private Sub chkMuteMplay2_Click()
    If chkMuteMPlay2.Value = 1 Then
        MPlay2.Mute = True
    Else
        MPlay2.Mute = False
    End If
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub cmdCenter_Click(): On Error Resume Next
    timBalSlideMPlay1.Enabled = False
    timBalSlideMPlay2.Enabled = False
    timBalSlideCenter.Enabled = True
End Sub

Private Sub cmdClose_Click()
    x = MsgBox("Really Quit?", vbOKCancel)
    If x = VbMsgBoxResult.vbOK Then
        End
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdMinimize_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub cmdFadeMPlay1_Click()
    timBalSlideMPlay1.Interval = 20
    timBalSlideCenter.Interval = 20
    timBalSlideMPlay2.Interval = 20
    timBalSlideMPlay2.Enabled = False
    timBalSlideCenter.Enabled = False
    timBalSlideMPlay1.Enabled = True
End Sub

Private Sub cmdFadeMPlay2_Click()
    timBalSlideMPlay1.Interval = 20
    timBalSlideCenter.Interval = 20
    timBalSlideMPlay2.Interval = 20
    timBalSlideMPlay1.Enabled = False
    timBalSlideCenter.Enabled = False
    timBalSlideMPlay2.Enabled = True
End Sub


Private Sub cmdSlowFadeMPlay1_Click()
    timBalSlideMPlay1.Interval = 75
    timBalSlideCenter.Interval = 75
    timBalSlideMPlay2.Interval = 75
    timBalSlideCenter.Enabled = False
    timBalSlideMPlay2.Enabled = False
    timBalSlideMPlay1.Enabled = True
End Sub

Private Sub cmdSlowFadeMPlay2_Click()
    timBalSlideMPlay1.Interval = 75
    timBalSlideCenter.Interval = 75
    timBalSlideMPlay2.Interval = 75
    timBalSlideCenter.Enabled = False
    timBalSlideMPlay1.Enabled = False
    timBalSlideMPlay2.Enabled = True
End Sub

Private Sub cmdSyncPlay_Click(): On Error Resume Next
    MPlay1.Play
    MPlay2.Play
End Sub

Private Sub cmdSyncStop_Click(): On Error Resume Next
    MPlay1.Stop
    MPlay2.Stop
End Sub

Private Sub FileBrowser1_Change()
    MPlay1.Open FileBrowser1.FileName
End Sub

Private Sub FileBrowser2_Change()
    MPlay2.Open FileBrowser2.FileName
End Sub

Private Sub Form_Load()
    Call BalSlide_Change
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub MPlay1_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    If MPlay1.PlayState = mpClosed Or MPlay1.PlayState = mpStopped Then
        x = 0
        hLED1.Value = x
        hLED2.Value = x
        ProgressBar1.Value = ProgressBar1.Min
    End If
End Sub

Private Sub MPlay2_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    If MPlay2.PlayState = mpClosed Or MPlay2.PlayState = mpStopped Then
        y = 0
        hLED3.Value = y
        hLED4.Value = y
        ProgressBar2.Value = ProgressBar2.Min
    End If
End Sub

Private Sub SysTray_MouseDown(Button As Integer, Id As Long)
    Me.WindowState = vbNormal
    Me.Show
    Me.SetFocus
End Sub

Private Sub timBalSlideCenter_Timer()
    If BalSlide.Value > 50 Then
        i = BalSlide.Value
        BalSlide.Value = i - 1
    ElseIf BalSlide.Value < 50 Then
        i = BalSlide.Value
        BalSlide.Value = i + 1
    ElseIf BalSlide.Value = 50 Then
        timBalSlideCenter.Enabled = False
    End If
End Sub

Private Sub timBalSlideMPlay1_Timer()
    If BalSlide.Value = BalSlide.Min Then timBalSlideMPlay1.Enabled = False: Exit Sub
    i = BalSlide.Value
    BalSlide.Value = i - 1
End Sub

Private Sub timBalSlideMPlay2_Timer(): On Error Resume Next
    If BalSlide.Value = BalSlide.Max Then timBalSlideMPlay2.Enabled = False: Exit Sub
    i = BalSlide.Value
    BalSlide.Value = i + 1
End Sub

Private Sub Timer1_Timer(): On Error Resume Next
    FileBrowser1.TextFieldEnabled = False
    FileBrowser1.ButtonEnabled = True
    FileBrowser2.TextFieldEnabled = False
    FileBrowser2.ButtonEnabled = True
    If MPlay1.PlayState = mpPlaying Then Call Vis1
    If MPlay2.PlayState = mpPlaying Then Call Vis2
    ProgressBar3.Value = (ProgressBar1.Value + ProgressBar2.Value)
End Sub

Public Sub Vis1(): On Error Resume Next
    x = Rnd(10) * 10
    hLED1.Value = x
    hLED2.Value = (x + Rnd(2) * 2)
    x = hLED1.Value + hLED2.Value
    If x > ProgressBar1.Max Then x = ProgressBar1.Max
    ProgressBar1.Value = x
End Sub

Public Sub Vis2(): On Error Resume Next
    y = Rnd(10) * 10
    hLED3.Value = y
    hLED4.Value = (y + Rnd(2) * 2)
    y = hLED3.Value + hLED4.Value
    If y > ProgressBar2.Max Then y = ProgressBar2.Max
    ProgressBar2.Value = y
End Sub

Private Sub Titlebar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DifX = x
    DifY = y
    AllX = x + Me.Left
    AllY = y + Me.Top
End Sub

Private Sub Titlebar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
            
    If Button = 1 Then
        
        With frmMain
                
            If x <> DifX Then
                
                If x < DifX Then
                    PosX = (.Left - (DifX - x))
                Else
                    PosX = (.Left + (x - DifX))
                End If
                
                If y < DifY Then
                    PosY = (.Top - (DifY - y))
                Else
                    PosY = (.Top + (y - DifY))
                End If
            End If
            
            .Move PosX, PosY
        
        End With
        
    End If
    
End Sub
