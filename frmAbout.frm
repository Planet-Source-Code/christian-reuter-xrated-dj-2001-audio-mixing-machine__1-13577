VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   2310
   ClientLeft      =   2340
   ClientTop       =   1650
   ClientWidth     =   5730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1594.403
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":447A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   1665
      Width           =   1260
   End
   Begin VB.Label cmdEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "nocturn@gmx.net"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2460
      MouseIcon       =   "frmAbout.frx":88F4
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   6
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Innen ausgefüllt
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1024.973
      Y2              =   1024.973
   End
   Begin VB.Label lblDescription 
      Caption         =   "Send comments to: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1020
      TabIndex        =   2
      Top             =   1140
      Width           =   1485
   End
   Begin VB.Label lblTitle 
      Caption         =   "Name der Anwendung"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   2145
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1035.327
      Y2              =   1035.327
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   2145
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "We don't take care of damages to your Computer caused by using this Software. YOU ARE USING THIS SOFTWARE ON YOUR OWN RISK!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   255
      TabIndex        =   3
      Top             =   1665
      Width           =   3630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AllX As Single
Dim AllY As Single
Dim PosX As Single
Dim PosY As Single
Dim DifX As Single
Dim DifY As Single
Dim x As Single
Dim y As Single

Private Sub cmdEmail_Click(): On Error Resume Next
    Dim Res As String
    Dim tmp As String * 255
    Dim Windir As String
    
    Res = GetWindowsDirectory(tmp, 255)
    Windir = left$(tmp, Res)
    Shell Windir & "\explorer.exe ""mailto:nocturn@gmx.net?subject=-=Xrated DJ-2001 Public BETA=-"" ", vbNormalFocus
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = _
        "Version " & App.Major & "." & App.Minor & " Build#" & App.Revision
    
    lblTitle.Caption = _
        "Xrated DJ-2001 Public Beta" & Chr(10) & _
        "(C)2000 Xrated Software"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DifX = x
    DifY = y
    AllX = x + Me.left
    AllY = y + Me.top
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
            
    If Button = 1 Then
        
        With frmAbout
                
            If x <> DifX Then
                
                If x < DifX Then
                    PosX = (.left - (DifX - x))
                Else
                    PosX = (.left + (x - DifX))
                End If
                
                If y < DifY Then
                    PosY = (.top - (DifY - y))
                Else
                    PosY = (.top + (y - DifY))
                End If
            End If
            
            .Move PosX, PosY
        
        End With
        
    End If

End Sub

