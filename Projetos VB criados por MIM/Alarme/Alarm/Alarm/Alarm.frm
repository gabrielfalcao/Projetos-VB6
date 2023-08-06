VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Deepak Naidu's -- Digital Alarm"
   ClientHeight    =   780
   ClientLeft      =   8895
   ClientTop       =   225
   ClientWidth     =   2280
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00626366&
   Icon            =   "Alarm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      Picture         =   "Alarm.frx":7BBA
      ScaleHeight     =   849.057
      ScaleMode       =   0  'User
      ScaleWidth      =   2250
      TabIndex        =   1
      ToolTipText     =   "Deepak Naidu"
      Top             =   0
      Width           =   2280
      Begin VB.OptionButton Radio2 
         Caption         =   "Pm"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   1080
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton Radio1 
         Caption         =   "Am"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   120
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "C"
         Height          =   195
         Left            =   1680
         TabIndex        =   10
         ToolTipText     =   "Cancel Alarm"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Set"
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         ToolTipText     =   "On Alarm"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         MaxLength       =   2
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         MaxLength       =   2
         TabIndex        =   7
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Height          =   135
         Left            =   1920
         TabIndex        =   4
         ToolTipText     =   "Set Alarm"
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         FillColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   1920
         Picture         =   "Alarm.frx":802F
         ScaleHeight     =   75
         ScaleWidth      =   195
         TabIndex        =   3
         ToolTipText     =   "Stop Alarm"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minute"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hour"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00B1B59F&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "DigitMed"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0066676A&
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1770
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4200
      Top             =   2040
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   1080
      TabIndex        =   13
      Top             =   1200
      Width           =   3255
   End
   Begin MediaPlayerCtl.MediaPlayer Media 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   2895
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   1
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   "c:\Alarm\Alarm.wav"
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
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim h As String
Dim m As String
Dim th As String
Dim tm As String
Dim rd As String
Dim radio As String









Private Sub Picture2_Click()
Picture2.Visible = False
Media.Volume = -9640
th = 0
tm = 0
h = 0
m = 0
Text1.Text = ""
Text2.Text = ""
Beep
End Sub






Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
MsgBox ("Enter numbers")
Text1.Text = ""
Text1.SetFocus
End If

End Sub

Private Sub Timer1_Timer()
Dim time As String
time = Now
time = Right(time, 11)
th = Left(time, 2)
th = Trim(th)
tm = Minute(time)
radio = Right(time, 2)
Label1.Caption = time

Label4.Caption = rd

If h = th And m = tm And rd = radio Then
 Media.Play
End If

End Sub



Private Sub Command1_Click()
Label2.Visible = True
Label3.Visible = True
Command2.Visible = True
Command3.Visible = True
Text1.Visible = True
Text2.Visible = True
Command1.Visible = False
Label1.Visible = False
Radio1.Visible = True
Radio2.Visible = True
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Then
o = MsgBox("   Invalid Entry !", vbOKOnly, "       Alert")
Text1.SetFocus
Else

h = Text1.Text
m = Text2.Text

If Radio1.Value = True Then
rd = "AM"
ElseIf Radio2.Value = True Then
rd = "PM"
End If

Label2.Visible = False
Label3.Visible = False
Command2.Visible = False
Command3.Visible = False
Text1.Visible = False
Text2.Visible = False
Command1.Visible = True
Label1.Visible = True
Picture2.Visible = True
Media.Volume = 0
Radio1.Visible = False
Radio2.Visible = False
End If

End Sub
Private Sub Command3_Click()
Label2.Visible = False
Label3.Visible = False
Command2.Visible = False
Command3.Visible = False
Text1.Visible = False
Text2.Visible = False
Picture2.Visible = False
Command1.Visible = True
Label1.Visible = True
Radio1.Visible = False
Radio2.Visible = False
Text1.Text = ""
Text2.Text = ""
h = ""
m = ""
tm = ""
th = ""
End Sub



Private Sub Form_Load()
Picture2.Visible = False
Label2.Visible = False
Label3.Visible = False
Command2.Visible = False
Command3.Visible = False
Text1.Visible = False
Text2.Visible = False
Label1.Visible = True
End Sub



