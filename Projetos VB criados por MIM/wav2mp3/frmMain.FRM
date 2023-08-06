VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JR MID2WAV"
   ClientHeight    =   1212
   ClientLeft      =   156
   ClientTop       =   156
   ClientWidth     =   6312
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1212
   ScaleWidth      =   6312
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   720
      Left            =   48
      TabIndex        =   6
      Top             =   2472
      Width           =   6192
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   312
         Top             =   216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CONVERT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   2088
         TabIndex        =   7
         Top             =   264
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "File to Save:"
      Height          =   1152
      Left            =   48
      TabIndex        =   3
      Top             =   1272
      Width           =   6192
      Begin VB.CheckBox Check1 
         Caption         =   "Convert in MP3"
         Height          =   324
         Left            =   1488
         TabIndex        =   8
         Top             =   312
         Width           =   1572
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save Where..."
         Height          =   300
         Left            =   96
         TabIndex        =   5
         Top             =   312
         Width           =   1284
      End
      Begin VB.TextBox txtSalva 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   96
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   5964
      End
      Begin MSComDlg.CommonDialog c2 
         Left            =   5928
         Top             =   768
         _ExtentX        =   677
         _ExtentY        =   677
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File to convert:"
      Height          =   1152
      Left            =   48
      TabIndex        =   1
      Top             =   24
      Width           =   6192
      Begin MSComDlg.CommonDialog c 
         Left            =   5688
         Top             =   288
         _ExtentX        =   677
         _ExtentY        =   677
         _Version        =   393216
      End
      Begin VB.CommandButton btnApri 
         Caption         =   "Open File..."
         Height          =   300
         Left            =   96
         TabIndex        =   0
         Top             =   312
         Width           =   1284
      End
      Begin VB.TextBox filename 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   96
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   5964
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mm As New MusicModule

Private Declare Function mciSendString Lib "winmm.dll" _
                                   Alias "mciSendStringA" _
                                   (ByVal lpstrCommand As String, _
                                   ByVal lpstrReturnString As String, _
                                   ByVal uReturnLength As Long, _
                                   ByVal hwndCallback As Long) As Long


Private Function RecordSound(filename As String) As Boolean
    preservetime = Time
    Dim Result&
    Dim ReturnString As String * 1024
    Result& = mciSendString("open new Type waveaudio Alias recsound", ReturnString, Len(ReturnString), 0) 'Start at the beginning
    Result& = mciSendString("set recsound time format ms bitspersample 16 channels 2 bytespersec 22500 samplespersec 44100", ReturnString, 1024, 0) 'CD Quality Sound
    Result& = mciSendString("record  recsound", ReturnString, Len(ReturnString), 0) 'Start Recording
   
   mm.Play
   Timer1.Enabled = True
   Do
   DoEvents
   x = x + 1
   Loop Until mm.EndOfSong
   
    Result& = mciSendString("save recsound " & txtSalva.text, ReturnString, Len(ReturnString), 0) 'Save recording
    Result& = mciSendString("close recsound", ReturnString, 1024, 0) 'Close/Stop recording
  Time = preservetime
  If Check1.Value = 1 Then frmConvert.Show: Form1.Enabled = False: Exit Function
  
  Frame1.Enabled = True
  Frame2.Enabled = True
  Command1.Enabled = True
  Command1.Caption = "CONVERT"

  End Function
 

Private Sub btnApri_Click()
On Error Resume Next
c.Filter = "Midi Files (*.mid)|*.mid|MP3 Files (*.mp3)|*.mp3|Wave Files (*.wav)|*.wav|All Files (*.*)|*.*"
'c.Filter = "MP3 Files (*.mp3)|*.mp3|Wave Files (*.wav)|*.wav|Midi Files (*.mid)|*.mid|All Files (*.*)|*.*"
c.ShowOpen
If c.filename = "" Then Exit Sub
If c.filename = " " Then Exit Sub
filename.text = c.filename
mm.filename = c.filename
Me.Height = 2850

End Sub

Private Sub Command1_Click()
On Error Resume Next
Frame1.Enabled = False
Frame2.Enabled = False
Command1.Enabled = False
Command1.Caption = "Please wait..."
Call RecordSound(filename.text)


End Sub

Private Sub Command2_Click()
On Error Resume Next
c2.Filter = "Wave Files (*.wav)|*.wav"
c2.ShowSave
If c2.filename = "" Then Exit Sub
If c2.filename = " " Then Exit Sub
txtSalva.text = c2.filename
fileWAV = c2.filename
Me.Height = 3600

End Sub

Private Sub Command3_Click()
frmBatch.Show
Unload Me
End Sub

Private Sub Form_Terminate()
On Error Resume Next
mm.StopPlay
End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
mm.StopPlay
End Sub


