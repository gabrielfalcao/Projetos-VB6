VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversione Batch"
   ClientHeight    =   5640
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   3660
   Icon            =   "frmBatch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   3660
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSalva 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3672
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Converti in MP3"
      Height          =   324
      Left            =   1032
      TabIndex        =   7
      Top             =   4800
      Width           =   1572
   End
   Begin VB.Frame Frame3 
      Height          =   972
      Left            =   96
      TabIndex        =   5
      Top             =   4608
      Width           =   3504
      Begin VB.CommandButton Command3 
         Caption         =   "CONVERTI"
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
         Left            =   720
         TabIndex        =   6
         Top             =   576
         Width           =   1935
      End
   End
   Begin VB.TextBox filename 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3696
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   264
      Visible         =   0   'False
      Width           =   2316
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleziona file:"
      Height          =   4428
      Left            =   96
      TabIndex        =   0
      Top             =   96
      Width           =   3492
      Begin MSComDlg.CommonDialog c1 
         Left            =   3048
         Top             =   1512
         _ExtentX        =   677
         _ExtentY        =   677
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aggiungi file...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   264
         Width           =   3228
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancella contenuto"
         Height          =   276
         Left            =   144
         TabIndex        =   2
         Top             =   4056
         Width           =   3228
      End
      Begin VB.ListBox List1 
         Height          =   3312
         Left            =   120
         TabIndex        =   1
         Top             =   768
         Width           =   3228
      End
   End
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.Clear
End Sub

Private Sub Command2_Click()
With c1
    .filename = ""
    .DialogTitle = "Apri"
    .CancelError = False
    .Filter = "Tutti i midi (*.mid)|*.mid"
    .ShowOpen
    If Len(.filename) = 0 Then
         Exit Sub
    End If
List1.AddItem .filename
End With

End Sub

Private Sub Command3_Click()
filename.text = List1.List(0)
txtSalva.text = List1.List(0) & ".wav"

On Error Resume Next
Frame1.Enabled = False
Frame2.Enabled = False
Command3.Enabled = False
Command3.Caption = "Attendere.."
Call RecordSound(filename.text)

End Sub


Private Function RecordSound(filename As String) As Boolean
    preservetime = Time
    Dim Result&
    Dim ReturnString As String * 1024
    
For v = 0 To List1.ListCount
filename.text = List1.List(v)
txtSalva.text = List1.List(v) & ".wav"
    
    Result& = mciSendString("open new Type waveaudio Alias recsound", ReturnString, Len(ReturnString), 0) 'Start at the beginning
    Result& = mciSendString("set recsound time format ms bitspersample 16 channels 2 bytespersec 22500 samplespersec 44100", ReturnString, 1024, 0) 'CD Quality Sound
    Result& = mciSendString("record  recsound", ReturnString, Len(ReturnString), 0) 'Start Recording
   
   mm.Play
   Do
   DoEvents
   x = x + 1
   Loop Until mm.EndOfSong
   
    Result& = mciSendString("save recsound " & txtSalva.text, ReturnString, Len(ReturnString), 0) 'Save recording
    'Result& = mciSendString("save recsound " & filename, ReturnString, Len(ReturnString), 0) 'Save recording
    Result& = mciSendString("close recsound", ReturnString, 1024, 0) 'Close/Stop recording
  Time = preservetime
  If Check1.Value = 1 Then frmConvert2.Show: Form1.Enabled = False
Next v
  
  Frame1.Enabled = True
  Frame2.Enabled = True
  Command3.Enabled = True
  Command3.Caption = "CONVERTI"

End Function

