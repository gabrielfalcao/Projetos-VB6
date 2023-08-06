VERSION 5.00
Begin VB.Form imgmovie 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   180
   ClientTop       =   0
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3750
      Left            =   2205
      ScaleHeight     =   3720
      ScaleWidth      =   5490
      TabIndex        =   9
      Top             =   1440
      Width           =   5520
   End
   Begin VB.HScrollBar movTimer 
      Height          =   255
      Left            =   3165
      TabIndex        =   7
      Top             =   5550
      Width           =   4605
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFD1AD&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   7845
      ScaleHeight     =   375
      ScaleWidth      =   960
      TabIndex        =   5
      Top             =   4200
      Width           =   990
      Begin VB.CheckBox chkstretch 
         BackColor       =   &H00EFD1AD&
         Caption         =   "Esticar"
         Height          =   195
         Left            =   75
         TabIndex        =   6
         Top             =   75
         Width           =   885
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   645
      Left            =   4800
      MousePointer    =   2  'Cross
      TabIndex        =   8
      Top             =   330
      Width           =   1380
   End
   Begin VB.Label cmdpause 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pausar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   8122
      TabIndex        =   4
      Top             =   3765
      Width           =   420
   End
   Begin VB.Label cmdrewind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rebobinar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   8025
      TabIndex        =   3
      Top             =   3210
      Width           =   615
   End
   Begin VB.Label cmdstop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   8167
      TabIndex        =   2
      Top             =   2250
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5385
      TabIndex        =   1
      Top             =   510
      Width           =   585
   End
   Begin VB.Label cmdplay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tocar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   8160
      TabIndex        =   0
      Top             =   1680
      Width           =   345
   End
End
Attribute VB_Name = "imgmovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Example code for the CAVI Class
'
' To try this example, do the following:
' 1. Create a new form
' 2. Add a command button called 'cmdPlay'
' 3. Add a command button called 'cmdStop'
' 4. Add a command button called 'cmdPause'
' 5. Add a command button called 'cmdRewind'
' 6. Add a check box called 'chkStretch'
' 7. Add an image control called 'ImgMovie'
' 8. Paste all the code from this example to the new form's module
' 9. Run the form

' This example assumes that the sample files are located in the
' directory named by the following constant.
Private Const mcstrExamplePath = "C:\OpenFX\MOVIES"
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim m_Rgn As CBMPRegion
Private mCaptionlessWindowMover As CCaptionlessWindowMover
Private mavi As CAVI

Private Sub chkStretch_Click()
  mavi.Stretch = (chkstretch = 1)
End Sub

Private Sub cmdPause_Click()
  mavi.Pause
End Sub

Private Sub cmd_Click(Index As Integer)

End Sub

Private Sub cmdPlay_Click()
  mavi.Play
End Sub

Private Sub cmdRewind_Click()
  mavi.Rewind
End Sub

Private Sub cmdStop_Click()
  mavi.StopPlay
End Sub

Private Sub Form_Load()
  Set mavi = New CAVI
Set m_Rgn = New CBMPRegion
Set mCaptionlessWindowMover = New CCaptionlessWindowMover
  Set mCaptionlessWindowMover.Form = Me
  m_Rgn.CreateFromPic Me.Picture, vbBlack
  SetWindowRgn hwnd, m_Rgn.Handle, True
  cmdplay.Caption = "Play"
  cmdstop.Caption = "Stop"
  cmdpause.Caption = "Pause"
  cmdrewind.Caption = "Rewind"
  chkstretch.Caption = "Stretch"

  ' position movie
  mavi.Left = imgmovie.Left / Screen.TwipsPerPixelX
  mavi.Top = imgmovie.Top / Screen.TwipsPerPixelX
  mavi.Width = imgmovie.Width / Screen.TwipsPerPixelX
  mavi.Height = imgmovie.Height / Screen.TwipsPerPixelX
  movTimer.Max = mavi.Length
  movTimer.Min = 1
  ' Specify file
  mavi.FileName = "C:\OpenFX\MOVIES\unnamed.avi"
  
  ' Set the parent of the movie window
  mavi.hWndParent = Picture2.hwnd
  
  ' Open movie
  mavi.OpenAVI
  
End Sub

Private Sub form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Handle the form's MouseDown event
  mCaptionlessWindowMover.HandleMouseDown X, Y
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Handle the form's MouseMove event
  mCaptionlessWindowMover.HandleMouseMove X, Y
End Sub

Private Sub form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Handle the form's MouseUp event
  mCaptionlessWindowMover.HandleMouseUp
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mavi = Nothing
    SetWindowRgn hwnd, 0, False
  m_Rgn.Destroy
  Set m_Rgn = Nothing
End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub movTimer_Change()
mavi.Position = movTimer.Value
End Sub
