VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   ControlBox      =   0   'False
   FillColor       =   &H00FFCCFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   5  'Size
   Picture         =   "Form1.frx":0E42
   ScaleHeight     =   5700
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   1755
      ScaleHeight     =   225
      ScaleWidth      =   1830
      TabIndex        =   2
      Top             =   2700
      Width           =   1890
      Begin VB.PictureBox status 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontTransparent =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   0
         MousePointer    =   1  'Arrow
         ScaleHeight     =   225
         ScaleWidth      =   1845
         TabIndex        =   3
         Top             =   0
         Width           =   1845
      End
   End
   Begin VB.Timer prog 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   60
      Top             =   60
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1515
      Left            =   1650
      MouseIcon       =   "Form1.frx":56424
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Be the crack!"
      Top             =   915
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   900
      Left            =   3960
      MouseIcon       =   "Form1.frx":5672E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Unload ME!"
      Top             =   4800
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim i
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim m_Rgn As CBMPRegion
Private mCaptionlessWindowMover As CCaptionlessWindowMover

Private Sub ctr_Click(Index As Integer)
Select Case Index
Case Is = 0
End
Case Is = 1
Me.BorderStyle = 3
Me.WindowState = 1
End Select
End Sub




Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseDown X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseMove X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseUp
End Sub

Private Sub Label1_Click()
Unload Me
End Sub
Private Sub Form_Load()
Set m_Rgn = New CBMPRegion
Set mCaptionlessWindowMover = New CCaptionlessWindowMover
  Set mCaptionlessWindowMover.Form = Me
  m_Rgn.CreateFromPic Me.Picture, vbWhite
  SetWindowRgn hwnd, m_Rgn.Handle, True
  status.Cls
  status.Print "Ready"
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
Me.BorderStyle = 3
Else
Me.BorderStyle = 1
End If
End Sub
Public Function progress(pb As Control, ByVal Percent As Integer, Optional ByVal ShowPercent = False)
    'Replacement for progress bar..looks nicer also
    Dim sNum                            As String    'use percent
    'Dim Num$
    If Not pb.AutoRedraw Then 'picture in memory ?
        pb.AutoRedraw = -1 'no, make one
    End If
    pb.Cls 'clear picture in memory
    pb.ScaleWidth = 100 'new sclaemodus
    pb.DrawMode = 10 'not XOR Pen Modus
    pb.Print "Cracking wait..."
    If ShowPercent = True Then
    num$ = Format$(Percent, "###0") + "%"
    pb.CurrentX = 50 - pb.TextWidth(num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(num$)) / 2
    pb.Print num$ 'print percent
    End If
    pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
    pb.Refresh 'show differents
End Function
Private Sub Form_Unload(Cancel As Integer)
  SetWindowRgn hwnd, 0, False
  m_Rgn.Destroy
  Set m_Rgn = Nothing
End Sub



Private Sub Label2_Click()

End Sub










Private Sub Label4_Click()
status.Cls
i = 0
prog.Enabled = True
End Sub

Private Sub prog_Timer()

If i < 100 Then
progress status, i, False
i = i + 1
Else
prog.Enabled = False
progress status, 100, False
status.Cls
status.Print "Cracked Succefully!"
End If
End Sub
