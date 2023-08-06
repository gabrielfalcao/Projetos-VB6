VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E2BAA9&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00E2BAA9&
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   60
      MaxLength       =   275
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   2160
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   540
      Top             =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0099372D&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   240
      Left            =   2235
      TabIndex        =   0
      Top             =   15
      Width           =   165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFFF00
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HC0C0FF

End Sub

Private Sub Text1_Click()
MsgBox Len(Text1.Text)
End Sub

Private Sub Timer1_Timer()
Static cnt As Integer
If cnt < Screen.Height - Me.Height - 8100 Then
cnt = cnt + 100
Me.Top = Screen.Height - cnt
Else
Timer1.Enabled = False
End If
End Sub
