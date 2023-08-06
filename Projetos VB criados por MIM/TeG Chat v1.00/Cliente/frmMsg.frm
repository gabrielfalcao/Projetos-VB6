VERSION 5.00
Begin VB.Form frmMSG 
   BackColor       =   &H00EFD1AD&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   9645
   ClientTop       =   7425
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   Picture         =   "frmMsg.frx":0000
   ScaleHeight     =   1500
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   960
      Top             =   915
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00EFD1AD&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   135
      MaxLength       =   275
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmMsg.frx":BBC2
      Top             =   135
      Width           =   2100
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   540
      Top             =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0099372D&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   210
      Left            =   2280
      TabIndex        =   0
      Top             =   -30
      Width           =   135
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer


Private Sub Form_Load()
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height
If Len(Text1.Text) >= 142 Then
Text1.Text = Right$(Text1.Text, 142) & "..."
End If
End Sub

Private Sub Form_LostFocus()
retorna
End Sub

Private Sub Label1_Click()
retorna
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFFF00
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HC0C0FF

End Sub

Private Sub ret_Timer()
Static cnt As Integer
If cnt > Screen.Height Then
cnt = cnt - 100
Me.Top = Screen.Height + cnt
Else
ret.Enabled = False
End If
End Sub

Private Sub Text1_Click()
frmChatCliente.WindowState = 0
retorna
End Sub

Private Sub Timer1_Timer()

If cnt < Screen.Height - Me.Height - 8100 Then
cnt = cnt + 100
Me.Top = Screen.Height - cnt
Else
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub
Private Sub retorna()
Timer1.Enabled = False
frmChatCliente.msg.Value = True
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height
cnt = 0
End Sub

Private Sub Timer2_Timer()
retorna
Timer2.Enabled = False
End Sub
