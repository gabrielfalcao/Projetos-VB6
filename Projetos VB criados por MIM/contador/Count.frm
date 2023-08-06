VERSION 5.00
Begin VB.Form Count 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   990
   ClientLeft      =   3090
   ClientTop       =   2460
   ClientWidth     =   13875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   990
   ScaleWidth      =   13875
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   915
      Top             =   690
   End
   Begin VB.PictureBox pc 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   13815
      TabIndex        =   4
      Top             =   0
      Width           =   13875
   End
   Begin VB.CommandButton Count 
      Caption         =   "&Ir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5430
      TabIndex        =   0
      Top             =   589
      Width           =   975
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   255
      TabIndex        =   3
      Top             =   480
      Width           =   4020
   End
   Begin VB.Label Label2 
      Height          =   315
      Left            =   8010
      TabIndex        =   2
      Top             =   615
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   555
      Width           =   105
   End
End
Attribute VB_Name = "Count"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Private Sub Count_Click()
    Label3.Caption = Time
Timer1.Enabled = True
End Sub

Private Sub ProgressBar1_Click()
Unload Me
End Sub
Private Function pbar(pb As Control, ByVal Percent As Integer, Optional ByVal ShowPercent = False)
    'Replacement for progress bar..looks nicer also
    Dim sNum                            As String    'use percent
    Dim Num$
    If Not pb.AutoRedraw Then 'picture in memory ?
        pb.AutoRedraw = -1 'no, make one
    End If
    pb.Cls 'clear picture in memory
    pb.ScaleWidth = 100 'new sclaemodus
    pb.DrawMode = 10 'not XOR Pen Modus
    If ShowPercent = True Then
    Num$ = Format$(Percent, "###0") + "%"
    pb.CurrentX = 50 - pb.TextWidth(Num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(Num$)) / 2
    pb.Print Num$ 'print percent
    End If
    pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
    pb.Refresh 'show differents
End Function

Private Sub Form_Load()
cnt = 0
End Sub

Private Sub Timer1_Timer()

If cnt <= 100 Then

   pbar pc, cnt, True
cnt = cnt + 1
End If

End Sub
