VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Assistente de Duelo 1.0"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form2"
   ScaleHeight     =   1740
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame d2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1380
      Left            =   2895
      TabIndex        =   4
      Top             =   210
      Width           =   2280
      Begin VB.TextBox pv2s 
         Height          =   360
         Left            =   315
         TabIndex        =   6
         Top             =   795
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&OK"
         Height          =   360
         Left            =   1515
         TabIndex        =   5
         Top             =   795
         Width           =   480
      End
      Begin VB.Label pv2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   315
         TabIndex        =   7
         Top             =   390
         Width           =   1680
      End
   End
   Begin VB.Frame d1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1380
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   2280
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   1740
         Top             =   -210
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   360
         Left            =   1515
         TabIndex        =   3
         Top             =   795
         Width           =   480
      End
      Begin VB.TextBox pv1s 
         Height          =   360
         Left            =   315
         TabIndex        =   2
         Top             =   795
         Width           =   1170
      End
      Begin VB.Label pv1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   315
         TabIndex        =   1
         Top             =   390
         Width           =   1680
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
pv1.Caption = pv1.Caption - pv1s.Text
End Sub

Private Sub Command2_Click()
pv2.Caption = pv2.Caption - pv12s.Text
End Sub

Private Sub Form_Load()
pv1.Caption = 8000
pv2.Caption = 8000
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
If pv1.Caption <= 0 Then
MsgBox "Vencedor: " & d2.Caption & "  Perdedor: " & d1.Caption
Timer1.Enabled = False
Unload Me
End If
If pv2.Caption <= 0 Then
MsgBox "Vencedor: " & d1.Caption & "  Perdedor: " & d2.Caption
Timer1.Enabled = False
Unload Me
End If
End Sub
