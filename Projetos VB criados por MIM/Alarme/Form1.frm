VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Despertador / Relógio"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   3090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Tocar"
      Height          =   495
      Left            =   2130
      TabIndex        =   5
      Top             =   2055
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   361
      TabIndex        =   4
      Text            =   "00:00:00"
      Top             =   1260
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ativar alarme"
      Height          =   450
      Left            =   623
      TabIndex        =   1
      Top             =   2688
      Width           =   1830
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   1448
      Top             =   2103
   End
   Begin VB.Label m 
      Height          =   405
      Left            =   2685
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Hora do Alarme:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   68
      TabIndex        =   2
      Top             =   768
      Width           =   2955
   End
   Begin VB.Label relógio 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   361
      TabIndex        =   0
      Top             =   146
      Width           =   2355
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
m.Caption = "ativado"
Me.Hide
Form2.Show
End Sub

Private Sub Command2_Click()
tocarwave ("C:\WINDOWS\MEDIA\DING.WAV")
End Sub

Private Sub Form_Load()
Text1.Text = Time
End Sub

Private Sub Timer1_Timer()
relógio.Caption = Time
If relógio.Caption = Text1.Text Then
Me.Show
Unload Form2
tocarwave ("C:\WINDOWS\MEDIA\DING.WAV")
End If
m.Caption = "desativado"
End Sub
