VERSION 5.00
Begin VB.Form escFun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "             Escolher Fundo"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   3315
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2505
      TabIndex        =   5
      Top             =   2535
      Width           =   645
   End
   Begin VB.Timer scanner 
      Interval        =   1
      Left            =   4755
      Top             =   1890
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fundos"
      Height          =   720
      Left            =   150
      TabIndex        =   0
      Top             =   2415
      Width           =   2325
      Begin VB.OptionButton f4 
         Caption         =   "4"
         Height          =   315
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   255
         Width           =   540
      End
      Begin VB.OptionButton f3 
         Caption         =   "3"
         Height          =   315
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   255
         Width           =   540
      End
      Begin VB.OptionButton f2 
         Caption         =   "2"
         Height          =   315
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   255
         Width           =   540
      End
      Begin VB.OptionButton f1 
         Caption         =   "1"
         Height          =   315
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   255
         Value           =   -1  'True
         Width           =   540
      End
   End
   Begin VB.Image i1 
      Height          =   585
      Left            =   210
      Picture         =   "escFun.frx":0000
      Stretch         =   -1  'True
      Top             =   255
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image i3 
      Height          =   585
      Left            =   210
      Picture         =   "escFun.frx":2132A
      Stretch         =   -1  'True
      Top             =   1425
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image i2 
      Height          =   585
      Left            =   210
      Picture         =   "escFun.frx":3ED94
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image i4 
      Height          =   585
      Left            =   210
      Picture         =   "escFun.frx":47A5C
      Stretch         =   -1  'True
      Top             =   2010
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Left            =   120
      Stretch         =   -1  'True
      Top             =   60
      Width           =   3075
   End
End
Attribute VB_Name = "escFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
duelo.Image2.Picture = Image1.Picture
Unload Me
End Sub

Private Sub Form_Load()
Image1.Picture = duelo.Image2.Picture
End Sub

Private Sub scanner_Timer()
If f1.Value = True Then
Image1.Picture = i1.Picture
End If
If f2.Value = True Then
Image1.Picture = i2.Picture
End If
If f3.Value = True Then
Image1.Picture = i3.Picture
End If
If f4.Value = True Then
Image1.Picture = i4.Picture
End If

End Sub
