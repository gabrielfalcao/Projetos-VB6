VERSION 5.00
Begin VB.Form menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kit de MP3 - Home Edition by Gabriel Falcão"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Gravadores de MP3"
      Height          =   1320
      Left            =   3120
      TabIndex        =   2
      Top             =   4800
      Width           =   2340
      Begin VB.CommandButton Command9 
         Caption         =   "Atualização do Easy CD Creator 5.0"
         Height          =   405
         Left            =   240
         TabIndex        =   11
         Top             =   765
         Width           =   1875
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Easy CD Creator 5.0"
         Height          =   405
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Players de MP3"
      Height          =   2280
      Left            =   5760
      TabIndex        =   1
      Top             =   2880
      Width           =   2340
      Begin VB.CommandButton Command10 
         Caption         =   "Music Match Jukeox"
         Height          =   405
         Left            =   240
         TabIndex        =   12
         Top             =   1575
         Width           =   1875
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Real Player 8.0"
         Height          =   405
         Left            =   240
         TabIndex        =   8
         Top             =   1170
         Width           =   1875
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Windows Media Player 7.0"
         Height          =   405
         Left            =   240
         TabIndex        =   7
         Top             =   765
         Width           =   1875
      End
      Begin VB.CommandButton Command4 
         Caption         =   "WinAmp"
         Height          =   405
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Baixadores de MP3"
      Height          =   2115
      Left            =   360
      TabIndex        =   0
      Top             =   2880
      Width           =   2445
      Begin VB.CommandButton Command8 
         Caption         =   "Napster"
         Height          =   405
         Left            =   255
         TabIndex        =   9
         Top             =   1515
         Width           =   1875
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Morpheus"
         Height          =   405
         Left            =   255
         TabIndex        =   5
         Top             =   1110
         Width           =   1875
      End
      Begin VB.CommandButton Command2 
         Caption         =   "WinMX"
         Height          =   405
         Left            =   255
         TabIndex        =   4
         Top             =   705
         Width           =   1875
      End
      Begin VB.CommandButton Command1 
         Caption         =   "KaZaA Media Desktop"
         Height          =   405
         Left            =   255
         TabIndex        =   3
         Top             =   300
         Width           =   1875
      End
   End
   Begin VB.Image about 
      Height          =   2250
      Left            =   360
      Top             =   120
      Width           =   8250
   End
   Begin VB.Image pic2 
      Height          =   375
      Left            =   3600
      Picture         =   "Form1.frx":08CA
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image pic1 
      Height          =   375
      Left            =   3240
      Picture         =   "Form1.frx":3D106
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exe(nome As String)
Call Shell(nome, vbNormalFocus)
End Sub
Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub about_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
about.Picture = pic2.Picture
End Sub

Private Sub about_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
about.Picture = pic1.Picture
End Sub

Private Sub Command1_Click()
exe "Kit\Kazaa.exe"
End Sub

Private Sub Command10_Click()
exe "Kit\MusicMatch.exe"
End Sub

Private Sub Command2_Click()
exe "Kit\winmx260.exe"
End Sub

Private Sub Command3_Click()
exe "Kit\Morpheus.exe"
End Sub

Private Sub Command4_Click()
exe "Kit\Winamp.exe"
End Sub

Private Sub Command5_Click()
exe "Kit\wmp7.exe"
End Sub

Private Sub Command6_Click()
exe "Kit\EasyCDCreator5\setup.exe"
End Sub

Private Sub Command7_Click()
exe "Kit\rp8.exe"
End Sub

Private Sub Command8_Click()
exe "Kit\napster.exe"
End Sub

Private Sub Command9_Click()
exe "Kit\Patch.exe"
End Sub

Private Sub Form_Load()
about.Picture = pic1.Picture
End Sub
