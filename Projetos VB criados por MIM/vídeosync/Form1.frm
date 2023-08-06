VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sincronizador de Audio e Vídeo por Gabriel Falcão"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8580
      TabIndex        =   2
      Top             =   6015
      Width           =   8580
      Begin VB.CommandButton Command1 
         Caption         =   "Selecionar vídeo"
         Height          =   285
         Left            =   1050
         TabIndex        =   6
         Top             =   15
         Width           =   2025
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "PLAY"
         Height          =   285
         Left            =   5250
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   15
         Width           =   915
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Selecionar audio"
         Height          =   285
         Left            =   3075
         TabIndex        =   4
         Top             =   15
         Width           =   2025
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "STOP"
         Height          =   285
         Left            =   6165
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   15
         Width           =   915
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1185
      Top             =   3165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir vídeo"
      Filter          =   "Vídeos|*.mpg;*.mpeg;*.avi;*.wmv"
   End
   Begin MSComDlg.CommonDialog cd2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir som"
      Filter          =   "AUDIO|*.mp3;*wav"
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   6000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8580
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   50
      mute            =   -1  'True
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   15134
      _cy             =   10583
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer2 
      Height          =   30
      Left            =   60
      TabIndex        =   1
      Top             =   6255
      Width           =   8475
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   14949
      _cy             =   53
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cd1.ShowOpen
WindowsMediaPlayer1.URL = cd1.FileName
WindowsMediaPlayer1.Controls.stop
WindowsMediaPlayer1.settings.volume = 0
End Sub

Private Sub Command2_Click()
HScroll1.Max = WindowsMediaPlayer1.currentMedia.duration
WindowsMediaPlayer1.Controls.play
WindowsMediaPlayer2.Controls.play
End Sub

Private Sub Command3_Click()
cd2.ShowOpen
WindowsMediaPlayer2.URL = cd2.FileName
WindowsMediaPlayer2.Controls.stop
End Sub

Private Sub Command4_Click()
WindowsMediaPlayer1.Controls.stop
WindowsMediaPlayer2.Controls.stop
End Sub

Private Sub HScroll1_Change()
WindowsMediaPlayer1.currentMedia.getMarkerTime HScroll1.Value
WindowsMediaPlayer2.currentMedia.getMarkerTime HScroll1.Value
End Sub
