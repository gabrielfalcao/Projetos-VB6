VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form pl 
   Caption         =   "Media Player"
   ClientHeight    =   3195
   ClientLeft      =   5550
   ClientTop       =   1755
   ClientWidth     =   4680
   Icon            =   "pl.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin WMPLibCtl.WindowsMediaPlayer mp1 
      Height          =   3090
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8070
      _cy             =   5450
   End
End
Attribute VB_Name = "pl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload cn
End Sub

Private Sub Form_Resize()
mp1.Width = Me.Width - 170
mp1.Height = Me.Height - 500
End Sub
