VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Osciloscópio"
   ClientHeight    =   1470
   ClientLeft      =   2100
   ClientTop       =   900
   ClientWidth     =   4410
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4410
   Begin Project1.usrMeter usrMeter1 
      Height          =   1470
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   2593
      BackColor       =   128
      FillColor       =   65535
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   465
      Top             =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
usrMeter1.StartMetering 0
End Sub

Private Sub Form_Resize()
usrMeter1.Width = Me.Width - 115
usrMeter1.Height = Me.Height - 115
End Sub

Private Sub Form_Unload(Cancel As Integer)
usrMeter1.StopMetering
End Sub

Private Sub Timer1_Timer()
usrMeter1.Visualize
End Sub
