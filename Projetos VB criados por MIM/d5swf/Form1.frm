VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Incúria Screensaver"
   ClientHeight    =   11520
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   15360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11085
      TabIndex        =   1
      Top             =   30
      Width           =   765
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   8805
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   11820
      _cx             =   20849
      _cy             =   15531
      FlashVars       =   ""
      Movie           =   "c:\incuria.swf"
      Src             =   "c:\incuria.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   "c:\incuria.swf"
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   -1  'True
      BGColor         =   "000000"
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
End
End Sub

Private Sub Form_Load()

ShockwaveFlash1.Play
ShockwaveFlash1.LoadMovie 0, "c:\sp.swf"
'ShockwaveFlash1.Left = Me.Width / 2 - ShockwaveFlash1.Width / 2
'ShockwaveFlash1.Top = Me.Height / 2 - ShockwaveFlash1.Height / 2
ShockwaveFlash1.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200

Command1.Left = Me.Width - Command1.Width
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static cnt As Integer
If cnt > 3 Then
End
Else
cnt = cnt + 1
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
End
End Sub

