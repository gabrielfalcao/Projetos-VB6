VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Menu Tinh@ Program"
   ClientHeight    =   8985
   ClientLeft      =   -30
   ClientTop       =   -30
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "frmMenu.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Image sai 
      Height          =   750
      Left            =   0
      Top             =   1935
      Width           =   2250
   End
   Begin VB.Image iopt 
      Height          =   255
      Left            =   9930
      Picture         =   "frmMenu.frx":240042
      Top             =   6630
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image ishutd 
      Height          =   255
      Left            =   9930
      Picture         =   "frmMenu.frx":2488A4
      Top             =   6630
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image isai 
      Height          =   255
      Left            =   9930
      Picture         =   "frmMenu.frx":24E83E
      Top             =   6630
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image iteg 
      Height          =   255
      Left            =   9930
      Picture         =   "frmMenu.frx":2540C8
      Top             =   6630
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Inaveg 
      Height          =   255
      Left            =   9930
      Picture         =   "frmMenu.frx":259CDA
      Top             =   6630
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image opt 
      Height          =   990
      Left            =   0
      Top             =   4485
      Width           =   2640
   End
   Begin VB.Image shutd 
      Height          =   810
      Left            =   0
      Top             =   2685
      Width           =   2250
   End
   Begin VB.Image teg 
      Height          =   780
      Left            =   0
      Top             =   1155
      Width           =   2250
   End
   Begin VB.Image naveg 
      Height          =   1155
      Left            =   0
      Top             =   0
      Width           =   2250
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub naveg_Click()
frmMain.Show
End Sub

Private Sub naveg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
naveg.Picture = Inaveg.Picture
End Sub

Private Sub naveg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
naveg.Picture = LoadPicture("")
End Sub

Private Sub opt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
opt.Picture = iopt.Picture
End Sub

Private Sub opt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
opt.Picture = LoadPicture("")
End Sub

Private Sub sai_Click()
End
End Sub

Private Sub sai_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
sai.Picture = isai.Picture
End Sub

Private Sub sai_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
sai.Picture = LoadPicture("")
End Sub

Private Sub shutd_Click()
On Error Resume Next
Dim intResp As Integer
intResp = MsgBox("Você Tem certeza que deseja Desligar o Computador ?", vbYesNo)
Select Case intResp
Case vbYes
exe ("RUNDLL.EXE user.exe,exitwindows")
Case Else
End Select

End Sub

Private Sub shutd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shutd.Picture = ishutd.Picture
End Sub

Private Sub shutd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shutd.Picture = LoadPicture("")
End Sub

Private Sub teg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
teg.Picture = iteg.Picture
End Sub

Private Sub teg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
teg.Picture = LoadPicture("")
End Sub
