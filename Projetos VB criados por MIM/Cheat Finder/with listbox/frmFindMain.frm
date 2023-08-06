VERSION 5.00
Begin VB.Form frmFindMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheat Finder 1.0"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4050
   Icon            =   "frmFindMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1980
      Left            =   180
      ScaleHeight     =   1980
      ScaleWidth      =   3750
      TabIndex        =   2
      Top             =   345
      Visible         =   0   'False
      Width           =   3750
      Begin VB.ListBox c2 
         Height          =   1230
         ItemData        =   "frmFindMain.frx":08CA
         Left            =   1815
         List            =   "frmFindMain.frx":08D1
         TabIndex        =   5
         Top             =   0
         Width           =   1815
      End
      Begin VB.ListBox c1 
         Height          =   1230
         ItemData        =   "frmFindMain.frx":08EF
         Left            =   0
         List            =   "frmFindMain.frx":08F6
         TabIndex        =   4
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   15
         TabIndex        =   3
         Top             =   1230
         Width           =   3630
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sobre o Cheat Finder 1.0"
      Height          =   300
      Left            =   2010
      TabIndex        =   1
      Top             =   15
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procurar Dicas"
      Height          =   300
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   1995
   End
End
Attribute VB_Name = "frmFindMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hpdebusca As String
Dim consesc As String
Dim letra As String
Dim console As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim num As Long
Dim q$
Dim lastN As Long
Dim N As Long
Dim W$
Dim Msg$
Dim Result
Dim canExit As Boolean
Private Sub c1_Change()
If c1.Text = "PC" Then
consesc = "console7"
End If
If c1.Text = "Playstation" Then
consesc = "console8"
End If
If c1.Text = "Playstation 2" Then
consesc = "console9"
End If
If c1.Text = "Game Boy Color" Then
consesc = "console2"
End If
If c1.Text = "Game Boy Advance" Then
consesc = "console3"
End If
If c1.Text = "XBOX" Then
consesc = "console10"
End If
If c1.Text = "Nintendo 64" Then
consesc = "console6"
End If
If c1.Text = "Super Nintendo" Then
consesc = "console12"
End If
If c1.Text = "Sega Saturn" Then
consesc = "console13"
End If
If c1.Text = "Game Cube" Then
consesc = "console5"
End If
If c1.Text = "Dreamcast" Then
consesc = "console1"
End If
If c1.Text = "Mega Drive" Then
consesc = "console11"
End If
End Sub

Private Sub c1_Click()
If c1.Text = "PC" Then
consesc = "console7"
End If
If c1.Text = "Playstation" Then
consesc = "console8"
End If
If c1.Text = "Playstation 2" Then
consesc = "console9"
End If
If c1.Text = "Game Boy Color" Then
consesc = "console2"
End If
If c1.Text = "Game Boy Advance" Then
consesc = "console3"
End If
If c1.Text = "XBOX" Then
consesc = "console10"
End If
If c1.Text = "Nintendo 64" Then
consesc = "console6"
End If
If c1.Text = "Super Nintendo" Then
consesc = "console12"
End If
If c1.Text = "Sega Saturn" Then
consesc = "console13"
End If
If c1.Text = "Game Cube" Then
consesc = "console5"
End If
If c1.Text = "Dreamcast" Then
consesc = "console1"
End If
If c1.Text = "Mega Drive" Then
consesc = "console11"
End If
End Sub

Private Sub c1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyAscii = 8 Then
c1.Text = "Escolha o Console:"
End If
End Sub

Private Sub c1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
c1.Text = "Escolha o Console:"
End If
End Sub

Private Sub c1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyAscii = 8 Then
c1.Text = "Escolha o Console:"
End If
End Sub

Private Sub c2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyAscii = 8 Then
c2.Text = "Escolha a letra inicial:"
End If
End Sub

Private Sub c2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
c2.Text = "Escolha a letra inicial:"
End If
End Sub

Private Sub c2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyAscii = 8 Then
c2.Text = "Escolha a letra inicial:"
End If
End Sub

Private Sub Command1_Click()
Picture1.Visible = True
End Sub

Private Sub Command2_Click()
frmAboutfinder.Show
End Sub

Private Sub Command3_Click()
If c1.Text = "PC" Then
consesc = "7"
End If
If c1.Text = "Playstation" Then
consesc = "8"
End If
If c1.Text = "Playstation 2" Then
consesc = "9"
End If
If c1.Text = "Game Boy Color" Then
consesc = "2"
End If
If c1.Text = "Game Boy Advance" Then
consesc = "3"
End If
If c1.Text = "XBOX" Then
consesc = "10"
End If
If c1.Text = "Nintendo 64" Then
consesc = "6"
End If
If c1.Text = "Super Nintendo" Then
consesc = "12"
End If
If c1.Text = "Sega Saturn" Then
consesc = "13"
End If
If c1.Text = "Game Cube" Then
consesc = "5"
End If
If c1.Text = "Dreamcast" Then
consesc = "1"
End If
If c1.Text = "Mega Drive" Then
consesc = "11"
End If
If c1.Text = "Escolha o Console:" Then
MsgBox "Escolha um Console!", , Me.Caption
Else
If c2.Text = "Escolha a letra inicial:" Then
MsgBox "Escolha a letra inicial!", , Me.Caption
Else
If c2.Text = "TUDO" Then
frmMain.Show
  frmMain.WebBrowser1.Navigate (hpdebusca & letra & "0" & console & consesc)
Picture1.Visible = False
Else
frmMain.Show
  frmMain.WebBrowser1.Navigate (hpdebusca & letra & c2.Text & console & consesc)
Picture1.Visible = False
End If
End If
End If
End Sub

Private Sub Form_Load()
hpdebusca = "http://www.ddj.com.br/jogos/index.asp?"
letra = "letra="
console = "&console="
c1.AddItem "PC"
c1.AddItem "Playstation"
c1.AddItem "Playstation 2"
c1.AddItem "Game Boy Color"
c1.AddItem "Game Boy Advance"
c1.AddItem "XBOX"
c1.AddItem "Nintendo 64"
c1.AddItem "Super Nintendo"
c1.AddItem "Sega Saturn"
c1.AddItem "Game Cube"
c1.AddItem "Dreamcast"
c1.AddItem "Mega Drive"
c2.AddItem "#"
c2.AddItem "a"
c2.AddItem "b"
c2.AddItem "c"
c2.AddItem "d"
c2.AddItem "e"
c2.AddItem "f"
c2.AddItem "g"
c2.AddItem "h"
c2.AddItem "i"
c2.AddItem "j"
c2.AddItem "k"
c2.AddItem "l"
c2.AddItem "m"
c2.AddItem "n"
c2.AddItem "o"
c2.AddItem "p"
c2.AddItem "q"
c2.AddItem "r"
c2.AddItem "s"
c2.AddItem "t"
c2.AddItem "u"
c2.AddItem "v"
c2.AddItem "w"
c2.AddItem "x"
c2.AddItem "y"
c2.AddItem "z"
c2.AddItem "TUDO"
End Sub
