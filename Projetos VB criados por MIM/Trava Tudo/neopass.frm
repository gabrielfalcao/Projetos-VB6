VERSION 5.00
Begin VB.Form frmPassword 
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ControlBox      =   0   'False
   Icon            =   "neopass.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&Sobre o Trava Tudo"
      Height          =   495
      Left            =   6855
      TabIndex        =   3
      Top             =   5940
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   165
      TabIndex        =   2
      Top             =   -165
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6810
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5070
      Width           =   1770
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   7320
      TabIndex        =   0
      Top             =   5580
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   11520
      Left            =   15
      Picture         =   "neopass.frx":0CCA
      Top             =   -60
      Width           =   15360
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)

Private Const SPI_SCREENSAVERRUNNING = 97
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40

Private Sub cmdAbout_Click()
    
    Load frmAbout
    frmAbout.Show
    
End Sub

Private Sub cmdOk_Click()

    Password = txtPassword.Text
    If Password = Text2.Text Then
        rtn = FindWindow("Shell_traywnd", "")
        Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
        ShowCursor (True)
        Load frmCorrect
        frmCorrect.Show
        Dim Ret  As Long
        Dim pOld As Boolean
        Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)

    Else
        Load frmWrong
        frmWrong.Show
    End If

End Sub

Private Sub Form_Load()

    Dim rtn As Long

    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    
    Dim Ret  As Long
    Dim pOld As Boolean
    Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
    
    ShowCursor (True)
    Text2.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\trvtd", "PTAESSST")
End Sub

Private Sub Form_Resize()
Image2.Width = Me.Width - 15
Image2.Height = Me.Height - 15
Image2.Left = Me.Width / 2 - Image2.Width / 2
Image2.Top = Me.Height / 2 - Image2.Height / 2
'Image1.Left = Image2.Width \ 2 - Image1.Width / 2
txtPassword.Left = Image2.Width \ 2 - txtPassword.Width / 2
cmdAbout.Left = Image2.Width \ 2 - cmdAbout.Width / 2
cmdOk.Left = Image2.Width \ 2 - cmdOk.Width / 2
End Sub

Private Sub Picture1_Click()

End Sub

