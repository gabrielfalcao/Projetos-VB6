VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H00000000&
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ControlBox      =   0   'False
   Icon            =   "neoass.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   165
      TabIndex        =   0
      Top             =   -165
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox Picture1 
      Height          =   4290
      Left            =   3833
      Picture         =   "neoass.frx":0442
      ScaleHeight     =   4230
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   3555
      Width           =   7575
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1350
         TabIndex        =   3
         Top             =   3450
         Width           =   1275
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H001B1FC5&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3015
         Width           =   1770
      End
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   420
         Picture         =   "neoass.frx":6827
         Stretch         =   -1  'True
         Top             =   1410
         Width           =   2940
      End
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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
                  Dim nero As String
    nero = "C:\Arquivos de programas\ahead\Nero\nero.exe"
    Call Shell(nero, 1)
    
        Dim Ret  As Long
        Dim pOld As Boolean
        Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
Unload Me
    Else
        rtn = FindWindow("Shell_traywnd", "")
        Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
        ShowCursor (True)

        Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
        Unload Me
        End
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

Private Sub Image2_Click()

End Sub
