VERSION 5.00
Begin VB.Form ModSenha 
   BackColor       =   &H00000000&
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4965
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   ScaleHeight     =   5445
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox i1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5235
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   4920
      Begin VB.TextBox Text2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001EDEE8&
         Height          =   2295
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "ModSenha.frx":0000
         Top             =   1035
         Width           =   4620
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Não exibir esta tela novamente"
         ForeColor       =   &H001EDEE8&
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   4245
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H001EDEE8&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   22
         TabIndex        =   2
         Text            =   "12345"
         Top             =   3885
         Width           =   4785
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3930
         TabIndex        =   1
         Top             =   4245
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ATENÇÃO"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   825
         Left            =   900
         TabIndex        =   9
         Top             =   90
         Width           =   3450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANCELAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1650
         TabIndex        =   8
         Top             =   4755
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Definir Senha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001EDEE8&
         Height          =   330
         Left            =   1380
         TabIndex        =   6
         Top             =   3405
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   37
         TabIndex        =   5
         Top             =   1189
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   1747
         TabIndex        =   4
         Top             =   1189
         Visible         =   0   'False
         Width           =   90
      End
   End
End
Attribute VB_Name = "ModSenha"
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
Private Sub Check1_Click()
If Check1.Value = 0 Then
Label3.Caption = "0"
Else
Label3.Caption = "1"
End If
End Sub

Private Sub Command1_Click()
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\trvtd"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\trvtd", "PTAESSST", (Text1.Text)
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\trvtd", "CHKval", (Label3.Caption)
frmPassword.Show
ModSenha.Hide
End Sub

Private Sub Form_Load()
Label3.Caption = Check1.Value
Label2.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\trvtd", "CHKval")
If Label2.Caption = "1" Then
frmPassword.Show
ModSenha.Hide
Else
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF&
End Sub

Private Sub Form_Resize()
i1.Left = Me.Width / 2 - i1.Width / 2
i1.Top = Me.Height / 2 - i1.Height / 2
End Sub

Private Sub i1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF&
End Sub

Private Sub Label4_Click()
        rtn = FindWindow("Shell_traywnd", "")
        Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
        ShowCursor (True)
        Load frmCorrect
        Dim Ret  As Long
        Dim pOld As Boolean
        Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
      End
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HC0C0FF
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H1EDEE8
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF&
End Sub
