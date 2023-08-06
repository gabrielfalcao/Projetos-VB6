VERSION 5.00
Begin VB.Form frmDesligator 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3255
   ClientLeft      =   735
   ClientTop       =   450
   ClientWidth     =   6765
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":08CA
   ScaleHeight     =   3255
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin poweroff.isButton Command1 
      Height          =   330
      Left            =   4830
      TabIndex        =   8
      Top             =   1260
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   582
      Icon            =   "Form3.frx":48678
      Style           =   8
      Caption         =   "PROGRAMAR"
      IconAlign       =   1
      iNonThemeStyle  =   1
      USeCustomColors =   -1  'True
      BackColor       =   6886418
      HighlightColor  =   12632064
      FontColor       =   10082551
      FontHighlightColor=   8388608
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0099D8F7&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3780
      TabIndex        =   6
      Text            =   "00:00:00"
      Top             =   1260
      Width           =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   60
      Left            =   2385
      Top             =   2490
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2940
      Top             =   2610
   End
   Begin poweroff.isButton Command3 
      Height          =   315
      Left            =   4770
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Icon            =   "Form3.frx":48694
      Style           =   8
      Caption         =   "Minimizar"
      IconAlign       =   1
      iNonThemeStyle  =   8
      USeCustomColors =   -1  'True
      HighlightColor  =   32767
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin poweroff.isButton Command2 
      Height          =   315
      Left            =   5865
      TabIndex        =   10
      Top             =   60
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Icon            =   "Form3.frx":486B0
      Style           =   8
      Caption         =   "Fechar"
      IconAlign       =   1
      iNonThemeStyle  =   8
      USeCustomColors =   -1  'True
      HighlightColor  =   21759
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   390
      Left            =   75
      TabIndex        =   7
      Top             =   45
      Width           =   435
   End
   Begin VB.Image imgminn 
      Height          =   30
      Left            =   6030
      Picture         =   "Form3.frx":486CC
      Stretch         =   -1  'True
      Top             =   675
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.megaaccesshp.hpg.com.br"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4065
      MouseIcon       =   "Form3.frx":48F7E
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2715
      Width           =   2325
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gabrielfalcao@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4065
      MouseIcon       =   "Form3.frx":49288
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2520
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gabriel Falcão"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4065
      TabIndex        =   3
      Top             =   2325
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programação e Design:"
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
      Height          =   195
      Left            =   3420
      TabIndex        =   2
      Top             =   2070
      Width           =   1995
   End
   Begin VB.Image imgMin 
      Height          =   15
      Left            =   5850
      Picture         =   "Form3.frx":49592
      Stretch         =   -1  'True
      Top             =   855
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgEnd 
      Height          =   15
      Left            =   5850
      Picture         =   "Form3.frx":49E44
      Stretch         =   -1  'True
      Top             =   855
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0099D8F7&
      Height          =   375
      Left            =   1725
      TabIndex        =   1
      Top             =   1020
      Width           =   135
   End
   Begin VB.Label lbltime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0099D8F7&
      Height          =   375
      Left            =   1725
      TabIndex        =   0
      Top             =   1575
      Width           =   135
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuMin 
         Caption         =   "Minimizar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmDesligator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prgmd As Boolean
Dim poweroff As String
Dim nid As NOTIFYICONDATA
Private Type LUID
         UsedPart As Long
         IgnoredForNowHigh32BitPart As Long
      End Type

      Private Type TOKEN_PRIVILEGES
        PrivilegeCount As Long
        TheLuid As LUID
        Attributes As Long
      End Type
Private mCaptionlessWindowMover As CCaptionlessWindowMover

Private Const SW_SHOW = 5
      Private Const EWX_SHUTDOWN As Long = 1
      Private Const EWX_FORCE As Long = 4
      Private Const EWX_REBOOT = 2
      Private Const EWX_POWEROFF As Long = 8

      Private Declare Function ExitWindowsEx Lib "user32" (ByVal _
           dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
      Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
      Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
         ProcessHandle As Long, _
         ByVal DesiredAccess As Long, TokenHandle As Long) As Long
      Private Declare Function LookupPrivilegeValue Lib "advapi32" _
         Alias "LookupPrivilegeValueA" _
         (ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
         As LUID) As Long
      Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
         (ByVal TokenHandle As Long, _
         ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
         , ByVal BufferLength As Long, _
      PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
      
Private Sub AdjustToken()
         Const TOKEN_ADJUST_PRIVILEGES = &H20
         Const TOKEN_QUERY = &H8
         Const SE_PRIVILEGE_ENABLED = &H2
         Dim hdlProcessHandle As Long
         Dim hdlTokenHandle As Long
         Dim tmpLuid As LUID
         Dim tkp As TOKEN_PRIVILEGES
         Dim tkpNewButIgnored As TOKEN_PRIVILEGES
         Dim lBufferNeeded As Long

         hdlProcessHandle = GetCurrentProcess()
         OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
            TOKEN_QUERY), hdlTokenHandle
         LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

         tkp.PrivilegeCount = 1
         tkp.TheLuid = tmpLuid
         tkp.Attributes = SE_PRIVILEGE_ENABLED
         AdjustTokenPrivileges hdlTokenHandle, False, _
         tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
     End Sub

'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.


Private Sub Command1_Click()
If MsgBox("Programação definida para desligar às " & Text1.Text & ", confirma a programação?", vbQuestion + vbYesNo, poweroff) = vbYes Then
Text1.Enabled = False
Timer1.Enabled = True
Command1.Enabled = False

Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height - 420 - 420
Me.Hide
Command3.Visible = True
mnuMin.Visible = True
   nid.cbSize = Len(nid)
   nid.hwnd = frmDesligator.hwnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = frmDesligator.Icon
   nid.szTip = "Clique duplo aqui para mostrar o programa..." & vbNullChar
   Shell_NotifyIcon NIM_ADD, nid
   MsgBox "O Programa está ao lado do relógio"

Text1.BackColor = &HFF&
Text1.ForeColor = &HFFFF&
Command1.Caption = "PROGRAMADO"
prgmd = True
End If
End Sub

Private Sub Command2_Click()
If prgmd = True Then
Dim MSG As Integer
MSG = MsgBox("Há um desligamento programado, deseja sair mesmo assim?", vbYesNo + vbExclamation, "Desligator 1.0 by Gabriel Falcão")
Select Case MSG
Case vbYes
Unload Me
Case vbNo
Exit Sub
Case Else
End Select
Else
If MsgBox("Deseja realmente sair do programa?", vbQuestion + vbYesNo, poweroff) = vbYes Then Unload Me
End If
End Sub



Private Sub Command3_Click()
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height - 420
Me.Hide
End Sub


Private Sub Form_Load()

poweroff = "PowerOff XP"
 Set mCaptionlessWindowMover = New CCaptionlessWindowMover
  Set mCaptionlessWindowMover.Form = Me
Text1.Text = time
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height - 420
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
mCaptionlessWindowMover.HandleMouseDown x, Y
End Sub

Private Sub Form_MouseMove _
   (Button As Integer, _
    Shift As Integer, _
    x As Single, _
    Y As Single)
    mCaptionlessWindowMover.HandleMouseMove x, Y
    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Label4.ForeColor = &HFFFFFF
    Label3.ForeColor = &HFFFFFF
 
    Dim MSG As Long
    Dim sFilter As String
    MSG = x / Screen.TwipsPerPixelX
    Select Case MSG
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
Me.Show

       Case WM_RBUTTONDOWN
Me.Show

       Case WM_RBUTTONUP
       Case WM_RBUTTONDBLCLK
    End Select
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
mCaptionlessWindowMover.HandleMouseUp
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Label3_Click()
  ShellExecute GetActiveWindow(), "Open", "mailto:gabrielfalcao@hotmail.com", "", 0&, 1
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = &H0&

 Label4.ForeColor = &HFFFF00
End Sub

Private Sub Label4_Click()
ShellExecute GetActiveWindow(), "Open", "http://www.megaaccesshp.hpg.com.br/main.htm", "", 0&, 1
End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label4.ForeColor = &H0&
    Label3.ForeColor = &HFFFFFF
End Sub

Private Sub Label5_Click()
PopupMenu mnuMain
End Sub

Private Sub mnuClose_Click()
Command2_Click
End Sub

Private Sub mnuMin_Click()
Command3_Click
End Sub

Private Sub Timer1_Timer()
If time = Text1.Text Then
Timer2.Enabled = False
Timer1.Enabled = False
AdjustToken
         ExitWindowsEx (EWX_POWEROFF), &HFFFF
         Unload Me
End If
End Sub

Private Sub Timer2_Timer()
lbltime.Caption = time
lbldate.Caption = Date
End Sub
