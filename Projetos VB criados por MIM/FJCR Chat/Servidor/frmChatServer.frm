VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmChatServer 
   BackColor       =   &H00EFD1AD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TeG Chat Lite - Servidor - por Gabriel Falcão - gabrielfalcao@hotmail.com"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   DrawMode        =   12  'Nop
   Icon            =   "frmChatServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Left            =   9750
      Top             =   6615
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3855
      TabIndex        =   8
      Top             =   5730
      Width           =   2745
   End
   Begin VB.TextBox txtPapo 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5535
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   60
      Width           =   9900
   End
   Begin VB.TextBox txtEnviar 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   300
      TabIndex        =   0
      Top             =   6120
      Width           =   8205
   End
   Begin VB.CommandButton cmdEnviar 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Enviar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8505
      TabIndex        =   3
      Top             =   6120
      Width           =   1170
   End
   Begin VB.TextBox txtApelido 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   915
      TabIndex        =   2
      Text            =   "G@gau"
      Top             =   5730
      Width           =   1665
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   9840
      Top             =   6540
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Get Reg"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4605
      TabIndex        =   1
      Top             =   11190
      Width           =   1875
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   6510
      Width           =   720
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   855
      TabIndex        =   6
      Top             =   6495
      Width           =   9105
   End
   Begin VB.Label lblApelido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apelido:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   270
      TabIndex        =   5
      Top             =   5775
      Width           =   585
   End
   Begin VB.Menu mnuchat 
      Caption         =   "chat"
      Visible         =   0   'False
      Begin VB.Menu mnudata 
         Caption         =   "texto aki"
      End
   End
End
Attribute VB_Name = "frmChatServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ss As String
Dim version As String
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim filename As String
Dim linedata As String
  Dim file As String
Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim nid As NOTIFYICONDATA
Private Sub LogErr()
If Err.Number <> 0 Then
Me.Caption = Err.Description
End If
End Sub
Private Sub cmdEnviar_Click()
Dim dados As String
If Socket.State <> sckConnected Then
MsgBox "Você não esta conectado", , Me.Caption
Socket.Close
Socket.Listen
lblInfo.Caption = "Aguardando conexão..."
Else
If txtApelido.Text = "" Then
lblInfo.Caption = "Escolha um apelido!"
txtApelido.SetFocus
Else
If txtEnviar.Text = "" Then
lblInfo.Caption = "Digite algo para enviar!"
txtEnviar.SetFocus
Else
dados = txtApelido.Text & ": " & vbCrLf & txtEnviar.Text
Socket.SendData dados
txtPapo.Text = txtPapo.Text & txtApelido.Text & ": " & vbCrLf & txtEnviar.Text & vbCrLf
txtEnviar.Text = ""
End If
End If
End If
End Sub
Private Sub Form_Load()
Socket.LocalPort = 2502
Text1.Text = Socket.LocalIP
Clipboard.SetText Socket.LocalIP
Socket.Listen
lblInfo.Caption = "Aguardando conexão..."
Me.WindowState = 1
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim msg As Long
    Dim sFilter As String
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Me.WindowState = 0
       Me.Show
       Case WM_LBUTTONDBLCLK
       Case WM_RBUTTONDOWN
       Case WM_RBUTTONUP
  Me.WindowState = 0
  Me.Show
        Case WM_RBUTTONDBLCLK
    End Select
    End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Socket.SendData "$cmdEXIT$"
End Sub

    Private Sub Form_Resize()
If Me.WindowState = 1 Then
   nid.cbSize = Len(nid)
   nid.hwnd = frmChatServer.hwnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = frmChatServer.Icon
   nid.szTip = "Clique aqui para maximizar o programa" & vbNullChar
   Shell_NotifyIcon NIM_ADD, nid
   Me.Hide
   Else
   Me.Show
   End If
End Sub
Private Sub lblApelido_Click()
Clipboard.SetText Socket.LocalIP
End Sub
Private Sub socket_ConnectionRequest(ByVal requestID As Long)
If Socket.State <> sckClosed Then
Socket.Close
End If
Socket.Accept requestID
lblInfo.Caption = "Cliente conectado. IP - " & Socket.RemoteHostIP
End Sub
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim dados As String
Socket.GetData dados
txtPapo.Text = txtPapo.Text + dados & vbCrLf
Me.SetFocus
End Sub
Private Sub Timer1_Timer()
Me.Caption = "TeG Chat Lite - Servidor - por Gabriel Falcão - gabrielfalcao@hotmail.com"
End Sub
Private Sub txtPapo_Change()
On Error Resume Next
txtPapo.SelStart = Len(txtPapo.Text)
End Sub
