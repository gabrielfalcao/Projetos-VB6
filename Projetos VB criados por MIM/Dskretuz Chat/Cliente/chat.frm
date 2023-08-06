VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmChatCliente 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   8910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "chat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "chat.frx":3F3A
   ScaleHeight     =   8910
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   9990
      ScaleHeight     =   690
      ScaleWidth      =   1560
      TabIndex        =   16
      Top             =   960
      Width           =   1620
      Begin VB.CommandButton cmdConectar 
         BackColor       =   &H00EFD1AD&
         Caption         =   "&Conectar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   1560
      End
      Begin VB.CommandButton cmdDesconectar 
         BackColor       =   &H00DADAFC&
         Caption         =   "&Desconectar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   1560
      End
   End
   Begin VB.OptionButton msg 
      Caption         =   "Option1"
      Height          =   150
      Left            =   10320
      TabIndex        =   8
      Top             =   330
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtPapo 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEE7D6&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1065
      Width           =   9390
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D6FEDA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6285
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   705
      Width           =   2685
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "&Enviar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10320
      TabIndex        =   5
      Top             =   7665
      Width           =   1215
   End
   Begin VB.TextBox txtEnviar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6FBE6&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   225
      TabIndex        =   4
      Text            =   "Olá"
      Top             =   7320
      Width           =   9405
   End
   Begin VB.TextBox txtApelido 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6FEDA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3525
      TabIndex        =   3
      Text            =   "««§|Ð""§k®ë†µz 5|§»» "
      Top             =   690
      Width           =   2055
   End
   Begin VB.TextBox txtIp 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6FEDA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   705
      TabIndex        =   2
      Top             =   690
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9645
      TabIndex        =   1
      Top             =   210
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9285
      TabIndex        =   0
      Top             =   210
      Width           =   330
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   9675
      Top             =   4455
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seu IP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5640
      TabIndex        =   15
      Top             =   750
      Width           =   585
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   285
      TabIndex        =   14
      Top             =   7095
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   930
      TabIndex        =   13
      Top             =   8175
      Width           =   6510
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   195
      Left            =   3885
      TabIndex        =   12
      Top             =   7950
      Width           =   600
   End
   Begin VB.Label lblApelido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apelido:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2790
      TabIndex        =   11
      Top             =   750
      Width           =   675
   End
   Begin VB.Label lblIp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   435
      TabIndex        =   10
      Top             =   750
      Width           =   225
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Height          =   525
      Left            =   30
      MousePointer    =   5  'Size
      TabIndex        =   9
      Top             =   105
      Width           =   10185
   End
End
Attribute VB_Name = "frmChatCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim version As String
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim Buffer As String
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
Private Type SECURITY_ATTRIBUTES
nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private mCaptionlessWindowMover As CCaptionlessWindowMover
Dim m_Rgn As CBMPRegion
Private Sub cmdConectar_Click()
On Error Resume Next
If txtIp.Text = "" Then
lblInfo.Caption = "Digite o IP do Servidor"
txtIp.SetFocus
  Else
  lblInfo.Caption = "Tentando conectar a " & txtIp.Text & "..."
      If Socket.State <> 7 Then
  Socket.Close
  Socket.Connect txtIp.Text, 2502
     Do
     If Socket.State = 7 Then GoTo 10
     If Socket.State = 9 Then GoTo 20
     DoEvents
     Loop
10 lblInfo.Caption = "Conectado a " & txtIp.Text
Exit Sub
20 lblInfo.Caption = "Erro na conexao a " & txtIp.Text
Socket.Close
Exit Sub
   End If
End If
End Sub
Private Sub cmdDesconectar_Click()
On Error Resume Next
Socket.Close
lblInfo.Caption = "««§|D''§k®ë†µz|§»» CHAT"
End Sub

Private Sub cmdEnviar_Click()
On Error Resume Next
If txtApelido.Enabled = True Then txtApelido.Enabled = False
If Socket.State <> sckConnected Then
MsgBox "Cliente Desconectado..."
 Else
 If txtApelido.Text = "" Then
 lblInfo.Caption = "Escolha um apelido!"
 txtApelido.SetFocus
   Else
   If txtEnviar.Text = "" Then
   lblInfo.Caption = "Digite algo para enviar!"
   txtEnviar.SetFocus
    Else
    Socket.SendData vbCrLf & txtApelido.Text & " - IP: " & Text1.Text & vbCrLf & txtEnviar.Text
    txtPapo.Text = txtPapo.Text & vbCrLf & txtApelido.Text & " - IP: " & Text1.Text & vbCrLf & txtEnviar.Text & vbCrLf
    txtEnviar.Text = ""
    End If
  End If
End If
End Sub

Private Sub cmdSair_Click()
On Error Resume Next
End
End Sub

Private Sub cmdSobre_Click()
On Error Resume Next
frmAbout.Show
End Sub

Private Sub cmdOptions_Click()
mnuOptions_Click
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Me.WindowState = 1
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then Unload Me
On Error Resume Next

Set m_Rgn = New CBMPRegion
Set mCaptionlessWindowMover = New CCaptionlessWindowMover
  Set mCaptionlessWindowMover.Form = Me
  m_Rgn.CreateFromPic Me.Picture, vbWhite
  SetWindowRgn hwnd, m_Rgn.Handle, True

lblInfo.Caption = "««§|D''§k®ë†µz|§»» CHAT"
Text1.Text = Socket.LocalIP
msg.Value = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Socket.Close
Unload frmAbout
Unload frmOptions
Unload frmMSG
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim msg As Long
    Dim sFilter As String
    msg = x / Screen.TwipsPerPixelX
    Select Case msg
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
Me.WindowState = 0
       Me.Show
       Case WM_RBUTTONDOWN

       Case WM_RBUTTONUP
  Me.WindowState = 0
  Me.Show
 
       Case WM_RBUTTONDBLCLK
    End Select
    End Sub
Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
   nid.cbSize = Len(nid)
   nid.hwnd = frmChatCliente.hwnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = frmChatCliente.Icon
   nid.szTip = "Clique aqui para maximizar o programa" & vbNullChar
   Shell_NotifyIcon NIM_ADD, nid
   Me.Hide
   Else
   Me.Show
   End If
End Sub

Private Sub mnuAbout_Click()
On Error Resume Next
frmAbout.Show
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Handle the form's MouseDown event
  mCaptionlessWindowMover.HandleMouseDown x, y
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Handle the form's MouseMove event
  mCaptionlessWindowMover.HandleMouseMove x, y
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Handle the form's MouseUp event
  mCaptionlessWindowMover.HandleMouseUp
End Sub

Private Sub mnuOptions_Click()
On Error Resume Next
frmOptions.Show
End Sub

Private Sub txtPapo_Change()
On Error Resume Next
txtPapo.SelStart = Len(txtPapo.Text)
End Sub
Private Sub socket_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Dados As String
Socket.GetData Dados
If Dados = "$cmdEXIT$" Then
Unload Me
Exit Sub
End If
If Me.WindowState = 1 Then
If msg.Value = False Then
frmMSG.Show
Load frmMSG
frmMSG.Text1.Text = Dados
Else
frmMSG.Text1.Text = Dados
frmMSG.Timer1.Enabled = True
End If
End If
txtPapo.Text = txtPapo.Text & Dados & vbCrLf
End Sub





