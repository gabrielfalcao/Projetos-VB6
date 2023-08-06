VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChatCliente 
   BackColor       =   &H00EFD1AD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FJCR Chat"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmChatCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11910
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1305
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin RichTextLib.RichTextBox txtpapo 
      Height          =   6630
      Left            =   60
      TabIndex        =   13
      Top             =   390
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   11695
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmChatCliente.frx":0CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOptions 
      BackColor       =   &H00FFF4EA&
      Caption         =   "Opções"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10965
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   15
      Width           =   930
   End
   Begin VB.CommandButton cmdConectar 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Conectar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   30
      Width           =   2025
   End
   Begin VB.CommandButton cmdDesconectar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Desconectar"
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
      Left            =   9675
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   45
      Width           =   1065
   End
   Begin VB.TextBox txtApelido 
      BackColor       =   &H00FFFF00&
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
      Left            =   3180
      TabIndex        =   7
      Top             =   30
      Width           =   2055
   End
   Begin VB.TextBox txtEnviar 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1710
      TabIndex        =   3
      Text            =   "Olá"
      Top             =   7140
      Width           =   8295
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "&Enviar"
      Default         =   -1  'True
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
      Left            =   10050
      TabIndex        =   2
      Top             =   7125
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
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
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   2685
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   7455
      Top             =   3105
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.OptionButton msg 
      Caption         =   "Option1"
      Height          =   195
      Left            =   3150
      TabIndex        =   12
      Top             =   2055
      Visible         =   0   'False
      Width           =   975
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
      Left            =   2445
      TabIndex        =   10
      Top             =   90
      Width           =   675
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   975
      TabIndex        =   6
      Top             =   7530
      Width           =   600
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1710
      TabIndex        =   5
      Top             =   7485
      Width           =   9270
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Texto:"
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
      Left            =   1035
      TabIndex        =   4
      Top             =   7185
      Width           =   540
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
      Left            =   5295
      TabIndex        =   0
      Top             =   90
      Width           =   585
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "Opções"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuchat 
      Caption         =   "a"
      Visible         =   0   'False
      Begin VB.Menu mnudata 
         Caption         =   "Texto"
      End
   End
End
Attribute VB_Name = "frmChatCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim version As String
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim buffer As String
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
Private Sub cmdConectar_Click()
On Error Resume Next

  lblInfo.Caption = "Tentando conectar a " & Text2.Text & "..."
      If Socket.State <> 7 Then
  Socket.Close
  Socket.Connect Text2.Text, 2502
     Do
     If Socket.State = 7 Then GoTo 10
     If Socket.State = 9 Then GoTo 20
     DoEvents
     Loop
10 lblInfo.Caption = "Conectado a " & Text2.Text
Exit Sub
20 lblInfo.Caption = "Erro na conexao a " & Text2.Text
Socket.Close
Exit Sub
   End If

End Sub
Private Sub cmdDesconectar_Click()
On Error Resume Next
Socket.Close
lblInfo.Caption = "TeG Chat Lite"
End Sub

Private Sub cmdEnviar_Click()
On Error Resume Next
If txtApelido.Enabled = True Then txtApelido.Enabled = False
If Socket.State <> sckConnected Then
MsgBox "O programa está desconectado..."
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
    txtpapo.Text = txtpapo.Text & vbCrLf & txtApelido.Text & " - IP: " & Text1.Text & vbCrLf & txtEnviar.Text & vbCrLf
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

Private Sub Form_Load()
If App.PrevInstance = True Then Unload Me
On Error Resume Next
Do Until txtApelido.Text <> Empty
txtApelido.Text = InputBox("Digite o seu apelido!!", Me.Caption, "Visitante")
Loop
Dim p_ret As String
If Text2.Text = Empty Then teste
lblInfo.Caption = "FJCR Chat - v1.0"
Text1.Text = Socket.LocalIP
msg.Value = False
End Sub
Private Sub teste()
Dim arquiv As String
Dim linedata As String
arquiv = "C:\ip.txt"
 Open arquiv For Input As #1

    Input #1, linedata
Text2.Text = linedata

  Close #1
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Socket.Close
Unload frmAbout
Unload frmOptions
Unload frmMSG
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim msg As Long
    Dim sFilter As String
    msg = X / Screen.TwipsPerPixelX
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

Private Sub mnuOptions_Click()
On Error Resume Next
frmOptions.Show
End Sub
Private Sub txtPapo_Change()
On Error Resume Next
txtpapo.SelStart = Len(txtpapo.Text)
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
txtpapo.Text = txtpapo.Text & Dados & vbCrLf
End Sub




