VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cliente"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin WinsockTwin.Winsock Winsock1 
      Left            =   555
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enviar Dados"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      TabIndex        =   4
      Top             =   1800
      Width           =   1560
   End
   Begin VB.PictureBox st 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FEFADE&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   5190
      TabIndex        =   3
      Top             =   3270
      Width           =   5250
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00DDF3A0&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   675
      Width           =   5115
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E8EAFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2220
      Width           =   5115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar a 220"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   1560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim data As String
Dim recdata As String
Private Sub Command1_Click()
  If Winsock1.State <> 7 Then
  Winsock1.CloseSck
  Winsock1.Connect "127.0.0.1", "2502"
     Do
     If Winsock1.State = 7 Then GoTo 10
     If Winsock1.State = 9 Then GoTo 20
     DoEvents
     Loop

10 stat "Conectado a porta 220"
Exit Sub

20 stat "Erro na conexao a porta 220"
Winsock1.CloseSck
Exit Sub
   End If


End Sub
Private Sub stat(status As String)
 st.Cls
st.Print status
End Sub
Private Sub Command3_Click()
data = Text1.Text
Winsock1.SendData data
End Sub



Private Sub Winsock1_closesck()
 
stat "Servidor Fechado"
End Sub

Private Sub Winsock1_Connect()
 
stat "Conectado na porta 220"
End Sub



Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData recdata
 
stat "Recebendo Dados..."
Text2.Text = Text2.Text & vbCrLf & recdata
 
stat "Dados Recebidos"
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 
stat "Erro"
End Sub

