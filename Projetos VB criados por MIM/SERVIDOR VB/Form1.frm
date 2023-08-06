VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Servidor"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   5040
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin WinsockTwin.Winsock Winsock1 
      Left            =   1215
      Top             =   1230
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDECF2&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3150
      Width           =   6150
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF8E5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   330
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   765
      Width           =   6180
   End
   Begin VB.PictureBox st 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFEDDC&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      ScaleHeight     =   225
      ScaleWidth      =   6645
      TabIndex        =   0
      Top             =   4725
      Width           =   6675
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sobre..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2100
      TabIndex        =   3
      Top             =   2445
      Width           =   975
   End
   Begin VB.Image endDOWN 
      Height          =   315
      Left            =   2220
      Picture         =   "Form1.frx":70B4C
      Top             =   1110
      Width           =   330
   End
   Begin VB.Image endUP 
      Height          =   315
      Left            =   2850
      Picture         =   "Form1.frx":71122
      Top             =   1095
      Width           =   330
   End
   Begin VB.Image endd 
      Height          =   315
      Left            =   6495
      Picture         =   "Form1.frx":716F8
      Top             =   0
      Width           =   330
   End
   Begin VB.Image sndDOWN 
      Height          =   480
      Left            =   1770
      Picture         =   "Form1.frx":71CCE
      Top             =   1065
      Width           =   1425
   End
   Begin VB.Image sndUP 
      Height          =   480
      Left            =   975
      Picture         =   "Form1.frx":74110
      Top             =   1590
      Width           =   1425
   End
   Begin VB.Image Command1 
      Height          =   480
      Left            =   4950
      Picture         =   "Form1.frx":76552
      Top             =   2295
      Width           =   1425
   End
   Begin VB.Image resDOWN 
      Height          =   480
      Left            =   1335
      Picture         =   "Form1.frx":78994
      Top             =   1140
      Width           =   1425
   End
   Begin VB.Image resUP 
      Height          =   480
      Left            =   1185
      Picture         =   "Form1.frx":7ADD6
      Top             =   1140
      Width           =   1425
   End
   Begin VB.Image Command2 
      Height          =   480
      Left            =   3510
      Picture         =   "Form1.frx":7D218
      Top             =   2265
      Width           =   1425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim data As String
Dim recdata As String
Private mCaptionlessWindowMover As CCaptionlessWindowMover
Private Sub Command1_Click()
data = Text1.Text
Winsock1.SendData data
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.Picture = sndDOWN.Picture
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.Picture = sndUP.Picture
End Sub

Private Sub Command2_Click()
Winsock1.CloseSck
Winsock1.Listen
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Command2.Picture = resDOWN.Picture
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Command2.Picture = resUP.Picture
End Sub

Private Sub endd_Click()
Unload Me
End Sub

Private Sub endd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
endd.Picture = endDOWN.Picture
End Sub

Private Sub endd_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
endd.Picture = endUP.Picture
End Sub

Private Sub Form_Load()
Set mCaptionlessWindowMover = New CCaptionlessWindowMover
  Set mCaptionlessWindowMover.Form = Me
Winsock1.LocalPort = 220
Winsock1.Listen
st.Cls
st.Print "Liberando conexão na porta 220..."
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Winsock1.CloseSck
End Sub

Private Sub Label1_Click()
Form2.Show
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = &HFFFFFF
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = &H0&
End Sub

Private Sub Winsock1_closesck()
st.Cls
st.Print "Servidor Fechado"
End Sub

Private Sub Winsock1_Connect()
st.Cls
st.Print "Conectado na porta 220"
End Sub



Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckclosesckd Then
Winsock1.CloseSck
End If
Winsock1.Accept requestID

st.Cls
st.Print "Cliente requerindo conexão"
If Winsock1.State = 7 Then GoTo 10
10 stat "Conectado a porta 220"
End Sub
Private Sub stat(status As String)
 st.Cls
st.Print status
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData recdata
st.Cls
st.Print "Recebendo Dados..."
Text2.Text = Text2.Text & vbCrLf & recdata
st.Cls
st.Print "Dados Recebidos"
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
st.Cls
st.Print "Erro"
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
