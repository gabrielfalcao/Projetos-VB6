VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0E42
   ScaleHeight     =   5025
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1305
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1350
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1305
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2820
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   2
      Top             =   255
      Width           =   1065
   End
   Begin Crack.Winsock Winsock1 
      Left            =   420
      Top             =   4005
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Timer prog 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4905
      Top             =   1140
   End
   Begin VB.PictureBox st 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1320
      MousePointer    =   1  'Arrow
      ScaleHeight     =   240
      ScaleWidth      =   3240
      TabIndex        =   0
      Top             =   3480
      Width           =   3240
   End
   Begin VB.Image endDOWN 
      Height          =   600
      Left            =   0
      Picture         =   "Form1.frx":630D4
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image endUP 
      Height          =   600
      Left            =   0
      Picture         =   "Form1.frx":641F6
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Label1 
      Height          =   600
      Left            =   5205
      Picture         =   "Form1.frx":65318
      Top             =   90
      Width           =   540
   End
   Begin VB.Label movy 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   2325
      TabIndex        =   1
      Top             =   60
      Width           =   1245
   End
   Begin VB.Image label4 
      Height          =   555
      Left            =   1875
      Picture         =   "Form1.frx":6643A
      Top             =   3855
      Width           =   2115
   End
   Begin VB.Image sndDown 
      Height          =   555
      Left            =   6465
      Picture         =   "Form1.frx":6A1C4
      Top             =   375
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Image sndUP 
      Height          =   555
      Left            =   6480
      Picture         =   "Form1.frx":6DF4E
      Top             =   495
      Visible         =   0   'False
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim i
 Dim data As String
Dim recdata As String
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim m_Rgn As CBMPRegion
Private mCaptionlessWindowMover As CCaptionlessWindowMover

Private Sub Command1_Click()
  If Winsock1.State <> 7 Then
  Winsock1.CloseSck
  Winsock1.Connect "127.0.0.1", "220"
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

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Picture = endDOWN.Picture
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Picture = endUP.Picture
End Sub

Private Sub label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
label4.Picture = sndDown.Picture
End Sub

Private Sub label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
label4.Picture = sndUP.Picture
End Sub

Private Sub movy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseDown X, Y
End Sub

Private Sub movy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseMove X, Y
End Sub

Private Sub movy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseUp
End Sub

Private Sub Label1_Click()
Winsock1.CloseSck
Unload Me
End Sub
Private Sub Form_Load()
Set m_Rgn = New CBMPRegion
Set mCaptionlessWindowMover = New CCaptionlessWindowMover
  Set mCaptionlessWindowMover.Form = Me
  m_Rgn.CreateFromPic Me.Picture, vbRed
  SetWindowRgn hwnd, m_Rgn.Handle, True
  st.Cls

End Sub

Public Function progress(pb As Control, ByVal Percent As Integer, Optional ByVal ShowPercent = False)
    'Replacement for progress bar..looks nicer also
    Dim sNum                            As String    'use percent
    'Dim Num$
    If Not pb.AutoRedraw Then 'picture in memory ?
        pb.AutoRedraw = -1 'no, make one
    End If
    pb.Cls 'clear picture in memory
    pb.ScaleWidth = 100 'new sclaemodus
    pb.DrawMode = 10 'not XOR Pen Modus
    pb.Print "            Enviando..."
    If ShowPercent = True Then
    num$ = Format$(Percent, "###0") + "%"
    pb.CurrentX = 50 - pb.TextWidth(num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(num$)) / 2
    pb.Print num$ 'print percent
    End If
    pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
    pb.Refresh 'show differents
End Function
Private Sub Form_Unload(Cancel As Integer)
  SetWindowRgn hwnd, 0, False
  m_Rgn.Destroy
  Set m_Rgn = Nothing
End Sub


Private Sub Label4_Click()
data = Text1.Text
Winsock1.SendData data
st.Cls
i = 0
prog.Enabled = True
End Sub

Private Sub prog_Timer()

If i < 100 Then
progress st, i, False
i = i + 1
Else
prog.Enabled = False
progress st, 100, False
st.Cls
st.Print "     Enviado!"
End If
End Sub
Private Sub stat(sst As String)
 st.Cls
st.Print sst
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

