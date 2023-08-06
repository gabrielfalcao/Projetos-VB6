VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   390
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1680
      Width           =   2850
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4425
      Top             =   1620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   345
      Left            =   3090
      TabIndex        =   5
      Top             =   1755
      Width           =   1470
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "Form1.frx":0000
      Left            =   3810
      List            =   "Form1.frx":0002
      TabIndex        =   4
      Top             =   75
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   825
      TabIndex        =   3
      Text            =   "gabrielfalcao@hotmail.com"
      Top             =   495
      Width           =   2925
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   825
      TabIndex        =   2
      Top             =   75
      Width           =   2925
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   915
      Width           =   2925
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   165
      TabIndex        =   0
      Top             =   1290
      Width           =   3585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Msg:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   165
      TabIndex        =   8
      Top             =   945
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   165
      TabIndex        =   7
      Top             =   525
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "De:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   165
      TabIndex        =   6
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim step

Private Sub Command1_Click()
Winsock1.RemoteHost = Combo1.Text
Winsock1.Connect Combo1.Text
step = 0
End Sub

Private Sub Form_Load()
Combo1.AddItem "192.168.0.1"
Combo1.AddItem "192.168.0.220"
Combo1.ListIndex = 0
End Sub

Private Sub Text5_Change()
If Text5.Text = "1" Then Winsock1.SendData ("HELO server" + vbCrLf)
If Text5.Text = "2" Then Winsock1.SendData ("mail from:" & Text2.Text + vbCrLf)
If Text5.Text = "3" Then Winsock1.SendData ("rcpt to:" & Text3.Text + vbCrLf)
If Text5.Text = "4" Then Winsock1.SendData ("data" + vbCrLf)
If Text5.Text = "5" Then Winsock1.SendData (Text4.Text + vbCrLf & "." + vbCrLf)
If Text5.Text = "6" Then Winsock1.SendData ("quit" + vbCrLf)
If Text5.Text = "7" Then Winsock1.Close
If Text5.Text = "7" Then List1.AddItem "-----------------------------------------"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Winsock1.GetData strData
List1.AddItem strData
step = step + 1
Text5.Text = step
End Sub



