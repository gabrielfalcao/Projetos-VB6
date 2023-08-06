VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "FTP por Gabriel Falcão"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   735
      Top             =   1605
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2010
      Left            =   3885
      ScaleHeight     =   2010
      ScaleWidth      =   3015
      TabIndex        =   11
      Top             =   60
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CheckBox Check1 
         Caption         =   "VbCrLf"
         Height          =   195
         Left            =   840
         TabIndex        =   14
         Top             =   765
         Width           =   810
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SendData"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   660
         TabIndex        =   13
         Top             =   1110
         Width           =   1665
      End
      Begin VB.TextBox tData 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   270
         TabIndex        =   12
         Text            =   "PUT"
         Top             =   285
         Width           =   2505
      End
   End
   Begin VB.TextBox tPort 
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
      IMEMode         =   3  'DISABLE
      Left            =   1245
      TabIndex        =   9
      Text            =   "21"
      Top             =   1290
      Width           =   2520
   End
   Begin VB.TextBox log 
      Height          =   1905
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2100
      Width           =   6930
   End
   Begin VB.TextBox tServer 
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
      Left            =   1245
      TabIndex        =   5
      Text            =   "ftp.hpg.com.br"
      Top             =   165
      Width           =   2520
   End
   Begin MSWinsockLib.Winsock w1 
      Left            =   1890
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
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
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Top             =   1680
      Width           =   1650
   End
   Begin VB.TextBox tPass 
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
      IMEMode         =   3  'DISABLE
      Left            =   1245
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "kimk14512"
      Top             =   915
      Width           =   2520
   End
   Begin VB.TextBox tUser 
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
      Left            =   1245
      TabIndex        =   0
      Text            =   "megaaccesshp"
      Top             =   540
      Width           =   2520
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Porta:"
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
      Left            =   510
      TabIndex        =   10
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOG:"
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
      Left            =   60
      TabIndex        =   8
      Top             =   1860
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor:"
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
      Left            =   150
      TabIndex        =   6
      Top             =   195
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
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
      Left            =   510
      TabIndex        =   3
      Top             =   975
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
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
      Left            =   270
      TabIndex        =   1
      Top             =   600
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
w1.Connect tServer, tPort
End Sub

Private Sub Command2_Click()
If Check1.Value = 1 Then
w1.SendData tData.Text & vbCrLf
Else
w1.SendData tData.Text
End If
End Sub

Private Sub log_Change()
log.SelStart = Len(log.Text)
End Sub

Private Sub tData_DblClick()
tData.Text = tData.Text & " " & App.EXEName & ".exe"
End Sub

Private Sub Timer1_Timer()
If w1.State = 7 Then
Picture1.Visible = True
Else
Picture1.Visible = False
End If
End Sub

Private Sub w1_Connect()
w1.SendData "USER " & tUser & vbCrLf
w1.SendData "pass " & tPass & vbCrLf
w1.SendData "LIST" & vbCrLf
w1.SendData "PASV" & vbCrLf
w1.SendData "USER " & tUser & vbCrLf
w1.SendData "pass " & tPass & vbCrLf
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
Dim dados As String
w1.GetData dados
log.Text = log.Text & vbCrLf & dados

End Sub

