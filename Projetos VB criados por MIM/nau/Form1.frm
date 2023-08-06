VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Norton Anti-Vírus Update Manager"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   25999
      Left            =   4290
      Top             =   150
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   1605
      TabIndex        =   2
      Text            =   "          Aguarde...          "
      Top             =   0
      Width           =   1515
   End
   Begin MSComctlLib.ProgressBar bb 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   795
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3885
      Top             =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1905
      TabIndex        =   1
      Top             =   300
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
'set timer to interval 10
Dim txt As String
Dim t As String
Let Text1.SelStart = 0
Let Text1.SelLength = 1
Let txt = Text1.SelText
Let Text1.SelText = ""
Let t = Text1.Text
Let Text1.Text = t & txt
End Sub

Private Sub Timer2_Timer()
If bb.Value < 100 Then
bb.Value = bb.Value + 1
Else
bb.Value = 1
End If
Label1.Caption = bb.Value & "%"
End Sub
