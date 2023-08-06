VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "U-0rMz"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Top             =   45
      Width           =   315
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   840
      ItemData        =   "Form1.frx":08CA
      Left            =   60
      List            =   "Form1.frx":08CC
      TabIndex        =   0
      Top             =   2145
      Width           =   2820
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   45
      TabIndex        =   4
      Top             =   3045
      Width           =   4815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Games"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0016ACED&
      Height          =   465
      Left            =   2265
      TabIndex        =   3
      Top             =   450
      Width           =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..:: WinGames ::.."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   900
      TabIndex        =   2
      Top             =   465
      Width           =   3150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jogar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3165
      TabIndex        =   1
      Top             =   2445
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   3330
      Left            =   -900
      Picture         =   "Form1.frx":08CE
      Stretch         =   -1  'True
      Top             =   -600
      Width           =   6675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim exe As String
List1.AddItem "Memória"
List1.AddItem "ChickenRun"
List1.AddItem "PeterPan"
List1.AddItem "ShipWar"
End Sub

Private Sub Label1_Click()
If List1.Text = "Memória" Then
exe = "Memoria.exe"
Call Shell(exe, vbNormalFocus)
End If
If List1.Text = "ChickenRun" Then
exe = "Galinha.exe"
Call Shell(exe, vbNormalFocus)
End If
If List1.Text = "PeterPan" Then
exe = "Peterpan.exe"
Call Shell(exe, vbNormalFocus)
End If
If List1.Text = "ShipWar" Then
exe = "Navinha.exe"
Call Shell(exe, vbNormalFocus)
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = &HFF&
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = &HFF00&
End Sub

Private Sub List1_Click()
Label1.Enabled = True
If List1.Text = "Memória" Then
Label4.Caption = " Jogo de Memória"
End If
If List1.Text = "ChickenRun" Then
Label4.Caption = "Atravesse a rua com a Galinha"
End If
If List1.Text = "PeterPan" Then
Label4.Caption = "Ajude o arqueiro a acertar o alvo"
End If
If List1.Text = "ShipWar" Then
Label4.Caption = "Jogo de Nave Espacial"
End If
End Sub
