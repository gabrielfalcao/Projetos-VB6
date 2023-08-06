VERSION 5.00
Begin VB.Form inicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Escolha o nome dos duelistas:"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   ControlBox      =   0   'False
   Icon            =   "inicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3870
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   630
      Left            =   1245
      TabIndex        =   4
      Top             =   915
      Width           =   1380
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Duelista 2"
      Height          =   735
      Left            =   1875
      TabIndex        =   2
      Top             =   135
      Width           =   1800
      Begin VB.TextBox d2 
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Duelista 1"
      Height          =   735
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   1800
      Begin VB.TextBox d1 
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   1530
      End
   End
End
Attribute VB_Name = "inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
duelo.d1.Caption = d1.Text
duelo.d2.Caption = d2.Text
duelo.Command3.Enabled = True

duelo.Show
Me.Hide
Unload Me
End Sub

Private Sub d1_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(d1.Text) >= 11 Then
MsgBox "Coloque no máximo 11 letras, caso contrário, poderá haver cortes nos nomes"
End If

End Sub

Private Sub d2_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(d2.Text) >= 11 Then
MsgBox "Coloque no máximo 11 letras, caso contrário, poderá haver cortes nos nomes"
End If
End Sub

Private Sub Form_Load()
d1.Text = duelo.d1.Caption
d2.Text = duelo.d2.Caption
End Sub
