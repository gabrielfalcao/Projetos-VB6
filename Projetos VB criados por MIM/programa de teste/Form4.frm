VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FCF089&
   BorderStyle     =   0  'None
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   4965
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00BEA2A2&
      Caption         =   "Sobre o programa"
      Height          =   345
      Left            =   225
      MaskColor       =   &H00BEA2A2&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4020
      UseMaskColor    =   -1  'True
      Width           =   8910
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   225
      TabIndex        =   1
      Top             =   165
      Width           =   3030
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BEA2A2&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   570
      Width           =   3030
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versão de Avaliação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3450
      TabIndex        =   6
      Top             =   4395
      Width           =   2490
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   8955
      TabIndex        =   4
      Top             =   60
      Width           =   390
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BEA2A2&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   225
      TabIndex        =   3
      Top             =   1215
      Width           =   8910
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00BEA2A2&
      Caption         =   "Texto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5040
      TabIndex        =   2
      Top             =   375
      Width           =   1320
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text2.Visible = True
Else
End If
If Check1.Value = 0 Then
Text2.Visible = False
Else
End If
End Sub

Private Sub Command1_Click()
Label1.Caption = Text1.Text

End Sub


Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub VScroll1_Change()
Label3.Top = vscroll1.Value
End Sub

Private Sub VScroll1_Scroll()
Label3.Top = vscroll1.Value
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
End Sub

Private Sub Command3_Click()
Form5.Show
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFCF089
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
End Sub
