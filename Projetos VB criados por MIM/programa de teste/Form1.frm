VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   4935
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00BEA2A2&
      Caption         =   "Sobre o programa"
      Height          =   720
      Left            =   195
      MaskColor       =   &H00BEA2A2&
      Picture         =   "Form1.frx":97FE4
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3315
      UseMaskColor    =   -1  'True
      Width           =   1740
   End
   Begin VB.PictureBox Command2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   8985
      Picture         =   "Form1.frx":98426
      ScaleHeight     =   300
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   75
      Width           =   315
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   4605
      TabIndex        =   17
      Text            =   "Título da Caixa de Mensagem"
      Top             =   4035
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00BEA2A2&
      Caption         =   "Exibir texto(sem formatação) em caixa de Mensagem"
      Height          =   345
      Left            =   420
      TabIndex        =   16
      Top             =   4005
      Width           =   4125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00BEA2A2&
      Caption         =   "Cores:"
      Height          =   3075
      Left            =   195
      TabIndex        =   9
      Top             =   225
      Width           =   1725
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Vermelho"
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   270
         TabIndex        =   15
         Top             =   405
         Width           =   1185
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Azul"
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   270
         TabIndex        =   14
         Top             =   825
         Width           =   1185
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "Verde"
         ForeColor       =   &H0000FF00&
         Height          =   420
         Left            =   270
         TabIndex        =   13
         Top             =   1245
         Width           =   1185
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00BEA2A2&
         Caption         =   "Preto"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   420
         TabIndex        =   12
         Top             =   2550
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00000000&
         Caption         =   "Amarelo"
         ForeColor       =   &H0000FFFF&
         Height          =   420
         Left            =   270
         TabIndex        =   11
         Top             =   1665
         Width           =   1185
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00000000&
         Caption         =   "Branco"
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   255
         TabIndex        =   10
         Top             =   2100
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   2175
         Left            =   255
         Top             =   375
         Width           =   1230
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00BEA2A2&
         BackStyle       =   1  'Opaque
         Height          =   2745
         Left            =   105
         Top             =   225
         Width           =   1515
      End
   End
   Begin VB.OptionButton Option10 
      BackColor       =   &H00BEA2A2&
      Caption         =   "Negrito e Itálico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2175
      TabIndex        =   5
      Top             =   2535
      Width           =   1980
   End
   Begin VB.OptionButton Option9 
      BackColor       =   &H00BEA2A2&
      Caption         =   "Itálico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2190
      TabIndex        =   4
      Top             =   1980
      Width           =   1095
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H00BEA2A2&
      Caption         =   "Negrito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2190
      TabIndex        =   3
      Top             =   2310
      Width           =   1095
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00BEA2A2&
      Caption         =   "Normal"
      Height          =   270
      Left            =   2190
      TabIndex        =   2
      Top             =   1710
      Value           =   -1  'True
      Width           =   1095
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
      Left            =   2010
      MaskColor       =   &H00BEA2A2&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   585
      UseMaskColor    =   -1  'True
      Width           =   3030
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2010
      TabIndex        =   0
      Top             =   180
      Width           =   3030
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BEA2A2&
      Caption         =   "Tipos de fonte:"
      Height          =   1830
      Left            =   2055
      TabIndex        =   8
      Top             =   1500
      Width           =   2145
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00BEA2A2&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00BEA2A2&
      Height          =   660
      Left            =   7410
      Top             =   3975
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   6255
      TabIndex        =   6
      Top             =   195
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   5295
      TabIndex        =   7
      Top             =   765
      Width           =   3330
   End
End
Attribute VB_Name = "Form1"
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
If Option1.Value = True Then
Label1.ForeColor = &HFF&
Else
End If

If Option2.Value = True Then
Label1.ForeColor = &HFF0000
Else
End If

If Option3.Value = True Then
Label1.ForeColor = &HFF00&
Else
End If
If Option4.Value = True Then
Label1.ForeColor = &H0&
Else
End If
If Option5.Value = True Then
Label1.ForeColor = &HFFFF&
Else
End If
If Option6.Value = True Then
Label1.ForeColor = &HFFFFFF
Else
End If
If Option7.Value = True Then
Label1.FontBold = False
Label1.FontItalic = False
Else
End If
If Option8.Value = True Then
Label1.FontItalic = False
Label1.FontBold = True
Else
End If
If Option9.Value = True Then
Label1.FontBold = False
Label1.FontItalic = True
Else
End If
If Option10.Value = True Then
Label1.FontBold = True
Label1.FontItalic = True
Else
End If
If Check1.Value = 1 Then
Text2.Visible = True
MsgBox Text1.Text, , Text2.Text
End If
End Sub


Private Sub Command2_Click()
Dim intResp As Integer
intResp = MsgBox("Você tem certeza que deseja sair?", vbYesNo, "Atenção!")
Select Case intResp
Case vbYes
End
Case Else
End Select
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BorderStyle = 1
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BorderStyle = 0
End Sub

Private Sub Command3_Click()
Form5.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim intResp As Integer
intResp = MsgBox("Você tem certeza que deseja sair?", vbYesNo, "Atenção!")
Select Case intResp
Case vbYes
End
Case Else
End Select
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub
