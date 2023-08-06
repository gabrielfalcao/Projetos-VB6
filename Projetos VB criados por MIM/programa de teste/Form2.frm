VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00BEA2A2&
   BorderStyle     =   0  'None
   Caption         =   "Senha"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0442
   ScaleHeight     =   3795
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   2295
      Top             =   2685
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   2250
      Locked          =   -1  'True
      PasswordChar    =   " "
      TabIndex        =   20
      Top             =   1455
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox bsair 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   5400
      Picture         =   "Form2.frx":4C484
      ScaleHeight     =   0.529
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   0.556
      TabIndex        =   19
      Top             =   3210
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   510
      IMEMode         =   3  'DISABLE
      Left            =   285
      TabIndex        =   16
      Top             =   735
      Width           =   5505
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   3765
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ver Senha"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   270
      Left            =   705
      TabIndex        =   18
      Top             =   2565
      Width           =   1125
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Esconder Senha"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   270
      Left            =   450
      TabIndex        =   17
      Top             =   2325
      Width           =   1785
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Apagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   390
      TabIndex        =   15
      Top             =   2805
      Width           =   1875
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Apagar Tudo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   435
      TabIndex        =   14
      Top             =   3165
      Width           =   1875
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   570
      TabIndex        =   13
      Top             =   1890
      Width           =   1455
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Entra"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   345
      Left            =   615
      TabIndex        =   12
      Top             =   1500
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3795
      TabIndex        =   11
      Top             =   2445
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3780
      TabIndex        =   10
      Top             =   1890
      Width           =   210
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3825
      TabIndex        =   9
      Top             =   1335
      Width           =   150
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4080
      TabIndex        =   8
      Top             =   2445
      Width           =   210
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4365
      TabIndex        =   7
      Top             =   2445
      Width           =   240
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4080
      TabIndex        =   6
      Top             =   1890
      Width           =   210
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4365
      TabIndex        =   5
      Top             =   1890
      Width           =   240
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4080
      TabIndex        =   4
      Top             =   1335
      Width           =   210
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4380
      TabIndex        =   3
      Top             =   1335
      Width           =   210
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4080
      TabIndex        =   2
      Top             =   2925
      Width           =   225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite a senha para uso total do programa:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   705
      TabIndex        =   0
      Top             =   300
      Width           =   4740
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   2205
      Left            =   3015
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   2370
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text1.PasswordChar = "*"
Else
End If
If Check1.Value = 0 Then
Text1.PasswordChar = ""
Else
End If
SendKeys "(TAB)"
SendKeys "(TAB)"
End Sub

Private Sub bsair_Click()
Dim intResp As Integer
intResp = MsgBox("Você tem certeza que deseja sair?", vbYesNo, "Atenção!")
Select Case intResp
Case vbYes
End
Case Else
End Select
End Sub

Private Sub bsair_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
bsair.BorderStyle = 1
End Sub

Private Sub bsair_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bsair.BorderStyle = 0
End Sub

Private Sub Command1_Click()
Form4.Show
Me.Hide
End Sub

Private Sub Picture1_Click()
If Text1.Text = "kimk" Then
Text1.Refresh
Form1.Show
Me.Visible = False
Else
Text1.Refresh
Form5.Show
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 1
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 0
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BorderStyle = 1
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BorderStyle = 0
End Sub

Private Sub Command2_Click()
If Text1.Text = Text2.Text Then
Text1.Refresh
Form1.Show
Me.Visible = False
Else
If Text1.PasswordChar = "*" Then
Text1.PasswordChar = ""
Else
End If
Text1.Text = "Tente Novamente"
Text1.ForeColor = &HFF&
Timer1.Enabled = True
End If
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BorderStyle = 1
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BorderStyle = 0
End Sub

Private Sub Command3_Click()
SendKeys "{backspace}"
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BorderStyle = 1
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BorderStyle = 0
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
SendKeys "{tab}"
SendKeys "{tab}"
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
Label10.BorderStyle = 0
Label11.BorderStyle = 0
Text2.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\teste2", "PTAESSST")
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

Private Sub Label10_Click()
SendKeys "{3}"
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.BorderStyle = 1
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
Label10.BorderStyle = 0
Label11.BorderStyle = 0
End Sub

Private Sub Label11_Click()
SendKeys "{0}"
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.BorderStyle = 1
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
Label10.BorderStyle = 0
Label11.BorderStyle = 0
End Sub

Private Sub Label12_Click()
Text1.Text = ""
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.BorderStyle = 1
Label12.ForeColor = &HFF&
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.BorderStyle = 0
Label12.ForeColor = &H80000008
End Sub

Private Sub Label13_Click()
If Text1.PasswordChar = "" Then
Text1.PasswordChar = "*"
Else
End If
Text1.Refresh
Text1.Refresh
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.BorderStyle = 1
End Sub

Private Sub Label13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.BorderStyle = 0
End Sub

Private Sub Label14_Click()
If Text1.PasswordChar = "*" Then
Text1.PasswordChar = ""
Else
End If
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.BorderStyle = 1
End Sub

Private Sub Label14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.BorderStyle = 0
End Sub

Private Sub Label2_Click()
SendKeys "{7}"
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 1
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
Label10.BorderStyle = 0
Label11.BorderStyle = 0
End Sub

Private Sub Label3_Click()
SendKeys "{8}"
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BorderStyle = 1
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
Label10.BorderStyle = 0
Label11.BorderStyle = 0
End Sub

Private Sub Label4_Click()
SendKeys "{9}"
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BorderStyle = 1
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
Label10.BorderStyle = 0
Label11.BorderStyle = 0
End Sub

Private Sub Label5_Click()
SendKeys "{4}"
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BorderStyle = 1
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
Label10.BorderStyle = 0
Label11.BorderStyle = 0
End Sub

Private Sub Label6_Click()
SendKeys "{5}"
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BorderStyle = 1
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
Label10.BorderStyle = 0
Label11.BorderStyle = 0
End Sub

Private Sub Label7_Click()
SendKeys "{6}"
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BorderStyle = 1
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
Label10.BorderStyle = 0
Label11.BorderStyle = 0
End Sub

Private Sub Label8_Click()
SendKeys "{1}"
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BorderStyle = 1
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
Label10.BorderStyle = 0
Label11.BorderStyle = 0
End Sub

Private Sub Label9_Click()
SendKeys "{2}"
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BorderStyle = 1
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
Label10.BorderStyle = 0
Label11.BorderStyle = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Text1.Text = Text2.Text Then
Text1.Refresh
Form1.Show
Me.Visible = False
Else
If Text1.PasswordChar = "*" Then
Text1.PasswordChar = ""
Else
End If
Text1.Text = "Tente Novamente"
Text1.ForeColor = &HFF&
Timer1.Enabled = True
End If
Else
End If
End Sub

Private Sub Timer1_Timer()
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
SendKeys "{Delete}"
Timer1.Enabled = False
End Sub
