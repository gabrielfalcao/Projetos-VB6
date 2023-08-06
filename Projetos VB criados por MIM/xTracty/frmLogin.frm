VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "xTracty Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4905
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   4605.529
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
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
      Left            =   1267
      TabIndex        =   1
      Text            =   "gabrielfalcao"
      Top             =   135
      Width           =   3480
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Entrar"
      Default         =   -1  'True
      Height          =   390
      Left            =   1882
      TabIndex        =   4
      Top             =   1050
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
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
      Left            =   1267
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   3480
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuário:"
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
      Index           =   0
      Left            =   157
      TabIndex        =   0
      Top             =   150
      Width           =   960
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Senha:"
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
      Index           =   1
      Left            =   157
      TabIndex        =   2
      Top             =   540
      Width           =   720
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean

Private Sub cmdOK_Click()

    If txtUserName.Text = LCase$("gabrielfalcao") Then
    If txtPassword = "kimk2502" Then
        LoginSucceeded = True
        Unload Me
    Else
        MsgBox "Senha inválida tente novamente", , "xTracty Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
    Else
        MsgBox "Usuário inválido tente novamente", , "xTracty Login"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
        End If
        If LoginSucceeded = True Then
        frmXtract.Show
        Unload Me
        End If
End Sub
