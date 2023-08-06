VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada Restrita"
   ClientHeight    =   1635
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6135
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   966.012
   ScaleMode       =   0  'User
   ScaleWidth      =   5760.432
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   2032
      TabIndex        =   1
      Top             =   180
      Width           =   3465
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1695
      TabIndex        =   4
      Top             =   1065
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   3300
      TabIndex        =   5
      Top             =   1065
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2032
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   570
      Width           =   3465
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Nome de Usuário:"
      Height          =   270
      Index           =   0
      Left            =   637
      TabIndex        =   0
      Top             =   195
      Width           =   1290
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Senha:"
      Height          =   270
      Index           =   1
      Left            =   637
      TabIndex        =   2
      Top             =   585
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
 End
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    
    If txtPassword = Len(txtUserName.Text) * 1973 Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        frmRegVal.Show
        Unload Me
    Else
        MsgBox "Senha inválida, tente novamente", , "LOGIN"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
  
End Sub

