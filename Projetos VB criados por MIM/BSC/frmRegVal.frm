VERSION 5.00
Begin VB.Form frmRegVal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrador e Validador BSC"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ControlBox      =   0   'False
   Icon            =   "frmRegVal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Configurar!"
      Default         =   -1  'True
      Height          =   375
      Left            =   2242
      TabIndex        =   4
      Top             =   1005
      Width           =   1650
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1845
      TabIndex        =   3
      Top             =   585
      Width           =   3480
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1845
      TabIndex        =   1
      Top             =   255
      Width           =   3480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Organização:"
      Height          =   195
      Left            =   810
      TabIndex        =   2
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   195
      Left            =   810
      TabIndex        =   0
      Top             =   270
      Width           =   465
   End
End
Attribute VB_Name = "frmRegVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Len(Text1.Text) < 2 Then
MsgBox "Nome inválido, tente novamente.", vbCritical, Me.Caption
Else
If Len(Text2.Text) < 2 Then
MsgBox "Organização inválida, tente novamente.", vbCritical, Me.Caption
Else
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "regsnam", Text1.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "regsorg", Text2.Text
'MsgBox "Parabéns, SiCon BSC registrado com sucesso!", vbInformation, Me.Caption
End If
End If
config.Show
Unload Me
End Sub
