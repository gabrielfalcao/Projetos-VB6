VERSION 5.00
Begin VB.Form frmCorrect 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificação de Senha"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   2040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Mudar Senha"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   1140
      Width           =   1755
   End
   Begin VB.CheckBox Command1 
      BackColor       =   &H00000000&
      Caption         =   "&OK"
      ForeColor       =   &H0000FF00&
      Height          =   390
      Left            =   405
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   495
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acesso Permitido"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   420
      TabIndex        =   0
      Top             =   105
      Width           =   1245
   End
End
Attribute VB_Name = "frmCorrect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
If Command1.Value = 1 Then
Command1.Value = 0
End If
End Sub

Private Sub Command2_Click()
ModSenha.Show
End Sub

