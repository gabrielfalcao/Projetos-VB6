VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H00BEA2A2&
   BorderStyle     =   0  'None
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   930
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00BEA2A2&
      Height          =   225
      Left            =   1545
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   465
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BEA2A2&
      Caption         =   "OK"
      Height          =   465
      Left            =   1905
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   225
      UseMaskColor    =   -1  'True
      Width           =   720
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00BEA2A2&
      Caption         =   "Sair"
      Height          =   465
      Left            =   2625
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   225
      UseMaskColor    =   -1  'True
      Width           =   720
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00BEA2A2&
      Caption         =   "Todas as Opções"
      Height          =   255
      Left            =   270
      TabIndex        =   1
      Top             =   195
      Value           =   -1  'True
      Width           =   1605
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00BEA2A2&
      Caption         =   "Sem senha"
      Height          =   255
      Left            =   270
      TabIndex        =   0
      Top             =   450
      Width           =   1605
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
Form2.Show
Me.Visible = False
Else
End If
If Option2.Value = True Then
Form4.Show
Me.Visible = False
Else
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form5.Show
End Sub
