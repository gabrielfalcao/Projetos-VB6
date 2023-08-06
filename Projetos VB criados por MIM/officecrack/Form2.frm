VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Lll"
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Español"
      Height          =   345
      Left            =   1890
      TabIndex        =   5
      Top             =   750
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "English"
      Height          =   345
      Left            =   1890
      TabIndex        =   4
      Top             =   405
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Português"
      Height          =   345
      Left            =   1890
      TabIndex        =   3
      Top             =   60
      Width           =   1140
   End
   Begin VB.Label Label3 
      Caption         =   "Seleccionar o idioma:"
      Height          =   270
      Left            =   105
      TabIndex        =   2
      Top             =   810
      Width           =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "Choose the language:"
      Height          =   270
      Left            =   105
      TabIndex        =   1
      Top             =   450
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Escolha a língua:"
      Height          =   270
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "Form2"
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
Form4.Show
Unload Me
End Sub
