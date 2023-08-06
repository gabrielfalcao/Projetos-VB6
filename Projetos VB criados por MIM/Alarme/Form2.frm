VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voltar ao relógio"
   ClientHeight    =   360
   ClientLeft      =   7980
   ClientTop       =   6735
   ClientWidth     =   1470
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   1470
   Begin VB.CommandButton Command1 
      Caption         =   "Voltar"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1365
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
