VERSION 5.00
Begin VB.Form Fundo 
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9675
   ControlBox      =   0   'False
   Icon            =   "Fundo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9675
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   10815
      Left            =   0
      ScaleHeight     =   10815
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   0
      Width           =   9675
      Begin VB.Image Image1 
         Height          =   7410
         Left            =   0
         Picture         =   "Fundo.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9600
      End
   End
End
Attribute VB_Name = "Fundo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Height = Picture1.Height
Image1.Width = Picture1.Width
End Sub

Private Sub Form_Resize()

Image1.Height = Picture1.Height
Image1.Width = Picture1.Width
End Sub

Private Sub Picture1_Resize()
Picture1.Height = Me.Height
Image1.Height = Picture1.Height
Image1.Width = Picture1.Width
End Sub
