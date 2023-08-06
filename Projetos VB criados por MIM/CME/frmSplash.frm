VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   165
      Picture         =   "frmSplash.frx":0000
      Top             =   285
      Width           =   8250
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Left = (Me.Width - Image1.Width) \ 2
Image1.Top = (Me.Height - Image1.Height) \ 2 'centre the form on the scre
End Sub

Private Sub Form_Resize()

If Me.Width < 9765 Then
Me.Width = 9765
Else
If Me.Height < 10830 Then
Me.Height = 10830
Else
Image1.Left = (Me.Width - Image1.Width) \ 2
Image1.Top = (Me.Height - Image1.Height) \ 2 'centre the form on the scre
End If
End If
End Sub

