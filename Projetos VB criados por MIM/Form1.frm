VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Live Update! Norton Anti-virus 2003"
   ClientHeight    =   195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   195
   ScaleWidth      =   765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Visible = False
MsgBox "Anti-virus atualizado com sucesso!"
Unload Me
End Sub
