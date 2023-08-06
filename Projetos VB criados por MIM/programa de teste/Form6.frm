VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   ClientHeight    =   165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   ScaleHeight     =   165
   ScaleWidth      =   165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If chkLoadTipsAtStartup.Value = 1 Then
Form3.Show
Me.Hide
Else
End If
End Sub
