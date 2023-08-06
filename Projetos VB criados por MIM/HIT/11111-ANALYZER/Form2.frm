VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Visualizador de Dialogs"
   ClientHeight    =   600
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2445
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
ShowWindow HHDL, 1
Call SetMenu(hWnd, MHDL)
DrawMenuBar hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
DestroyMenu HHDL
Form1.SetFocus
End Sub
