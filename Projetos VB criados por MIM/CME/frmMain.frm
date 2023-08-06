VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000F&
   Caption         =   "MiCoMe - Micro Controle Micro Empresarial"
   ClientHeight    =   6735
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10020
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu ctrl 
      Caption         =   "Controles"
      Begin VB.Menu estq 
         Caption         =   "Estoque"
      End
      Begin VB.Menu cli 
         Caption         =   "Clientes"
      End
      Begin VB.Menu forn 
         Caption         =   "Forncecedores"
      End
      Begin VB.Menu fabri 
         Caption         =   "Fabricantes"
      End
   End
   Begin VB.Menu Util 
      Caption         =   "Utilitários"
      Begin VB.Menu txtEdit 
         Caption         =   "Editor de Textos"
      End
      Begin VB.Menu agnd 
         Caption         =   "Agenda"
      End
      Begin VB.Menu calc 
         Caption         =   "Calculadora"
      End
   End
   Begin VB.Menu sai 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub agnd_Click()
On Error Resume Next
Dim agenda As String
If Len(App.Path) > 3 Then
agenda = App.Path & "\agenda.exe"
Else
agenda = App.Path & "agenda.exe"
End If
Call Shell(agenda, vbNormalFocus)
End Sub

Private Sub calc_Click()
On Error Resume Next
Dim calculadora As String
If Len(App.Path) > 3 Then
calculadora = App.Path & "\calculadora.exe"
Else
calculadora = App.Path & "calculadora.exe"
End If
Call Shell(calculadora, vbNormalFocus)

End Sub

Private Sub cli_Click()
frmClientes.Show
End Sub

Private Sub estq_Click()
frmEstoque.Show
End Sub

Private Sub fabri_Click()
frmFabricantes.Show
End Sub

Private Sub forn_Click()
frmFornecedores.Show
End Sub

Private Sub MDIForm_Load()
frmSplash.Show
frmSplash.Width = Me.Width - 500
'frmSplash.Height = Me.Height - 900
End Sub

Private Sub sai_Click()
End
End Sub

Private Sub txtEdit_Click()
On Error Resume Next
Dim editor As String
If Len(App.Path) > 3 Then
editor = App.Path & "\Editor.exe"
Else
editor = App.Path & "Editor.exe"
End If
Call Shell(editor, vbNormalFocus)

End Sub
