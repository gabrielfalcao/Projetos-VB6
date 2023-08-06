VERSION 5.00
Begin VB.MDIForm Principal 
   Appearance      =   0  'Flat
   BackColor       =   &H00F1F1F1&
   Caption         =   "MISIME - Micro Sistema Micro-Empresarial"
   ClientHeight    =   4215
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8685
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu cadastr 
      Caption         =   "Cadastro"
      Begin VB.Menu Clientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu Produtos 
         Caption         =   "Produtos"
      End
      Begin VB.Menu Pedidos 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu Funcionarios 
         Caption         =   "Funcionários"
      End
      Begin VB.Menu Despesas 
         Caption         =   "Despesas"
      End
      Begin VB.Menu Fornecedores 
         Caption         =   "Fornecedores"
      End
   End
   Begin VB.Menu Util 
      Caption         =   "Utilitários"
      Begin VB.Menu text_editor 
         Caption         =   "Editor de Textos"
      End
      Begin VB.Menu calc 
         Caption         =   "Calculadora"
      End
   End
   Begin VB.Menu about_misime 
      Caption         =   "Sobre o MiSiMe"
   End
   Begin VB.Menu quit 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_misime_Click()
frmAbout.Show
End Sub

Private Sub calc_Click()
Dim exe As String
exe = "Calc.exe"
Call Shell(exe)
End Sub

Private Sub Clientes_Click()
frmClientes.Show
frmClientes.WindowState = 2

Unload Fundo
End Sub

Private Sub Despesas_Click()
frmDespesas.Show
frmDespesas.WindowState = 2

Unload Fundo
End Sub

Private Sub Fornecedores_Click()
frmFornecedores.Show
frmFornecedores.WindowState = 2

Unload Fundo
End Sub

Private Sub Funcionarios_Click()
frmFuncionários.Show
frmFuncionários.WindowState = 2

Unload Fundo
End Sub

Private Sub MDIForm_Load()
Fundo.Show
Fundo.WindowState = 2
End Sub

Private Sub Pedidos_Click()
frmPedidos.Show
frmPedidos.WindowState = 2

Unload Fundo
End Sub

Private Sub Picture1_Resize()

End Sub

Private Sub Produtos_Click()
frmProdutos.Show
frmProdutos.WindowState = 2

Unload Fundo
End Sub

Private Sub quit_Click()
Dim intResp As Integer
intResp = MsgBox("Ao sair, todos os arquivos não salvos serão perdidos, deseja realmente sair?", vbYesNo, App.Title)
Select Case intResp
Case vbYes
Unload Me
Case Else
End Select
End Sub

Private Sub text_editor_Click()
Dim editordetexto As String
editordetexto = "Editor.exe"
Call Shell(editordetexto, vbMaximizedFocus)
End Sub
