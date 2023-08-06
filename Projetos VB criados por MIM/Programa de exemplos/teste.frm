VERSION 5.00
Begin VB.Form frmPrincipal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programa de Exemplos - Manipulação de Variáveis e Arquivos"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "teste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerData 
      Interval        =   8000
      Left            =   3720
      Top             =   3750
   End
   Begin VB.Timer timerHora 
      Interval        =   60
      Left            =   4155
      Top             =   3750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selecione o seu Sexo para liberar o programa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1200
      Left            =   3525
      TabIndex        =   1
      Top             =   120
      Width           =   5265
      Begin VB.CommandButton Command2 
         Caption         =   "Feminino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2805
         TabIndex        =   3
         Top             =   315
         Width           =   990
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Indefinido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2805
         TabIndex        =   2
         Top             =   720
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Masculino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1785
         TabIndex        =   0
         Top             =   315
         Width           =   990
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5130
      Index           =   0
      Left            =   0
      ScaleHeight     =   5130
      ScaleWidth      =   9600
      TabIndex        =   4
      Top             =   15
      Visible         =   0   'False
      Width           =   9600
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   375
         Width           =   3000
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   120
         TabIndex        =   17
         Top             =   735
         Width           =   3000
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   120
         TabIndex        =   16
         Top             =   2220
         Width           =   3030
      End
      Begin VB.CommandButton botaoAbrir 
         Caption         =   "Abrir(*.exe)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3795
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2430
         Width           =   1890
      End
      Begin VB.CommandButton botaoCopiar 
         Caption         =   "COPIAR >>"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3795
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3210
         Width           =   1890
      End
      Begin VB.DirListBox Dir2 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1890
         Left            =   6330
         TabIndex        =   13
         Top             =   2505
         Width           =   3000
      End
      Begin VB.DriveListBox Drive2 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6330
         TabIndex        =   12
         Top             =   2190
         Width           =   3000
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   1
         Left            =   45
         ScaleHeight     =   405
         ScaleWidth      =   9480
         TabIndex        =   7
         Top             =   4665
         Width           =   9510
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Data:"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1343
            TabIndex        =   11
            Top             =   75
            Width           =   810
         End
         Begin VB.Label lblData 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   2138
            TabIndex        =   10
            Top             =   75
            Width           =   3720
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Hora:"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5858
            TabIndex        =   9
            Top             =   75
            Width           =   810
         End
         Begin VB.Label lblHora 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   6653
            TabIndex        =   8
            Top             =   75
            Width           =   1605
         End
      End
      Begin VB.CommandButton botaoDeletar 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Deletar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3795
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3600
         Width           =   1890
      End
      Begin VB.CommandButton botaoRenomeia 
         Caption         =   "Renomear"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3795
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2820
         Width           =   1890
      End
      Begin VB.Label observações1 
         BackColor       =   &H000000FF&
         Caption         =   $"teste.frx":628A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   810
         Left            =   3240
         TabIndex        =   21
         Top             =   1350
         Visible         =   0   'False
         Width           =   5985
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origem:"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   225
         TabIndex        =   20
         Top             =   135
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6330
         TabIndex        =   19
         Top             =   1950
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_GotFocus() 'evento de: quando o "command1" recebe foco
Command2.SetFocus ' faz com que o foco seja passado para o "command2"
Command2.Default = True ' faz com que o command2 seja o botão padrão
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
On Error Resume Next
MsgBox "MENTIRA!", vbCritical, Me.Caption
End Sub

Private Sub Command2_Click()
On Error Resume Next
MsgBox "Erro no sistema, sexo incompatível com o usuário" & vbCrLf & "ACESSO NEGADO!", vbCritical, Me.Caption
End Sub

Private Sub botaoAbrir_Click()
On Error Resume Next
On Error GoTo erroabrir
Dim arquivo As String ' Declara a variável "arquivo" como sendo uma string(valores de texto)
Dim extensao As String 'Declara a varíável extensao
Dim executavel As Boolean 'declara a varíavel executave como sendo BOOLEAN(verdadeiro ou falso)
'============
If Len(File1.Path) = 3 Then ' se a quantidade de caracteres do diretório do file1 for igual a 3 (por exemplo: "c:\") o arquivo não precisa de barra(\) antes...
arquivo = File1.Path & File1.FileName
Else ' caso contrário ele precisa de adicionar a barra "\"
arquivo = File1.Path & "\" & File1.FileName
End If
'==========
If Right$(arquivo, 3) = "exe" Then 'se os três últimos caracteres(ou seja: da direira para a esquerda da varíavel ARQUIVO forem = a "EXE" então ele diz que é um arquivo execútável(EXECUTAVEL=TRUE)
executável = True '                                               se quiser saber os tres primeiros(esquerda para a direita) vc tem que trocar o right$ por left$
Else ' caso contrário diz que não é executável...
executável = False
End If

If executável = True Then 'se for executável:
Call Shell(arquivo, vbNormalFocus) '      ele chama(call) o shell(aplicativo) que é a variável arquivo e faz ele ter o FOCO(vbNormalFocus)
End If 'Importante !! TODA VEZ QUE VC CRIAR UM "IF" VC TEM QUE CRIAR UM "END IF"!!!!!
erroabrir:
If Err.Number <> 0 Then
MsgBox "OCORREU UM ERRO" & vbCrLf & vbCrLf & "Nº: " & Err.Number & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "No Objeto: " & Err.Source, vbCritical, "Erro ao Copiar"
End If
End Sub

Private Sub botaoCopiar_Click()
On Error GoTo erroCopiar
Dim origem As String ' Declara a variável "origem" como sendo uma string(valores de texto)
Dim destino As String ' Declara a variável "destino" como sendo uma string(valores de texto)
If File1.Path <> Dir2.Path Then ' se o diretório de file1 for diferente do diretório de dir2 então...
If Len(File1.Path) = 3 Then
origem = File1.Path & File1.FileName
Else
origem = File1.Path & "\" & File1.FileName
End If
'==========
If Len(Dir2.Path) = 3 Then
destino = Dir2.Path & File1.FileName
Else
destino = Dir2.Path & "\" & File1.FileName
End If
FileCopy origem, destino
'=======
MsgBox "Arquivo copiado com sucesso!", , Me.Caption
Else
MsgBox "Não é possível copiar para a mesma pasta!", vbInformation, "Cópia de Arquivos"
End If
erroCopiar:  'tratar erros          vbcrlf dá o espaço de uma linha
If Err.Number <> 0 Then
MsgBox "OCORREU UM ERRO" & vbCrLf & vbCrLf & "Nº: " & Err.Number & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "No Objeto: " & Err.Source, vbCritical, "Erro ao Copiar"
End If
End Sub

Private Sub botaoDeletar_Click()
On Error Resume Next
Dim arquivo As String ' Declara a variável "arquivo" como sendo uma string(valores de texto)
If Len(File1.Path) = 3 Then ' se a quantidade de caracteres do diretório do file1 for igual a 3 (por exemplo: "c:\") o arquivo não precisa de barra(\) antes...
arquivo = File1.Path & File1.FileName
Else ' caso contrário ele precisa de adicionar a barra "\"
arquivo = File1.Path & "\" & File1.FileName
End If
If MsgBox("Deseja realmente EXCLUIR: " & vbCrLf & arquivo & " ?", vbYesNo, "Confirmação de exclusão") = vbYes Then
Kill arquivo ' exibe caixa de mensagem perguntado se tem certeza que quer deletar o arquivo...
End If
File1.Refresh 'atualiza o file1
End Sub

Private Sub botaoRenomeia_Click()
Dim pasta As String ' Declara a variável "arquivo" como sendo uma string(valores de texto)
Dim novonome As String
On Error Resume Next
If Len(File1.Path) = 3 Then ' se a quantidade de caracteres do diretório do file1 for igual a 3 (por exemplo: "c:\") o arquivo não precisa de barra(\) antes...
pasta = File1.Path
Else ' caso contrário ele precisa de adicionar a barra "\"
pasta = File1.Path & "\"
End If
novonome = InputBox("Digite o novo nome do arquivo:", "Renomear Arquivo")
If novonome <> File1.FileName Then
FileCopy pasta & File1.FileName, pasta & novonome
If Len(novonome) > 5 Then Kill pasta & File1.FileName 'se o número de caracteres digitados na caixa de dialogo forem maiores do que 5 então ele deleta o rauivo de nome anterior
'a função da linha acima é necessária pois se o usuário cancelar o arquivo original não é deletado...
If novonome <> Empty Then MsgBox "Arquivo renomeado!" & vbCrLf & "DE: " & File1.FileName & " PARA: " & novonome
File1.Refresh 'atualiza o file1
End If

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'quando passar o mouse em cima do botão...
If Command1.Top = 315 Then 'se a altura do botão no formulário for = a 315 então
Command1.Top = 690 ' colocar o botão na altura 690
Else 'caso contrário...
Command1.Top = 315 'colocar na altura 315
End If 'finaliza if
End Sub

Private Sub Command3_Click()
On Error Resume Next
MsgBox "Sexo confirmado!" & vbCrLf & "ACESSO PERMITIDO", vbQuestion, Me.Caption
Picture1(0).Enabled = True 'libera o programa(que está dentro do objeto picture1
Picture1(0).Visible = True 'mostra controles que estão dentro da picture1
Frame1.Visible = False ' some escolha de sexo
'O B S E R V A Ç Ã O!!!!!
'O controle picture serve também para guardar qualquer controle desenhável do formulário
End Sub

Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1.Path
'faz o diretório de arquivos do objeto file ser igual ao do objeto dir1
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
'dir1.path -  é a pasta ATUAL do objeto dir1
'drive1.drive é o drive atual do objeto drive
'Dir1.Path = Drive1.Drive faz o diretório mudar ao mudar o drive

End Sub

Private Sub Drive2_Change()
On Error Resume Next
Dir2.Path = Drive2.Drive
End Sub

Private Sub File1_Click()
On Error Resume Next
'ao clicar ele chama a sub que libera os botões
liberarBotoes
End Sub

Private Sub Form_Load()
On Error Resume Next
lblData.Caption = Day(Date) & " de " & MonthName(Month(Date)) & " de " & Year(Date) 'data por extenso
End Sub

Private Sub Label5_Click()

End Sub

Private Sub TimerData_Timer()
On Error Resume Next
lblData.Caption = Day(Date) & " de " & MonthName(Month(Date)) & " de " & Year(Date) 'data por extenso
End Sub

Private Sub timerHora_Timer()
On Error Resume Next
lblHora.Caption = Time
End Sub

'Você também pode criar uma sub personalizada para servir de função ou retornar dados, no caso abaixo eu vou usar como a função de habilitar os botões de copiar, abrir, e deletar arquivos...
Private Sub liberarBotoes()
On Error Resume Next

If Right$(File1.FileName, 3) = "EXE" Or Right$(File1.FileName, 3) = "exe" Then 'se os três últimos caracters do nome do arquivo de FILE1 for EXE ou exe então...
botaoAbrir.Enabled = True 'libera o botão Abrir
Else
botaoAbrir.Enabled = False 'desabilita o botão abrir
End If
botaoDeletar.Enabled = True
botaoCopiar.Enabled = True
botaoRenomeia.Enabled = True
End Sub
