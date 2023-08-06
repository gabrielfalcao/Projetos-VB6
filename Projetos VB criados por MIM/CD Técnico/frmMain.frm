VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "..::CD Técnico::.. - ::Menu Principal::"
   ClientHeight    =   6360
   ClientLeft      =   225
   ClientTop       =   600
   ClientWidth     =   8955
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFF00&
      Caption         =   "Abrir Pasta de Jogos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3390
      Width           =   2280
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Update Windows 98SE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5610
      TabIndex        =   28
      Top             =   3930
      Width           =   2280
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3645
      ScaleHeight     =   450
      ScaleWidth      =   5055
      TabIndex        =   24
      Top             =   5715
      Width           =   5085
      Begin VB.CommandButton Command13 
         Caption         =   "Instalar Winrar"
         Height          =   390
         Left            =   3345
         TabIndex        =   26
         Top             =   30
         Width           =   1680
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Instalar Winzip"
         Height          =   390
         Left            =   1665
         TabIndex        =   25
         Top             =   30
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Necessário"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   120
         TabIndex        =   27
         Top             =   60
         Width           =   1500
      End
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Ferramentas de HD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5610
      TabIndex        =   23
      Top             =   3660
      Width           =   2280
   End
   Begin VB.Frame Frame9 
      Caption         =   "Gravação de CDs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   135
      TabIndex        =   20
      Top             =   4365
      Width           =   4320
      Begin VB.CommandButton Command11 
         Caption         =   "Instalar Suporte Para Qualquer Gravadora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   135
         TabIndex        =   22
         Top             =   690
         Width           =   3510
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Instalar Easy CD Creator 5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   135
         TabIndex        =   21
         Top             =   285
         Width           =   3510
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Utilitários de Windows"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   4537
      TabIndex        =   17
      Top             =   4365
      Width           =   4350
      Begin VB.ComboBox cbWindows 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   270
         Width           =   4140
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Executar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1770
         TabIndex        =   18
         Top             =   690
         Width           =   1080
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Internet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   135
      TabIndex        =   14
      Top             =   3135
      Width           =   4320
      Begin VB.ComboBox cdInternet 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   270
         Width           =   4035
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Executar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1590
         TabIndex        =   15
         Top             =   690
         Width           =   1080
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Audio/Som"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   4537
      TabIndex        =   11
      Top             =   1905
      Width           =   4350
      Begin VB.CommandButton Command4 
         Caption         =   "Executar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   13
         Top             =   690
         Width           =   1080
      End
      Begin VB.ComboBox cbAudio 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   270
         Width           =   4140
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DVD/Vídeo/Codecs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   135
      TabIndex        =   8
      Top             =   1905
      Width           =   4320
      Begin VB.ComboBox cbDVD 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   270
         Width           =   4035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Executar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1575
         TabIndex        =   9
         Top             =   690
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Segurança"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   4537
      TabIndex        =   5
      Top             =   675
      Width           =   4350
      Begin VB.ComboBox cbSegurança 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   270
         Width           =   4140
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Executar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1695
         TabIndex        =   6
         Top             =   690
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hacker"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   135
      TabIndex        =   1
      Top             =   675
      Width           =   4320
      Begin VB.CommandButton Command1 
         Caption         =   "Executar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1575
         TabIndex        =   4
         Top             =   690
         Width           =   1080
      End
      Begin VB.ComboBox cbHacker 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         ItemData        =   "frmMain.frx":08CA
         Left            =   135
         List            =   "frmMain.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   4050
      End
   End
   Begin VB.CommandButton cmdEND 
      Caption         =   "Encerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   60
      Width           =   1230
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   1095
      TabIndex        =   29
      Top             =   930
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "www.tebugho.i8.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   555
      TabIndex        =   32
      Top             =   5910
      Width           =   2325
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Criado por 7£r@70s"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   31
      Top             =   5700
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Categorias"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   450
      Left            =   3555
      TabIndex        =   2
      Top             =   75
      Width           =   1860
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdpath As String
Dim f
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdEND_Click()
Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
Select Case cbHacker.ListIndex
Case 0
exec cdpath & "Cracks e Ferramentas Hacker", "XPsp1crk.exe"
Case 1
exec cdpath & "Cracks e Ferramentas Hacker", "WinXP.Activation.v1.1.Portuguese.exe"
Case 2
exec cdpath & "Cracks e Ferramentas Hacker\Bomb Generator", "Bomb Generator.exe"
Case 3
exec cdpath & "Cracks e Ferramentas Hacker\Office XP", "Crack Office XPBR.exe"
Case Else
End Select
End Sub
Private Sub exec(diretorio As String, arquivo As String)
Call ShellExecute(Me.hWnd, "Open", arquivo, "", diretorio, 1)
End Sub


Private Sub Command10_Click()
On Error Resume Next
exec cdpath & "Gravador de CD\EasyCDCreator5", "setup.exe"
End Sub

Private Sub Command11_Click()
On Error Resume Next
exec cdpath & "Gravador de CD\EasyCDCreator5\Atualizacao_EasyCDCreator", "ecdc_v5.3.2.34_basic_bp.exe"
End Sub

Private Sub Command12_Click()
On Error Resume Next
exec cdpath & "Winzip 8.1", "Winzip 8.1.exe"
End Sub

Private Sub Command13_Click()
On Error Resume Next
exec cdpath, "wrar34b2.exe"
End Sub

Private Sub Command14_Click()
On Error Resume Next
Call ShellExecute(Me.hWnd, "Open", cdpath & "Ferramentas para HD", "", cdpath & "Ferramentas para HD", 1)
End Sub

Private Sub Command2_Click()
On Error Resume Next
Select Case cbSegurança.ListIndex
Case 0
exec cdpath & "Cracks e Ferramentas Hacker", "FixBlast.exe"
Case 1
exec cdpath & "Updates Microsoft", "WindowsXP_Patch_NOBlast.exe"
Case 2
exec cdpath & "Internet\Firewall", "zlsSetup_45_538.exe"
Case 3
exec cdpath & "Internet\Firewall", "Keygen_ZONE_ALARM_ALL_VERSIONS.exe"
Case 4
exec cdpath & "Grava Conversa Telefônica", "setup.exe"
Case Else
End Select
End Sub

Private Sub Command3_Click()
On Error Resume Next
Select Case cbDVD.ListIndex
Case 0
exec cdpath & "DVD e Vídeo", "Dr.Divx 1.0.4 Full + keygen + how to crack(hardest) in portuguese.exe"
Case 1
exec cdpath & "DVD e Vídeo", "xvcd.exe"
Case 2
exec cdpath & "DVD e Vídeo\Plug-Ins\Divx", "divx503.exe"
Case 3
exec cdpath & "DVD e Vídeo\Plug-Ins\Indeo", "setup.exe"
Case 4
exec cdpath & "DVD e Vídeo\Plug-Ins\Xvid", "xvid.exe"
Case 5
exec cdpath & "DVD e Vídeo\Power DVD", "fo-pdvd4.exe"
Case 6
exec cdpath & "DVD e Vídeo\Power DVD\crack", "EPS-PowerDVD4.exe"
Case 7
exec cdpath & "DVD e Vídeo\Quick Time 6", "QuickTimeInstaller.exe"
Case 8
Call ShellExecute(Me.hWnd, "Open", cdpath & "DVD e Vídeo\NIKE", "", cdpath & "DVD e Vídeo\NIKE", 1)
Case 9
exec cdpath & "DVD e Vídeo\Windows Media Player 9", "Windows XP.exe"
Case 10
exec cdpath & "DVD e Vídeo\Windows Media Player 9", "95-98-98SE-ME.exe"
Case Else
End Select
End Sub

Private Sub Command4_Click()
On Error Resume Next
Select Case cbAudio.ListIndex
Case 0
exec cdpath & "Som", "mmsetup_9000156_PTB.exe"
Case 1
exec cdpath & "Som", "Keymaker.exe"
Case 2
exec cdpath & "Som", "mmjb9_crack.exe"
Case 3
exec cdpath & "Som", "como crackear.txt"
Case Else
End Select
End Sub

Private Sub Command5_Click()
On Error Resume Next
Select Case cdInternet.ListIndex
Case 0
exec cdpath & "Internet\Internet Explorer 6\Instalação", "ie6setup.exe"
Case 1
exec cdpath & "Internet\Shareaza 2.1", "Shareaza_2.1.0.0.exe"
Case 2
exec cdpath & "Plug-Ins\Flash", "iflash.exe"
Case 3
exec cdpath & "Plug-Ins\Java Virtual Machine XP", "msjavx86.exe"
Case 4
exec cdpath & "Plug-Ins\Removedor do IE", "infsetup.exe"
Case 5
exec cdpath & "Internet\Anti-Popup", "No-Ads.exe"
Case 6
exec cdpath & "Internet\Anti-Popup", "PopPepper.exe"
Case 7
exec cdpath & "Internet\Baixa-Sites", "Httrack.exe"
Case 8
exec cdpath & "Internet\Bate-papo", "PalaceUserWin.exe"
Case 9
exec cdpath & "Internet\Bate-papo", "PalaceBR.exe"
Case 10
exec cdpath & "Internet\Controle de Pulso", "Controle de Pulso.exe"
Case 11
exec cdpath & "Internet\Gerenciador de Downloads", "Download Acelerator 5.3.exe"
Case 12
exec cdpath & "Internet\Gerenciador de Downloads", "setup.exe"
Case 13
exec cdpath & "Internet\Gerenciador de Downloads", "FlashGet 1.40.exe"
Case 14
exec cdpath & "Internet\Messenger 6.2", "SetupDl.exe"
Case 15
Call ShellExecute(Me.hWnd, "Open", cdpath & "Internet\Messenger 6.2\Skins", "", cdpath & "Internet\Messenger 6.2\Skins", 1)
Case Else
End Select
End Sub

Private Sub Command6_Click()
On Error Resume Next
MkDir "C:\Jogos de Windows"

For f = 0 To File1.ListCount
File1.ListIndex = f
FileCopy cdpath & "Jogos" & File1.FileName, "C:\Jogos de Windows\" & File1.FileName
Next f
MsgBox "Arquivos Copiados na pasta C:\Jogos de Windows com SUCESSO!"
End Sub

Private Sub Command7_Click()

exec cdpath & "Ferramentas para HD", "Windows 98SE Update.EXE"
End Sub

Private Sub Command8_Click()
On Error Resume Next
Select Case cbSegurança.ListIndex
Case 0
exec cdpath & "Utilitários de Windows\Backup de Arquivos", "Backup Fácil.exe"
Case 1
exec cdpath & "Utilitários de Windows\Backup de Drivers", "Backup de Drivers.exe"
Case 2
exec cdpath & "Utilitários de Windows\DirectX 9", "dxsetup.exe"
Case 3
exec cdpath & "Utilitários de Windows\Faxina no PC", "Easy Cleaner.exe"
Case 4
exec cdpath & "Utilitários de Windows\Limpa Memória sem Reiniciar", "dsetup(o melhor).exe"
Case 5
exec cdpath & "Utilitários de Windows\Limpa Memória sem Reiniciar", "Fast Defrag.exe"
Case 6
exec cdpath & "Utilitários de Windows\Recuperador de Arquivos Deletados", "setup.exe"
Case Else
End Select
End Sub

Private Sub Command9_Click()
On Error Resume Next
Call ShellExecute(Me.hWnd, "Open", cdpath & "Jogos", "", cdpath & "Jogos", 1)
End Sub

Private Sub Form_Load()
On Error Resume Next
'Carrega a variável cdpath
If Right$(App.Path, 1) <> "\" Then
cdpath = App.Path & "\"
Else
cdpath = App.Path
End If
File1.Path = cdpath & "Jogos"
'Preenche as combo boxes
cbHacker.AddItem "Crack Windows XP"
cbHacker.AddItem "Ativador do WIndows XP BR"
cbHacker.AddItem "Gerador de BOMBs para windows 95 e 98"
cbHacker.AddItem "Crack do Office XP"
cbSegurança.AddItem "Removedor do Worm MSBlast"
cbSegurança.AddItem "Atualização do Windows Anti-Blast"
cbSegurança.AddItem "Firewall Zone Alarm 4.5 PRO"
cbSegurança.AddItem "Gerador de Serial do Zone Alarm"
cbSegurança.AddItem "Gravador de Conversas Telefônicas"
cbDVD.AddItem "DR.Divx Transforma Qualquer vídeo em Divx"
cbDVD.AddItem "X-VCD Player - Reproduz VCDs"
cbDVD.AddItem "CODEC Divx 5"
cbDVD.AddItem "CODEC Indeo"
cbDVD.AddItem "CODEC Xvid"
cbDVD.AddItem "Power DVD 4 - Player de DVDs"
cbDVD.AddItem "Gerador de Serial do Power DVD"
cbDVD.AddItem "Quick Time 6"
cbDVD.AddItem "Abrir pasta com vídeos da NIKE"
cbDVD.AddItem "Windows Media Player 9 para Windows XP"
cbDVD.AddItem "Windows Media Player 9 para Windows 9x"
cdInternet.AddItem "Instalar Internet Explorer 6"
cdInternet.AddItem "Shareaza 2.1 - O Melhor Compartilhador"
cdInternet.AddItem "PLUGIN Shockwave Flash"
cdInternet.AddItem "PLUGIN Java Virtual Machine"
cdInternet.AddItem "Removedor do Internet Explorer"
cdInternet.AddItem "No-Ads - Anti-Popups"
cdInternet.AddItem "PopPepper - Anti-Popups"
cdInternet.AddItem "Baixador de Sites"
cdInternet.AddItem "The Palace - Bate Papo"
cdInternet.AddItem "Tradutor do The Palace para Português"
cdInternet.AddItem "Controle de Pulsos"
cdInternet.AddItem "Download Acelerator - Gerenciador de Downloads"
cdInternet.AddItem "NetAnts - O Melhor Gerenciador de Downloads"
cdInternet.AddItem "FlashGet - Gerenciador de Downloads"
cdInternet.AddItem "Instalar MSN Messenger 6.2"
cdInternet.AddItem "Abrir pasta com Skins do Messenger 6.2"
cbAudio.AddItem "Instalar Music Match Jukebox 9 BR"
cbAudio.AddItem "Gerador de Serial do Music Match"
cbAudio.AddItem "Crack do Music Match Jukebox 9"
cbAudio.AddItem "Tutorial de Como Cracker o Music Match"
cbWindows.AddItem "Backup de Arquivos"
cbWindows.AddItem "Backup de Drivers"
cbWindows.AddItem "DirectX9"
cbWindows.AddItem "Faxina no PC"
cbWindows.AddItem "DramAtic - Libera Memória(O Melhor)"
cbWindows.AddItem "Fast Defrag - Libera Memória"
cbWindows.AddItem "Recuperador de Arquivos Deletados"

cbAudio.ListIndex = 0

cbWindows.ListIndex = 0
cdInternet.ListIndex = 0
cbDVD.ListIndex = 0
cbHacker.ListIndex = 0
cbSegurança.ListIndex = 0


End Sub

