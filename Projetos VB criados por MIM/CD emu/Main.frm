VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emuladores - by 7erabyte - www.tebugho.kit.net"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "tebugho"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   120
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Já escolhi!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      Picture         =   "Main.frx":08CA
      TabIndex        =   9
      Top             =   5400
      Width           =   2295
   End
   Begin VB.ListBox l1 
      BackColor       =   &H00C9EDFC&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   480
      TabIndex        =   8
      Top             =   5520
      Width           =   6735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "!!Playstation!!"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MaMe32k(Fliperama)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Callus(Fliperama)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Master System"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GameBoy"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5985
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sair"
      Height          =   615
      Left            =   8640
      Picture         =   "Main.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   735
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   2775
      Left            =   2993
      TabIndex        =   0
      Top             =   2400
      Width           =   3495
      _cx             =   6165
      _cy             =   4895
      FlashVars       =   ""
      Movie           =   "anim.swf"
      Src             =   "anim.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created, developed and designed by Gabriel Falcão - gabrielfalcao@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   11
      Top             =   2040
      Width           =   6885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "!!Liberar Jogos!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   720
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1200
      Picture         =   "Main.frx":1A5E
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   2595
      Left            =   840
      Picture         =   "Main.frx":2328
      Top             =   0
      Width           =   8040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exe As String
Dim dir As String
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

If Option1.Value = True Then
exe = App.Path & "gboy\" & "NO$GMB.exe"
Call Shell(exe, vbNormalFocus)
Else
dir = App.Path & "gboy\"
MkDir "c:\Gameboy"
FileCopy dir & "*.*", "c:\Gameboy"
MsgBox "A pasta foi copiada para:" & "c:\Gameboy"
End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.SimpleText = "Emulador de Gameboy"
End Sub

Private Sub Command3_Click()
If Option1.Value = True Then
exe = App.Path & "Master\" & "MASSAGE.exe"
Call Shell(exe, vbNormalFocus)
Else
dir = App.Path & "Master\"
MkDir "c:\MasterSys"
FileCopy dir & "*.*", "c:\MasterSys"
MsgBox "A pasta foi copiada para:" & "c:\MasterSys"
End If
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.SimpleText = "Emulador de Master System"
End Sub

Private Sub Command4_Click()
If Option1.Value = True Then
exe = App.Path & "Callus\" & "CALLUS95.exe"
Call Shell(exe, vbNormalFocus)
Else
dir = App.Path & "Callus\"
MkDir "c:\Callus"
FileCopy dir & "*.*", "c:\Callus"
MsgBox "A pasta foi copiada para:" & "c:\Callus"
End If

End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.SimpleText = "Emulador de Fliperamas estilo NEO GEO"
End Sub

Private Sub Command5_Click()
If Option1.Value = True Then
exe = App.Path & "MAME32k\" & "MAME32k.exe"
Call Shell(exe, vbNormalFocus)
Else
dir = App.Path & "MAME32k\"
MkDir "c:\MAME32k"
FileCopy dir & "*.*", "c:\MAME32k"
MsgBox "A pasta foi copiada para:" & "c:\MAME32k"
End If
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.SimpleText = "Emulador de Fliperama"
End Sub

Private Sub Command6_Click()
l1.Enabled = True
MsgBox "Atenção para os tipos de emulador de playsation disponíveis", , Me.Caption
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.SimpleText = "Emulador de Playstation"
End Sub

Private Sub Command7_Click()
If l1.ListIndex = 0 Then
MsgBox "O arquivo modificado para abrir em Pentium 3 ou 4 e em Windows XP necessita do emulador original na mesma pasta!!", , Me.Caption
If Option1.Value = True Then
exe = App.Path & "Playstation\" & "Connectix VGS 1.41.exe"
Call Shell(exe, vbNormalFocus)
Else
dir = App.Path & "Playstation\"
MkDir "c:\Playstation"
FileCopy dir & "*.*", "c:\Playstation"
MsgBox "A pasta foi copiada para:" & "c:\Playstation"
End If
Else
If Option1.Value = True Then
exe = App.Path & "Playstation\" & "VGSVideoPatchXP.exe"
Call Shell(exe, vbNormalFocus)
Else
dir = App.Path & "Playstation\"
MkDir "c:\Playstation"
FileCopy dir & "*.*", "c:\Playstation"
MsgBox "A pasta foi copiada para:" & "c:\Playstation"
End If
End If
End Sub

Private Sub Form_Load()
l1.AddItem "Emulador de Playsation para Windows98-98SE-ME Comum e Processador AMD"
l1.AddItem "Emulador de Playsation para WindowsXP e/ou Processador Pentium3 ou 4"
If Len(App.Path) = 3 Then
ShockwaveFlash1.Movie = App.Path & "anim.swf"
Else
ShockwaveFlash1.Movie = App.Path & "\anim.swf"
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF&
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF&
End Sub

Private Sub l1_Click()
Command7.Enabled = True
End Sub

Private Sub Option2_Click()

End Sub

Private Sub Label1_Click()
MsgBox "Visite o site: www.tebugho.kit.net para pegar roms, cracks, seriais e mais... ou visite nosso parceiro e desenvolvedor de sistemas: www.megaaccesshp.hpg.com.br", , Me.Caption
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF0000
End Sub
