VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anti-Vírus McAfee Virus Scan 6.05 - Com Crack"
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Crackear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1065
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2685
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Instruções de instalação"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3120
      Left            =   6495
      TabIndex        =   3
      Top             =   2130
      Width           =   2940
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2175
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "Main.frx":08CA
         Top             =   825
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "!!!Atenção!!!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   705
         TabIndex        =   6
         Top             =   345
         Width           =   1590
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Instalar Anti-Vírus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1065
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sair"
      Height          =   615
      Left            =   7350
      Picture         =   "Main.frx":08D5
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5265
      Width           =   1230
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3105
      Left            =   2745
      TabIndex        =   0
      Top             =   1995
      Width           =   3750
      _cx             =   4200919
      _cy             =   4199781
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
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   1080
      Picture         =   "Main.frx":119F
      Stretch         =   -1  'True
      Top             =   3075
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   2595
      Left            =   840
      Picture         =   "Main.frx":26CE
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
If Len(App.Path) = 3 Then
exe = App.Path & "setup.exe"

Else
exe = App.Path & "\setup.exe"
End If
Call Shell(exe)
End Sub

Private Sub Command3_Click()
If Len(App.Path) = 3 Then
exe = App.Path & "crack.exe"
Else
exe = App.Path & "\crack.exe"
End If
Call Shell(exe)
End Sub

Private Sub Form_Load()
Text1.Text = "Logo após a instalação, clique em crackear, daí ao abrir o programa de crack com o título: 'Crack for...', clique em 'Source' e procure o arquivo chamado:'vshwin32.exe' dentro da pasta onde foi instalado o anti-vírus(normalmente na pasta arquivos de programas), daí é só clicar em ok e depois em crack! - Pronto!! o seu antivírus já está pronto para usar!! (Este processo é necessário pois, senão, o ani-vírus só vai funcionar durante 90 dias, caso haja algum problema com o crack, procure no site: 'www.astalavista.com'"
If Len(App.Path) = 3 Then
ShockwaveFlash1.Movie = App.Path & "anim.swf"
Else
ShockwaveFlash1.Movie = App.Path & "\anim.swf"
End If
End Sub
