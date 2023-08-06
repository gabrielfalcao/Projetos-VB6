VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E2BAA9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kurt C. MP3  Changer - por Gabriel Falcão - gabrielfalcao@hotmail.com"
   ClientHeight    =   3735
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   8295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8295
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   4935
      ScaleHeight     =   690
      ScaleWidth      =   2775
      TabIndex        =   33
      Top             =   2865
      Width           =   2805
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   1050
         ScaleHeight     =   570
         ScaleWidth      =   1650
         TabIndex        =   37
         Top             =   45
         Width           =   1680
         Begin VB.CommandButton Command1 
            Caption         =   "Gravar Dados"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   75
            TabIndex        =   39
            Top             =   60
            Width           =   1500
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Sobre o programa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   75
            TabIndex        =   38
            Top             =   285
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   45
         ScaleHeight     =   570
         ScaleWidth      =   990
         TabIndex        =   34
         Top             =   45
         Width           =   1020
         Begin VB.CommandButton Command4 
            Caption         =   "Colar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   36
            Top             =   285
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Copiar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   35
            Top             =   60
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00EFD1AD&
      Caption         =   "V  MP3 Player  V"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1470
      MaskColor       =   &H00D89970&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3375
      Width           =   1380
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Adicionar à Playlist"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   75
      MaskColor       =   &H00D89970&
      TabIndex        =   21
      Top             =   3375
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Escolher pasta..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2850
      MaskColor       =   &H00D89970&
      TabIndex        =   19
      Top             =   3375
      Width           =   1425
   End
   Begin VB.TextBox track 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7125
      MaxLength       =   2
      TabIndex        =   18
      Top             =   1500
      Width           =   1095
   End
   Begin VB.TextBox year 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7125
      MaxLength       =   4
      TabIndex        =   17
      Top             =   2025
      Width           =   1095
   End
   Begin VB.TextBox comments 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4335
      MaxLength       =   28
      TabIndex        =   14
      Top             =   2550
      Width           =   2820
   End
   Begin VB.TextBox id 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7125
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2550
      Width           =   1095
   End
   Begin VB.TextBox genre 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7125
      TabIndex        =   10
      Top             =   975
      Width           =   1095
   End
   Begin VB.TextBox album 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4335
      TabIndex        =   8
      Top             =   2025
      Width           =   2820
   End
   Begin VB.TextBox artist 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4335
      TabIndex        =   6
      Top             =   1500
      Width           =   2820
   End
   Begin VB.TextBox title 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4335
      TabIndex        =   4
      Top             =   975
      Width           =   2805
   End
   Begin VB.TextBox file 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4335
      TabIndex        =   2
      Top             =   450
      Width           =   3885
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      Left            =   60
      Pattern         =   "*.mp3"
      TabIndex        =   0
      Top             =   203
      Width           =   4215
   End
   Begin VB.ListBox pl 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   225
      Index           =   1
      ItemData        =   "frmMain.frx":0CCA
      Left            =   1335
      List            =   "frmMain.frx":0CCC
      TabIndex        =   22
      Top             =   975
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picMP3 
      Appearance      =   0  'Flat
      BackColor       =   &H00D89970&
      ForeColor       =   &H80000008&
      Height          =   1785
      Left            =   1725
      ScaleHeight     =   1755
      ScaleWidth      =   4815
      TabIndex        =   29
      Top             =   3885
      Width           =   4845
      Begin VB.ListBox pl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF4EA&
         Height          =   1395
         Index           =   0
         ItemData        =   "frmMain.frx":0CCE
         Left            =   1230
         List            =   "frmMain.frx":0CD0
         TabIndex        =   30
         Top             =   285
         Width           =   3525
      End
      Begin VB.Image imgNext 
         Height          =   300
         Index           =   0
         Left            =   675
         Picture         =   "frmMain.frx":0CD2
         Top             =   975
         Width           =   330
      End
      Begin VB.Image imgPrev 
         Height          =   300
         Index           =   0
         Left            =   345
         Picture         =   "frmMain.frx":1264
         Top             =   975
         Width           =   330
      End
      Begin VB.Image imgStop 
         Height          =   300
         Index           =   0
         Left            =   675
         Picture         =   "frmMain.frx":17F6
         Top             =   660
         Width           =   330
      End
      Begin VB.Image imgPause 
         Height          =   300
         Index           =   0
         Left            =   345
         Picture         =   "frmMain.frx":1D88
         Top             =   660
         Width           =   330
      End
      Begin VB.Image imgPlay 
         Height          =   300
         Index           =   0
         Left            =   150
         Picture         =   "frmMain.frx":231A
         Top             =   345
         Width           =   1050
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MiniPlayer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   60
         TabIndex        =   31
         Top             =   75
         Width           =   4695
      End
   End
   Begin VB.Timer tmrMPlayer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1095
      Top             =   900
   End
   Begin VB.Timer tmrFile 
      Interval        =   10
      Left            =   1170
      Top             =   870
   End
   Begin VB.Image imgNext 
      Height          =   300
      Index           =   1
      Left            =   5032
      Picture         =   "frmMain.frx":33EC
      Top             =   4875
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPrev 
      Height          =   300
      Index           =   1
      Left            =   4552
      Picture         =   "frmMain.frx":397E
      Top             =   4875
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgStop 
      Height          =   300
      Index           =   1
      Left            =   4072
      Picture         =   "frmMain.frx":3F10
      Top             =   4875
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPause 
      Height          =   300
      Index           =   1
      Left            =   3592
      Picture         =   "frmMain.frx":44A2
      Top             =   4875
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPlay 
      Height          =   300
      Index           =   1
      Left            =   2392
      Picture         =   "frmMain.frx":4A34
      Top             =   4875
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgNext 
      Height          =   300
      Index           =   2
      Left            =   5032
      Picture         =   "frmMain.frx":5B06
      Top             =   4515
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPrev 
      Height          =   300
      Index           =   2
      Left            =   4552
      Picture         =   "frmMain.frx":5B72
      Top             =   4515
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgStop 
      Height          =   300
      Index           =   2
      Left            =   4072
      Picture         =   "frmMain.frx":5BDC
      Top             =   4515
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPause 
      Height          =   300
      Index           =   2
      Left            =   3592
      Picture         =   "frmMain.frx":5C36
      Top             =   4515
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPlay 
      Height          =   300
      Index           =   2
      Left            =   2392
      Picture         =   "frmMain.frx":5C97
      Top             =   4515
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgEject 
      Height          =   300
      Index           =   1
      Left            =   5512
      Picture         =   "frmMain.frx":5D06
      Top             =   4875
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgEject 
      Height          =   300
      Index           =   2
      Left            =   5512
      Picture         =   "frmMain.frx":6298
      Top             =   4515
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblmode 
      Alignment       =   2  'Center
      BackColor       =   &H00D89970&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1620
      TabIndex        =   28
      Top             =   900
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblbitrate 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1740
      TabIndex        =   27
      Top             =   1290
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00D89970&
      Caption         =   "kbps"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1740
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblKhz 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1740
      TabIndex        =   25
      Top             =   1290
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00D89970&
      Caption         =   "kHz"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1740
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1d 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   45
      Left            =   -15
      Top             =   3780
      Width           =   8565
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivos de Música:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   120
      TabIndex        =   20
      Top             =   30
      Width           =   1425
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ano:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   7125
      TabIndex        =   16
      Top             =   1800
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº da trilha:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   7125
      TabIndex        =   15
      Top             =   1275
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comentário:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   4335
      TabIndex        =   13
      Top             =   2325
      Width           =   915
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo da ID3:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   7125
      TabIndex        =   11
      Top             =   2325
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gênero:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   7125
      TabIndex        =   9
      Top             =   750
      Width           =   585
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Álbum:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   4335
      TabIndex        =   7
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Artista:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   4335
      TabIndex        =   5
      Top             =   1275
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Título:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   4335
      TabIndex        =   3
      Top             =   750
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Arquivo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   4335
      TabIndex        =   1
      Top             =   225
      Width           =   1335
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1335
      TabIndex        =   23
      Top             =   975
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgSlider 
      Height          =   180
      Left            =   1350
      Picture         =   "frmMain.frx":62F3
      Top             =   975
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   225
      Left            =   1335
      Top             =   1005
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currpos As String
Dim CurrentTime As String
Dim TotalFrames As String
Dim TotalTime As String
Dim FramesPerSecond As String
Dim Paused As Boolean
Dim MP3Path As String
Dim PLVisible As Boolean
Dim temp As String
Dim ok As Boolean
Dim showplaya As Boolean
Dim j As Integer
Dim strCommand As String
Dim SlideFlag As Boolean
Dim IX, IY, TX, TY, FX, FY
Dim tmptit As String
Dim tmpart As String

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type
Private mp3 As clsID3v1Tag
Private SysTray As New clsSystrayIcon
Dim mp3file As String, v As Single
Dim tmpTitle As String, tmpTrack As String, tmpAlbum As String, tmpArtist As String, tmpComment As String, tmpYear As String


Public Sub CloseMP3()
  Dim Result As String
  Result = CloseMultimedia(AliasName)
  If Result = "Success" Then
    tmrMPlayer3.Enabled = False
    currpos = 0
  End If
End Sub

Public Sub OpenMP3(FileName As String)
  Dim typeDevice As String
  Dim Result As String
  typeDevice = "MPEGVideo"
  Result = OpenMultimedia(picMP3.hWnd, AliasName, FileName, typeDevice)
  If Result = "Success" Then
    On Error Resume Next
    lblName = pl(0).List(pl(0).ListIndex)
    ReadMP3Header FileName
    End If
End Sub

Public Sub PauseMP3()
  Dim Result As String
  Result = PauseMultimedia(AliasName)
End Sub

Public Sub PlayMP3()
  Dim Result As String
  imgSlider.Move 24, 870: imgSlider.Visible = False
  Result = PlayMultimedia(AliasName, 0, 0)
End Sub

Public Sub ResumeMP3()
  Dim Result As String
  Result = ResumeMultimedia(AliasName)
End Sub

Public Sub StopMP3()
  Dim Result As String
  If SlideFlag = True Then
    Result = StopMultimedia(AliasName)
    Exit Sub
  End If
  imgSlider.Visible = False
  lblName = "MiniPlayer"
  lblbitrate = ""
  lblKhz = ""
  lblmode = ""
  lblTime = ":"
  Result = StopMultimedia(AliasName)
End Sub


Private Sub Command1_Click()
mp3.FileName = mp3file
mp3.title = title.Text
mp3.artist = artist.Text
mp3.album = album.Text
If comments <> Empty Then
mp3.Comment = comments.Text
Else
mp3.Comment = "ID Editado usando Kurt C."
End If
mp3.track = track.Text
 mp3.year = year.Text

If MsgBox("Tem certeza de que gostaria de sobrescrever os dados da MP3?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
If mp3.WriteTag = True Then MsgBox "Gravado com sucesso!"
End If
File1_Click
End Sub

Private Sub Command2_Click()
frmAbout.Show
End Sub

Private Sub Command3_Click()
tmpTitle = title.Text
tmpArtist = artist.Text
tmpAlbum = album.Text
If comments <> Empty Then
tmpComment = comments.Text
Else
tmpComment = "ID Editado usando Kurt C."
End If
tmpTrack = track.Text
 tmpYear = year.Text
End Sub

Private Sub Command4_Click()
Dim ans As Boolean
Dim cnt As Integer
 If title.Text = Empty Then
 title.Text = tmpTitle
 cnt = 1
 End If
 If artist.Text = Empty Then
 artist.Text = tmpArtist
 cnt = cnt + 1
 End If
 If album.Text = Empty Then
 album.Text = tmpAlbum
 cnt = cnt + 1
 End If
If comments.Text = Empty Then
comments.Text = tmpComment
cnt = cnt + 1
End If
 If track.Text = tmpTrack = Empty Then
 track.Text = tmpTrack
 cnt = cnt + 1
 End If
 If year.Text = tmpYear = Empty Then
 year.Text = tmpYear
 cnt = cnt + 1
 End If
 If cnt < 6 Then
If MsgBox("Tem certeza que deseja sobrescrever os dados atuais?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
 title.Text = tmpTitle
 artist.Text = tmpArtist
 album.Text = tmpAlbum
 comments.Text = tmpComment
 track.Text = tmpTrack
 year.Text = tmpYear
 End If
 End If
End Sub

Private Sub Command5_Click()
  Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
  
  szTitle = vbCr & vbCr & "Selecione a Pasta Desejada:"
  
  With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    File1.Path = sBuffer
  End If
End Sub

Private Sub Command6_Click()
If ok = True Then
pl(1).AddItem mp3file
If title.Text = Empty Then
tmptit = "Música Desconhecida"
Else
tmptit = title.Text
End If
If artist.Text = Empty Then
tmpart = "Artista Desconhecido"
Else
tmpart = artist.Text
End If
pl(0).AddItem tmptit & " por " & tmpart
pl(0).ListIndex = pl(0).ListCount - 1
pl(1).ListIndex = pl(1).ListCount - 1
End If
End Sub

Private Sub Command7_Click()
showplaya = True
Command7.Visible = False
Command6.Visible = True
resizeMe
End Sub

Private Sub file_Change()
ok = True
Command6.Enabled = True
End Sub

Private Sub File1_Click()
file.Text = Empty
title.Text = Empty
artist.Text = Empty
album.Text = Empty
id.Text = Empty
comments.Text = Empty
track.Text = Empty
year.Text = Empty

If Len(File1.Path) = 3 Then
mp3file = File1.Path & File1.FileName
Else
mp3file = File1.Path & "\" & File1.FileName
End If
file.Text = mp3file
 mp3.FileName = mp3file
If mp3.ReadTag(v) = False Then Exit Sub

title.Text = mp3.title
artist.Text = mp3.artist
album.Text = mp3.album
id.Text = mp3.id
comments.Text = mp3.Comment
track.Text = mp3.track
year.Text = mp3.year

End Sub

Private Sub File1_DblClick()
If showplaya = True Then Command6_Click
End Sub

Private Sub Form_Load()
Set mp3 = New clsID3v1Tag
File1.Path = "C:\"
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub tmrMPlayer3_Timer()
  Dim Percent As Long
  Dim min As Integer
  Dim sec As Integer
  currpos = GetCurrentMultimediaPos(AliasName)
  CurrentTime = Val(currpos) / Val(FramesPerSecond)
  If SlideFlag = False Then
    imgSlider.left = 26 + Int((currpos / TotalFrames) * 235)
  End If
  min = CurrentTime \ 60
  sec = CurrentTime - (min * 60)
  If sec = "-1" Then sec = "0"
  lblTime = Format$(min, "00") & ":" & Format$(sec, "00")
  'If AreMultimediaAtEnd(AliasName, 0) = True Then
   ' PlayNext
  'End If
End Sub
Private Sub tmrFile_Timer()
  strCommand = GetSetting(App.title, "Config", "Command", "")
  If strCommand <> "" Then
    'ParseCommand strCommand
    SaveSetting App.title, "Config", "Command", ""
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  StopMP3
  CloseMP3
End Sub

Private Sub imgEject_Click(Index As Integer)
  Dim CD As New clsDialog
  temp = CD.OpenDialog(Me, "Arquivos MP3 (*.MP3) |*.mp3|", "Abrir Arquivo", strLOpenPath)
  If temp = "" Then Exit Sub
  'pl(1).Clear
  'pl(0).Clear
  file = LTrim$(temp)
  pl(1).AddItem file
  file.Text = file
  For j = Len(file) To 1 Step -1
    If Mid$(file, j, 1) = "\" Then Exit For
  Next
  P2$ = P1$: p$ = left$(file, j): X$ = Mid$(file, j + 1)
  If left$(file, 1) = "\" Then P2$ = left$(P1$, 2)
  b$ = X$: i = InStr(X$, ".")
  If i > 0 Then b$ = left$(X$, i - 1)
If ok = True Then
pl(1).AddItem mp3file
If title.Text = Empty Then
tmptit = "Música Desconhecida"
Else
tmptit = title.Text
End If
If artist.Text = Empty Then
tmpart = "Artista Desconhecido"
Else
tmpart = artist.Text
End If
pl(0).AddItem tmptit & " por " & tmpart
pl(0).ListIndex = pl(0).ListCount - 1
pl(1).ListIndex = pl(1).ListCount - 1
End If
  WriteINI "MPlayer3", "LOpenPath", CStr(p$)
  strLOpenPath = p$
  imgPlay_Click (0)
End Sub

Private Sub imgEject_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgEject(0) = imgEject(2)
End Sub

Private Sub imgEject_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgEject(0) = imgEject(1)
End Sub

Private Sub imgNext_Click(Index As Integer)
  If pl(0).ListIndex = pl(0).ListCount - 1 Then Exit Sub
  pl(0).ListIndex = pl(0).ListIndex + 1
  If pl(0).ListCount = 0 Then Exit Sub
  StopMP3
  CloseMP3
  OpenMP3 pl(1).List(pl(0).ListIndex)
  DoEvents
  PlayMP3
  
  
End Sub

Private Sub imgNext_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgNext(0) = imgNext(2)
End Sub

Private Sub imgNext_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgNext(0) = imgNext(1)
End Sub

Private Sub imgPause_Click(Index As Integer)
  PauseMP3
  Paused = True
End Sub

Private Sub imgPause_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPause(0) = imgPause(2)
End Sub

Private Sub imgPause_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPause(0) = imgPause(1)
End Sub

Private Sub imgPlay_Click(Index As Integer)
On Error Resume Next
  If Paused = True Then
    ResumeMP3
    Paused = False
    Exit Sub
  End If
  If strFilePath <> "" Then
    If strFilePath = pl(1).List(pl(0).ListIndex) Then
      PlayMP3
    Else
        If pl(0).ListCount = 0 Then Exit Sub
  StopMP3
  CloseMP3
  OpenMP3 pl(1).List(pl(0).ListIndex)
  DoEvents
  PlayMP3
  End If
    Else
    If pl(1).ListCount = 0 Then Exit Sub
    If pl(0).SelCount = 0 Then
      pl(0).ListIndex = 0
     If pl(0).ListCount = 0 Then Exit Sub
  StopMP3
  CloseMP3
  OpenMP3 pl(1).List(pl(0).ListIndex)
  DoEvents
  PlayMP3
  End If
    End If
  

End Sub

Private Sub imgPlay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPlay(0) = imgPlay(2)
End Sub

Private Sub imgPlay_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPlay(0) = imgPlay(1)
End Sub

Private Sub imgStop_Click(Index As Integer)
On Error Resume Next
  StopMP3
  CloseMP3
  SetFocus
End Sub

Private Sub imgStop_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgStop(0) = imgStop(2)
End Sub

Private Sub imgStop_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgStop(0) = imgStop(1)
End Sub

Private Sub lblAbout_Click()
  frmAbout.Show 1
End Sub
Private Sub imgPrev_Click(Index As Integer)
On Error Resume Next
  If pl(0).ListIndex <= 0 Then Exit Sub
  pl(0).ListIndex = pl(0).ListIndex - 1
  If pl(0).ListCount = 0 Then Exit Sub
  StopMP3
  CloseMP3
  OpenMP3 pl(1).List(pl(0).ListIndex)
  DoEvents
  PlayMP3
  
End Sub

Private Sub imgPrev_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPrev(0) = imgPrev(2)
End Sub

Private Sub imgPrev_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPrev(0) = imgPrev(1)
End Sub

Private Sub imgSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
  If SlideFlag = False Then
    IX = X: FX = imgSlider.left
    TX = Screen.TwipsPerPixelX
    SlideFlag = True
  End If
End Sub

Private Sub imgSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If SlideFlag = True Then
    pos = FX + (X - IX) / TX
    If pos < 24 Then pos = 24
    If pos > 260 Then pos = 260
    FX = pos: imgSlider.left = pos
  End If
End Sub

Private Sub imgSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim From As String
  From = Int(((imgSlider.left - 24) / 229) * TotalFrames)
  StopMP3
  Result = PlayMultimedia(AliasName, From, 0)
  SlideFlag = False
End Sub
Private Sub resizeMe()
If showplaya = False Then
Me.Height = 4110
Else
Me.Height = 6105
End If
End Sub

Private Sub pl_DblClick(Index As Integer)
If Index = 0 Then
  If pl(0).ListCount = 0 Then Exit Sub
  StopMP3
  CloseMP3
  OpenMP3 pl(1).List(pl(0).ListIndex)
  DoEvents
  PlayMP3
  End If
End Sub
