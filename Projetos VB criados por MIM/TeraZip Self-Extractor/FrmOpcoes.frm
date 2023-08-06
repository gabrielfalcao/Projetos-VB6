VERSION 5.00
Begin VB.Form FrmOpcoes 
   BackColor       =   &H00EFD1AD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opções"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAcao 
      Caption         =   "A&ssociar"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox CmbLanguage 
      BackColor       =   &H00FFFAEA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton CmdAcao 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton CmdAcao 
      Caption         =   "&Aplicar"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox ChkDosFormat 
      BackColor       =   &H00EFD1AD&
      Caption         =   "Usar nomes em formato MS-DOS (nomes curtos)"
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
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4695
   End
   Begin VB.CheckBox ChkUseDirectoryInfo 
      BackColor       =   &H00EFD1AD&
      Caption         =   "Restaurar diretório original dos arquivos "
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
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4695
   End
   Begin VB.CheckBox ChkOverwrite 
      BackColor       =   &H00EFD1AD&
      Caption         =   "Sempre substituir arquivos ao descompactar"
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4695
   End
   Begin VB.ComboBox CmbLevel 
      BackColor       =   &H00FFFAEA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FrmOpcoes.frx":0000
      Left            =   1920
      List            =   "FrmOpcoes.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackColor       =   &H00EFD1AD&
      Caption         =   "Linguagem:"
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
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackColor       =   &H00EFD1AD&
      Caption         =   "Nível de compactação:"
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
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "FrmOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SFind As String
Dim LFind As Long

Private Sub CmdAcao_Click(Index As Integer)
 Select Case Index
  Case 0
    If CmbLevel.Text = "" Then CmbLevel.ListIndex = CmbLevel.ListCount - 1
    If CmbLanguage.Text = "" Then CmbLanguage.ListIndex = 0
    LFind = Left(CmbLanguage.Text, 4)
    SaveSetting App.EXEName, "last", "language", LFind
    MyLanguage = LFind
    With FrmMenu.Zipit1
       .CompressionLevel = CmbLevel.ListIndex
       .Overwrite = CBool(ChkOverwrite.Value)
       .UseDOS83Format = CBool(ChkDosFormat.Value)
       .UseDirectoryInfo = CBool(ChkUseDirectoryInfo.Value)
    End With
    Call FrmMenu.LoadCaption
    Call FrmMenu.EMenus
    Unload Me
  
  Case 1
    Dim CAss As New CAssociate
     With CAss
      .Title = LoadResString(MyLanguage + 141)
      .Class = "MCunha98.Zip"
      .Command = FixPath(App.Path) & App.EXEName & ".exe -o"
      .DefaultIcon = FixPath(App.Path) & App.EXEName & ".exe, 101"
      .Associate
      
      .Command = FixPath(App.Path) & App.EXEName & ".exe -a"
      .DefaultIcon = FixPath(App.Path) & App.EXEName & ".exe, 101"
      .Associate "shell\" & LoadResString(MyLanguage + 142) & "\command"
      
      .Command = FixPath(App.Path) & App.EXEName & ".exe -d"
      .DefaultIcon = FixPath(App.Path) & App.EXEName & ".exe, 101"
      .Associate "shell\" & LoadResString(MyLanguage + 143) & "\command"
     End With
  
  Case 2
    Unload Me
 End Select
End Sub

Private Sub Form_Load()
MyLanguage = GetSetting(App.EXEName, "last", "language", 0)
LoadCaption

With FrmMenu.Zipit1
 CmbLevel.ListIndex = .CompressionLevel
 ChkOverwrite.Value = IIf(.Overwrite = True, 1, 0)
 ChkDosFormat.Value = IIf(.UseDOS83Format = True, 1, 0)
 ChkUseDirectoryInfo.Value = IIf(.UseDirectoryInfo = True, 1, 0)
End With

With CmbLanguage
  .Clear
  .AddItem "0000 - Brazilian Portuguese"
  .AddItem "1000 - English"
  SFind = Format(MyLanguage, "0000")
   For LFind = 0 To .ListCount - 1
    If Left(.List(LFind), 4) = SFind Then
     CmbLanguage.ListIndex = LFind
     Exit For
    End If
   Next LFind
  If .ListIndex = -1 Then .ListIndex = 0
End With




Me.Icon = FrmMenu.Icon
End Sub

Private Sub LoadCaption()
 With Me
  .Caption = Replace(LoadResString(MyLanguage + 104), "&", "")
  .L(0).Caption = LoadResString(MyLanguage + 119)
  .L(1).Caption = LoadResString(MyLanguage + 159)
  .ChkOverwrite.Caption = LoadResString(MyLanguage + 120)
  .ChkUseDirectoryInfo.Caption = LoadResString(MyLanguage + 121)
  .ChkDosFormat.Caption = LoadResString(MyLanguage + 122)
  .CmdAcao(0).Caption = LoadResString(MyLanguage + 149)
  .CmdAcao(1).Caption = LoadResString(MyLanguage + 163)
  .CmdAcao(2).Caption = LoadResString(MyLanguage + 127)
  
 End With
End Sub
