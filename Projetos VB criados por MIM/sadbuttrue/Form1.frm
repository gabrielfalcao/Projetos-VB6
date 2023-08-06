VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Metallica - Sad But True"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   660
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   4590
      Width           =   5130
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Descompactar"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3435
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5580
      Width           =   1980
   End
   Begin VB.DirListBox Dir1 
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
      Height          =   2385
      Left            =   150
      TabIndex        =   2
      Top             =   990
      Width           =   5280
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Height          =   345
      Left            =   165
      TabIndex        =   0
      Top             =   585
      Width           =   2925
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Criado por Gabriel Falcão     gabrielfalcao@hotmail.com   www.gabrielfalcao.i8.com"
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
      Height          =   795
      Left            =   210
      TabIndex        =   7
      Top             =   5415
      Width           =   3495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caminho escolhido..."
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   195
      TabIndex        =   6
      Top             =   4320
      Width           =   2400
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Desculpe-me o gráfico estar ruim... é q eu criei este prog agora, e com pressa..."
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
      Height          =   795
      Left            =   120
      TabIndex        =   3
      Top             =   3585
      Width           =   4275
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   90
      X2              =   5430
      Y1              =   3465
      Y2              =   3465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Miguinhaaaa.... escolhe o lugar ond o arquivo vai ficar..."
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
      Height          =   585
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Width           =   4275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo err
Dim nome As String
Dim p_ret As String
nome = Text1.Text
p_ret = StrConv(LoadResData("ROCK", "SOM"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
MsgBox "Descompactado com sucesso em: " & nome
err:
If err.Number <> 0 Then MsgBox " - Deu pau contate o seu miguim... e informe os seguintes dados: " & "Descrição: " & err.Description & " Número: " & err.Number & "Fonte: " & err.Source

End Sub

Private Sub Dir1_Change()
If Len(Dir1.Path) = 3 Then
Text1.Text = Dir1.Path & "Metallica - Sad But True.mp3"
Else
Text1.Text = Dir1.Path & "\Metallica - Sad But True.mp3"
End If

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
If Len(Dir1.Path) = 3 Then
Text1.Text = Dir1.Path & "Metallica - Sad But True.mp3"
Else
Text1.Text = Dir1.Path & "\Metallica - Sad But True.mp3"
End If

End Sub

