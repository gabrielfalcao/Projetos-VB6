VERSION 5.00
Begin VB.Form frmXtract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "xTracty - Auto-extrator de resources"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmXtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Extrair"
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
      Height          =   375
      Left            =   1350
      TabIndex        =   6
      Top             =   2700
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Caption         =   "Escolha a origem:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   120
      TabIndex        =   4
      Top             =   135
      Width           =   3540
      Begin VB.ListBox List1 
         Height          =   1620
         ItemData        =   "frmXtract.frx":628A
         Left            =   120
         List            =   "frmXtract.frx":628C
         TabIndex        =   5
         Top             =   255
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Escolha o destino:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5370
      Left            =   3780
      TabIndex        =   0
      Top             =   135
      Width           =   3540
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   4140
         Width           =   3285
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00E7F4FA&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3465
         Left            =   105
         TabIndex        =   2
         Top             =   615
         Width           =   3285
      End
      Begin VB.DriveListBox Drive1 
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
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   3270
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "por Gabriel Falcão gabrielfalcao@hotmail.com www.gabrielfalcao.i8.com"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   780
      Left            =   840
      TabIndex        =   8
      Top             =   4680
      Width           =   2085
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xTracty 1.00 Auto-Extrator de Resources"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1215
      Left            =   975
      TabIndex        =   7
      Top             =   3420
      Width           =   1815
   End
End
Attribute VB_Name = "frmXtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim path As String
Dim dest As String
Dim orig As String



Private Sub Command2_Click()
Dim nome As String
Dim p_ret As String
nome = Text1.Text
p_ret = StrConv(LoadResData(orig, "BMP"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
MsgBox "Extraído com sucesso!"
End Sub

Private Sub Dir1_Change()
retDestino
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub
Private Sub retDestino()
If Len(Dir1.path) = 3 Then
path = Dir1.path
Else
path = Dir1.path & "\"
End If
dest = path & orig
Text1.Text = dest
End Sub

Private Sub Form_Load()
Dir1.path = "C:\"
List1.AddItem "MS1.BMP"
List1.AddItem "MS2.BMP"
List1.ListIndex = 0
retDestino
End Sub

Private Sub List1_Click()
orig = List1.Text
retDestino
Command2.Enabled = True
End Sub
