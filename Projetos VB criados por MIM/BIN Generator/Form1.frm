VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de Arquivos Binários"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   4515
      TabIndex        =   7
      Top             =   270
      Width           =   510
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3030
      TabIndex        =   6
      Top             =   2325
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   3000
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3735
      Top             =   180
   End
   Begin VB.CommandButton Command1 
      Height          =   825
      Left            =   3510
      Picture         =   "Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   810
      Width           =   1905
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   375
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   90
      TabIndex        =   0
      Top             =   810
      Width           =   2895
   End
   Begin VB.TextBox txtBin 
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":4758
      Top             =   2010
      Width           =   2715
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clique em GERAR para criar um arquivo Binário..."
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
      Left            =   15
      TabIndex        =   5
      Top             =   2745
      Width           =   5910
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pasta de Destino:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   675
      TabIndex        =   2
      Top             =   105
      Width           =   1470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
Command1.Enabled = False
txtBin.Text = Empty
End Sub

Private Sub Command2_Click()
Dim mp3 As String
mp3 = StrConv(txtBin.Text, vbUnicode)
  If Len(Dir1.Path) = 3 Then
    Open Dir1.Path & "File.mp3" For Binary As #1
    Else
    Open Dir1.Path & "\File.mp3" For Binary As #1
    End If
   
    Put #1, , mp3
    Close #1
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value < 3000 Then

txtBin.Text = Int(Rnd * 2) & txtBin.Text
Label3.Caption = "Gerando arquivo binário, caractere atual: " & ProgressBar1.Value & "/3000, Aguarde..."
ProgressBar1.Value = ProgressBar1.Value + 1
Else
Label3.Caption = "Gerando arquivo binário, caractere atual: " & "3000/3000, Aguarde..."
ProgressBar1.Value = 3000
Timer1.Enabled = False
Command1.Enabled = True
 ' pega a numeração disponivel de arquivo(freefile)
    If Len(Dir1.Path) = 3 Then
    Open Dir1.Path & "File.bin" For Binary As #1
    Else
    Open Dir1.Path & "\File.bin" For Binary As #1
    End If
   
    Print #1, txtBin.Text
    Close #1
    MsgBox "Arquivo gerado em: " & Dir1.Path & ", com o nome de: File.bin"
    Label3.Caption = "Clique em GERAR para criar um arquivo Binário..."
End If
End Sub

