VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gerador de Vírus Falsos"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1462
      TabIndex        =   3
      Text            =   "1"
      Top             =   1170
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   2490
      Width           =   3285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gerar Vírus Falsos"
      Height          =   840
      Left            =   540
      Picture         =   "Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   15
      Width           =   2745
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.teratos.blog-se.com.br"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2340
      TabIndex        =   10
      Top             =   3015
      Width           =   2085
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programador: Gabriel Falcão - gabrielfalcao@hotmail.com - www.gabrielfalcao.i8.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   345
      TabIndex        =   9
      Top             =   2820
      Width           =   6120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programador: Gabriel Falcão - gabrielfalcao@hotmail.com - www.gabrielfalcao.i8.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   2835
      Width           =   6120
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":1194
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   660
      Left            =   60
      TabIndex        =   7
      Top             =   1530
      Width           =   3885
   End
   Begin VB.Label Label4 
      BackColor       =   &H00EAEAF4&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form1.frx":1221
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3870
      TabIndex        =   6
      Top             =   45
      Width           =   2820
   End
   Begin VB.Label Label3 
      Height          =   210
      Left            =   1455
      TabIndex        =   5
      Top             =   420
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nº de arquivos a gerar:"
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
      Left            =   930
      TabIndex        =   4
      Top             =   855
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Destino:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1470
      TabIndex        =   2
      Top             =   2220
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.teratos.blog-se.com.br"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   2355
      TabIndex        =   11
      Top             =   3030
      Width           =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VRSTR As String
Private Sub Command1_Click()
Dim X, Z
Dim cnt
For X = 1 To Text2.Text      ' loop de 3 vezes
    Z = FreeFile   ' pega a numeração disponivel de arquivo(freefile)
    Open Text1.Text & "Vírus" & X & ".TXT" For Output As #Z 'Cria um arquivo Z vai ser o arquivo disponível ou seja ainda não utilizado
    Print #Z, VRSTR  ' Grava este texto dentro do arquivo
    Close #Z  'fecha o arquivo
  
Next X
For cnt = 1 To Text2.Text
If cnt = Text2.Text Then MsgBox Text2.Text & " Arquivos criados com sucesso!"
Next cnt
End Sub

Private Sub Form_Load()
VRSTR = "X5O!P%@AP[4\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*"
End Sub

