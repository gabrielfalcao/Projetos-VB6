VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviar E-Mail"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Compilar HTML"
      Height          =   300
      Left            =   2685
      TabIndex        =   6
      Top             =   900
      Width           =   1290
   End
   Begin VB.TextBox peça 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFAEA&
      Height          =   285
      Index           =   3
      Left            =   1665
      TabIndex        =   4
      Top             =   510
      Width           =   3525
   End
   Begin VB.TextBox peça 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFAEA&
      Height          =   285
      Index           =   1
      Left            =   1665
      TabIndex        =   1
      Top             =   240
      Width           =   3525
   End
   Begin VB.TextBox html 
      Appearance      =   0  'Flat
      Height          =   3075
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1530
      Width           =   9645
   End
   Begin VB.TextBox peça 
      Appearance      =   0  'Flat
      Height          =   3825
      Index           =   4
      Left            =   8250
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "Form1.frx":0000
      Top             =   4875
      Visible         =   0   'False
      Width           =   7080
   End
   Begin VB.TextBox peça 
      Appearance      =   0  'Flat
      Height          =   3825
      Index           =   2
      Left            =   8250
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0297
      Top             =   4875
      Visible         =   0   'False
      Width           =   7080
   End
   Begin VB.TextBox peça 
      Appearance      =   0  'Flat
      Height          =   3825
      Index           =   0
      Left            =   8250
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":02D0
      Top             =   4875
      Visible         =   0   'False
      Width           =   7080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código do HTML:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   1305
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Assunto:"
      Height          =   195
      Index           =   1
      Left            =   990
      TabIndex        =   5
      Top             =   570
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail Destinatário:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   255
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub Command1_Click()
html.Text = Empty
If peça(1).Text = Empty Then
MsgBox "Campo de Email vazio!"
Exit Sub
Else
If peça(3).Text = Empty Then
MsgBox "Campo de Assunto Vazio!"
Exit Sub
End If
End If
For i = 0 To 4
html.Text = html.Text & peça(i).Text
Next i
End Sub

