VERSION 5.00
Begin VB.Form frmIntro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fake Suite"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Não Concordo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2790
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Concordo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2790
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1290
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmIntro.frx":08CA
      Top             =   795
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Se estiver de acordo com os termos acima clique em ""Concordo, caso ao  contrário, clique em ""Não Concordo""."
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   60
      TabIndex        =   3
      Top             =   2085
      Width           =   4575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Atenção Importante, leia o texto abaixo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   435
      Width           =   4110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gerador de Fakes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1335
      TabIndex        =   0
      Top             =   75
      Width           =   1695
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmMenu.Show
Unload Me
End Sub

Private Sub Command2_Click()
End
End Sub
