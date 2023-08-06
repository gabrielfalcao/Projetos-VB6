VERSION 5.00
Begin VB.Form frmSearchInList 
   BackColor       =   &H00CC9999&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busca em Listbox por Gabriel Falcão"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   ControlBox      =   0   'False
   Icon            =   "frmSearchInList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   233
      ScaleHeight     =   690
      ScaleWidth      =   2775
      TabIndex        =   5
      Top             =   2805
      Width           =   2805
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "www.megaaccesshp.hpg.com.br"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   405
         Width           =   2355
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "gabrielfalcao@hotmail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   390
         TabIndex        =   7
         Top             =   210
         Width           =   1920
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gabriel Falcão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   855
         TabIndex        =   6
         Top             =   15
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2190
      TabIndex        =   4
      Top             =   3570
      Width           =   855
   End
   Begin VB.ListBox lstDados 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Columns         =   2
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      ItemData        =   "frmSearchInList.frx":6852
      Left            =   150
      List            =   "frmSearchInList.frx":6854
      TabIndex        =   2
      Top             =   960
      Width           =   2970
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   360
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dados:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   3
      Top             =   705
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   1
      Top             =   105
      Width           =   660
   End
End
Attribute VB_Name = "frmSearchInList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Este projeto foi desenvolvido por Gabriel Falcão ''
'' Função: Busca "Letra por Letra" numa listbox o   ''
'' critério escrino na textbox                      ''
'' E-MAIL: gabrielfalcao@hotmail.com                ''
'' SITE: www.megaaccesshp.hpg.com.br                ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
End
End Sub

Private Sub cmdSair_Click()
End
End Sub

Private Sub Form_Load()
lstDados.AddItem "Gabriel"
lstDados.AddItem "Rafael"
lstDados.AddItem "Daniel"
lstDados.AddItem "Ariel"
lstDados.AddItem "Maria"
lstDados.AddItem "Marina"
lstDados.AddItem "Marília"
lstDados.AddItem "Mariana"
lstDados.AddItem "João"
lstDados.AddItem "Joel"
lstDados.AddItem "Joab"
lstDados.AddItem "Juan"
lstDados.AddItem "Pedro"
lstDados.AddItem "Ana"
lstDados.AddItem "Camila"
lstDados.AddItem "Sabrina"
End Sub

Private Sub lstDados_Click()
'txtSearch.Text = lstDados.Text
'txtSearch.SetFocus
End Sub

Private Sub txtSearch_Change()
search$ = UCase$(txtSearch.Text)
searchlen = Len(search$)
If searchlen Then
For i = 0 To lstDados.ListCount - 1
If UCase$(Left$(lstDados.List(i), searchlen)) = search$ Then
lstDados.ListIndex = i
Exit For
End If
Next
End If
'If Len(txtSearch.Text) > 1 Then txtSearch.Text = Empty
'txtSearch.SetFocus
End Sub

Private Sub txtSearch_Click()
txtSearch.Text = Empty
End Sub
