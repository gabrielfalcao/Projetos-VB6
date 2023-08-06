VERSION 5.00
Begin VB.Form frmOPT 
   BackColor       =   &H00A4E3AC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opções"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   ControlBox      =   0   'False
   Icon            =   "frmOPT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox savCOD1 
      BackColor       =   &H00A4E3AC&
      Caption         =   "Salvar Código 1"
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1335
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RECUPERAR DADOS PRIMÁRIOS"
      Height          =   1035
      Left            =   180
      Picture         =   "frmOPT.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Recupera o arquivo contendo todos os DRTs primários"
      Top             =   150
      Width           =   2835
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00A4E3AC&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   3300
      TabIndex        =   0
      Top             =   1770
      Width           =   3300
      Begin VB.CommandButton Command3 
         Caption         =   "&Aplicar"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1650
         TabIndex        =   5
         Top             =   0
         Width           =   1650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   345
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1650
      End
   End
   Begin VB.TextBox Text1 
      Height          =   240
      Left            =   780
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmOPT.frx":1794
      Top             =   2805
      Width           =   135
   End
End
Attribute VB_Name = "frmOPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD1", savCOD1.Value
Unload Me
End Sub

Private Sub Command2_Click()
Dim Resposta
Resposta = MsgBox("Se os dados atuais não estiverem corrompidos, a restauração primária irá sobrepor possíveis novos DRTs cadastrados, deseja continuar?", vbYesNo, Me.Caption)
If Resposta = vbYes Then
Open "C:\BSC\drt.txt" For Output As #1
Print #1, Text1.Text
Close #1
MsgBox "Substituído com sucesso!", , Me.Caption
End If
Command3.Enabled = True
End Sub

Private Sub Command3_Click()
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD1", savCOD1.Value
Command3.Enabled = False
End Sub

Private Sub savCOD1_Click()
Command3.Enabled = True
End Sub

Private Sub savCOD1_KeyPress(KeyAscii As Integer)
Command3.Enabled = True
End Sub

Private Sub savCOD2_Click()
Command3.Enabled = True
End Sub

Private Sub savCOD2_KeyPress(KeyAscii As Integer)
Command3.Enabled = True
End Sub
