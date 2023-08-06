VERSION 5.00
Begin VB.Form agenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "agenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000080FF&
      Height          =   2010
      Left            =   4125
      ScaleHeight     =   1950
      ScaleWidth      =   1380
      TabIndex        =   6
      Top             =   465
      Width           =   1440
      Begin VB.CommandButton Command1 
         Caption         =   "&Sair"
         Height          =   510
         Index           =   2
         Left            =   45
         TabIndex        =   9
         Top             =   1365
         Width           =   1305
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Excluir"
         Height          =   510
         Index           =   1
         Left            =   45
         TabIndex        =   8
         Top             =   735
         Width           =   1305
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Incluir"
         Height          =   510
         Index           =   0
         Left            =   45
         TabIndex        =   7
         Top             =   105
         Width           =   1305
      End
   End
   Begin VB.TextBox Text1 
      DataField       =   "TELEFONE"
      DataSource      =   "Data1"
      Height          =   330
      Index           =   2
      Left            =   195
      TabIndex        =   5
      Top             =   2115
      Width           =   2385
   End
   Begin VB.TextBox Text1 
      DataField       =   "ENDERECO"
      DataSource      =   "Data1"
      Height          =   330
      Index           =   1
      Left            =   195
      TabIndex        =   4
      Top             =   1350
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      DataField       =   "NOME"
      DataSource      =   "Data1"
      Height          =   330
      Index           =   0
      Left            =   195
      TabIndex        =   3
      Top             =   585
      Width           =   3495
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "dBASE III;"
      DatabaseName    =   "C:\teste"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "AGENDA"
      Top             =   2880
      Width           =   5850
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Telefone:"
      Height          =   195
      Left            =   195
      TabIndex        =   2
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Endereço:"
      Height          =   195
      Left            =   195
      TabIndex        =   1
      Top             =   1035
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   270
      Width           =   465
   End
End
Attribute VB_Name = "agenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As Recordset

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
rs.AddNew
Text1(0).SetFocus
Case 1
If MsgBox("Confirma exclusão deste registro? " & rs!nome, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
rs.Delete
rs.MoveNext
If rs.EOF Then
MsgBox "Não há registro, adicione algum registro!", vbExclamation
Command1_Click (0)
Else
rs.MoveLast
End If
End If


Case 2
End
End Select
End Sub

Private Sub Form_Activate()
Set rs = Data1.Recordset

If rs.RecordCount = 0 Then
MsgBox "Não há registro, adicione algum registro!", vbExclamation
Command1_Click (0)
End If
End Sub

