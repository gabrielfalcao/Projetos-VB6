VERSION 5.00
Begin VB.Form frmLista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "1º Arcanjo Downloader 2.0"
   ClientHeight    =   4035
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5520
   Begin VB.ListBox lNom 
      Appearance      =   0  'Flat
      BackColor       =   &H00ECEDFF&
      DataField       =   "Nome"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      ItemData        =   "frmLista.frx":0CCA
      Left            =   113
      List            =   "frmLista.frx":0CCC
      TabIndex        =   2
      Top             =   315
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   113
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1815
      Width           =   5295
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   113
      Top             =   2595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Baixar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   3195
      TabIndex        =   0
      Top             =   3450
      Width           =   2025
   End
   Begin VB.Data Data1 
      Connect         =   "Text;"
      DatabaseName    =   "E:\Gabriel\Meus Arquivos\Projetos VB criados por MIM\Downloader2\Downloader"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   113
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   0  'Table
      RecordSource    =   "ListDown#txt"
      Top             =   795
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox nom 
      DataField       =   "Nome"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   113
      TabIndex        =   7
      Top             =   465
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox descr 
      Appearance      =   0  'Flat
      DataField       =   "Descricao"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   113
      TabIndex        =   6
      Top             =   465
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ListBox ldesc 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      ItemData        =   "frmLista.frx":0CCE
      Left            =   113
      List            =   "frmLista.frx":0CD0
      TabIndex        =   4
      Top             =   945
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.ListBox lCam 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      ItemData        =   "frmLista.frx":0CD2
      Left            =   113
      List            =   "frmLista.frx":0CD4
      TabIndex        =   3
      Top             =   945
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox cami 
      DataField       =   "Caminho"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   113
      TabIndex        =   5
      Top             =   465
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   113
      TabIndex        =   11
      Top             =   1530
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programas Disponíveis para Download:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   90
      Width           =   3720
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caminho:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   113
      TabIndex        =   9
      Top             =   2745
      Width           =   900
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   113
      TabIndex        =   8
      Top             =   2985
      Width           =   5295
   End
End
Attribute VB_Name = "frmLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With frmDown2
.Caption = "Baixando " & lNom.Text
.txtURL.Text = lCam.Text
.Show
End With
Unload Me
End Sub

Private Sub Form_Activate()
Dim A As String
Dim b As String
A = (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
b = Data1.Recordset.RecordCount
For i = A To b
lNom.AddItem nom.Text
ldesc.AddItem descr.Text
lCam.AddItem cami.Text
Data1.Recordset.MoveNext
Next i
Command1.Enabled = False
lNom.ListIndex = 0
End Sub

Private Sub lNom_Click()
ldesc.ListIndex = lNom.ListIndex
lCam.ListIndex = lNom.ListIndex
Text1.Text = ldesc.Text
Label2.Caption = lCam.Text
Command1.Enabled = True
End Sub

