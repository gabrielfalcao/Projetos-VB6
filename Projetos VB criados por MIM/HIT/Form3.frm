VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizador de Conteúdo"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6990
   Begin VB.CommandButton Command1 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   2948
      TabIndex        =   2
      Top             =   3975
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3525
      Left            =   23
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   405
      Width           =   6945
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   23
      TabIndex        =   1
      Top             =   30
      Width           =   6945
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tip As String
Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Command2_Click()

End Sub


Private Sub Form_Activate()
If Mid(Text1, 1, 4) = "RIFF" Then tip = "Tipo: Arquivo de Mídia"
If Mid(Text1, 1, 5) = "GIF89" Then tip = "Tipo: Arquivo GIF!"
If Mid(Text1, 1, 2) = "PK" Then tip = "Tipo: Arquivo ZIP!"
If Right$(Label1.Caption, 1) <> "!" Then Label1.Caption = Label1.Caption & tip
End Sub

