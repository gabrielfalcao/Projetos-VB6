VERSION 5.00
Begin VB.Form duelo 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assistente de Duelos 1.3 - :: Duelo em Curso::"
   ClientHeight    =   2925
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8265
   Icon            =   "duelo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   8265
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4170
      Top             =   2280
   End
   Begin VB.Timer sorteador 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   4290
      Top             =   945
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3495
      Top             =   960
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0073F780&
      Caption         =   "Sortear quem começa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2805
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   540
      Width           =   2490
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2385
      Left            =   315
      ScaleHeight     =   2325
      ScaleWidth      =   2445
      TabIndex        =   7
      Top             =   300
      Width           =   2505
      Begin VB.Frame d1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Duelista 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2280
         Left            =   60
         TabIndex        =   8
         Top             =   0
         Width           =   2340
         Begin VB.CommandButton Command4 
            Caption         =   "OK"
            Height          =   345
            Left            =   1590
            TabIndex        =   17
            Top             =   1740
            Width           =   465
         End
         Begin VB.TextBox pv1a 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   360
            Left            =   255
            TabIndex        =   16
            Top             =   1725
            Width           =   1335
         End
         Begin VB.TextBox pv1s 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   360
            Left            =   270
            TabIndex        =   10
            Top             =   1155
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "OK"
            Height          =   345
            Left            =   1605
            TabIndex        =   9
            Top             =   1170
            Width           =   465
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qntdade a aumentar:"
            Height          =   210
            Left            =   420
            TabIndex        =   18
            Top             =   1515
            Width           =   1515
         End
         Begin VB.Label pv1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "8000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   270
            TabIndex        =   13
            Top             =   525
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pontos de Vida:"
            Height          =   210
            Left            =   555
            TabIndex        =   12
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qntdade a subtrair:"
            Height          =   210
            Left            =   435
            TabIndex        =   11
            Top             =   945
            Width           =   1395
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2355
      Left            =   5310
      ScaleHeight     =   2295
      ScaleWidth      =   2385
      TabIndex        =   0
      Top             =   315
      Width           =   2445
      Begin VB.Frame d2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Duelista 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2250
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2340
         Begin VB.CommandButton Command5 
            Caption         =   "OK"
            Height          =   345
            Left            =   1575
            TabIndex        =   20
            Top             =   1755
            Width           =   465
         End
         Begin VB.TextBox pv2a 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   19
            Top             =   1740
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "OK"
            Height          =   345
            Left            =   1620
            TabIndex        =   3
            Top             =   1170
            Width           =   465
         End
         Begin VB.TextBox pv2s 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   360
            Left            =   270
            TabIndex        =   2
            Top             =   1155
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qntdade a aumentar:"
            Height          =   210
            Left            =   405
            TabIndex        =   21
            Top             =   1530
            Width           =   1515
         End
         Begin VB.Label pv2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "8000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   270
            TabIndex        =   6
            Top             =   525
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pontos de Vida:"
            Height          =   210
            Left            =   570
            TabIndex        =   5
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qntdade a subtrair:"
            Height          =   210
            Left            =   435
            TabIndex        =   4
            Top             =   945
            Width           =   1395
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3225
      Top             =   1650
   End
   Begin VB.Label Label8 
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1020
      Left            =   3375
      TabIndex        =   22
      Top             =   1215
      Width           =   2040
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Programador: Gabriel Falcão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2655
      TabIndex        =   14
      Top             =   -15
      Width           =   2850
   End
   Begin VB.Menu changename 
      Caption         =   "Escolher nome dos duelistas"
   End
   Begin VB.Menu abt 
      Caption         =   "Sobre o Assistente"
   End
   Begin VB.Menu sai 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "duelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub abt_Click()
sobre.Show
End Sub

Private Sub changename_Click()
inicio.Show
End Sub

Private Sub Command1_Click()
pv1.Caption = pv1.Caption - pv1s.Text

End Sub

Private Sub Command2_Click()
pv2.Caption = pv2.Caption - pv2s.Text

End Sub

Private Sub Command3_Click()
If Command3.Caption = "Parar sorteio" Then
sorteador.Enabled = False
Timer3.Enabled = True
End If
If Command3.Caption = "Sortear quem começa" Then
sorteador.Enabled = True

Command3.Caption = "Parar sorteio"
End If


End Sub

Private Sub Command4_Click()
pv1.Caption = pv1.Caption + pv1a.Text

End Sub

Private Sub Command5_Click()

pv2.Caption = pv2.Caption + pv2a.Text

End Sub

Private Sub escff_Click()
escFun.Show
End Sub

Private Sub Form_Load()
pv1.Caption = 8000
pv2.Caption = 8000
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_Click()

End Sub

Private Sub sai_Click()
Unload Me
End
End Sub

Private Sub sorteador_Timer()
Dim primeiro As String
Select Case (1 + Int(Rnd() * 11))

    Case 1
        Me.Caption = "Assistente de Duelos 1.3 - :: Duelo em Curso:: - Primeiro a jogar: " & d1.Caption
        primeiro = d1.Caption
        
    Case 2
        Me.Caption = "Assistente de Duelos 1.3 - :: Duelo em Curso:: - Primeiro a jogar: " & d2.Caption
        primeiro = d2.Caption
        
    Case 3
        Me.Caption = "Assistente de Duelos 1.3 - :: Duelo em Curso:: - Primeiro a jogar: " & d1.Caption
        primeiro = d1.Caption
        
    Case 4
        Me.Caption = "Assistente de Duelos 1.3 - :: Duelo em Curso:: - Primeiro a jogar: " & d2.Caption
        primeiro = d2.Caption


    Case 5
        Me.Caption = "Assistente de Duelos 1.3 - :: Duelo em Curso:: - Primeiro a jogar: " & d1.Caption
        primeiro = d1.Caption
        
    Case 6
        Me.Caption = "Assistente de Duelos 1.3 - :: Duelo em Curso:: - Primeiro a jogar: " & d2.Caption
        primeiro = d2.Caption


    Case 7
        Me.Caption = "Assistente de Duelos 1.3 - :: Duelo em Curso:: - Primeiro a jogar: " & d1.Caption
        primeiro = d1.Caption
        
    Case 8
        Me.Caption = "Assistente de Duelos 1.3 - :: Duelo em Curso:: - Primeiro a jogar: " & d2.Caption
        primeiro = d2.Caption


    Case 9
        Me.Caption = "Assistente de Duelos 1.3 - :: Duelo em Curso:: - Primeiro a jogar: " & d1.Caption
        primeiro = d1.Caption
        
    Case 10
        Me.Caption = "Assistente de Duelos 1.3 - :: Duelo em Curso:: - Primeiro a jogar: " & d2.Caption
        primeiro = d2.Caption
    Case 11
        Me.Caption = "Assistente de Duelos 1.3 - :: Duelo em Curso:: - Primeiro a jogar: " & d1.Caption
        primeiro = d1.Caption
End Select
'Command3.Enabled = False

End Sub

Private Sub Timer1_Timer()
Dim intResp As Integer
If pv1.Caption <= 0 Then
intResp = MsgBox("Vencedor: " & d2.Caption & "  Perdedor: " & d1.Caption & "   Deseja começar outro duelo?", vbYesNo, "!!!Fim do Duelo!!!")
Timer1.Enabled = False

End If
If pv2.Caption <= 0 Then
intResp = MsgBox("Vencedor: " & d1.Caption & "  Perdedor: " & d2.Caption & "   Deseja começar outro duelo?", vbYesNo, "!!!Fim do Duelo!!!")
Timer1.Enabled = False

End If
Select Case intResp
Case vbYes
Dim a As String
If Len(App.Path) = 3 Then
a = App.Path & App.EXEName
Else
a = App.Path & "\" & App.EXEName
End If
Call Shell(a, vbNormalFocus)
Unload Me
Case vbNo
Unload Me
Case Else
End Select


End Sub

Private Sub Timer2_Timer()
If pv1.Caption > 8000 Then
pv1.Caption = 8000
End If
If pv2.Caption > 8000 Then
pv2.Caption = 8000
End If

End Sub

Private Sub Timer3_Timer()
If sorteador.Enabled = False Then
Command3.Caption = "Sortear quem começa"
Command3.Enabled = False
End If
Timer3.Enabled = False

End Sub
