VERSION 5.00
Object = "{EDE6871F-B292-4B86-B602-523B7F4DC820}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form splash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Gerador de Bombs"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   405
      Left            =   540
      TabIndex        =   2
      Top             =   2940
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Não Concordo"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   64
      BCOLO           =   64
      FCOL            =   255
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "splash.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   450
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "splash.frx":001C
      Top             =   645
      Width           =   3900
   End
   Begin Chameleon.chameleonButton chameleonButton2 
      Height          =   405
      Left            =   2550
      TabIndex        =   3
      Top             =   2940
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Concordo"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16384
      BCOLO           =   16384
      FCOL            =   65280
      FCOLO           =   65280
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "splash.frx":0183
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   8888
      Left            =   2355
      Top             =   1665
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Height          =   225
      Left            =   4650
      TabIndex        =   5
      Top             =   3795
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leia com atenção..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   1710
      TabIndex        =   4
      Top             =   3750
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ATENÇÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1695
      TabIndex        =   0
      Top             =   105
      Width           =   1410
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nome As String

Private Sub chameleonButton1_Click()
Unload Me
End Sub

Private Sub chameleonButton2_Click()
'MsgBox "Bombs criados neste programa só funcionam perfeitamente até a versão 98SE do MS Windows", vbInformation, Me.Caption
frmBombGen.Show
Unload Me
End Sub



    


Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
Dim p_ret As String
p_ret = StrConv(LoadResData("FUNDO", "SOM"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
End Sub

Private Sub Form_Load()
If Len(App.Path) = 3 Then
nome = App.Path & "som.mp3"
Else
nome = App.Path & "\som.mp3"
End If

End Sub

Private Sub Label3_Click()
chameleonButton2.Enabled = True
End Sub

Private Sub Timer1_Timer()
chameleonButton2.Enabled = True
End Sub
