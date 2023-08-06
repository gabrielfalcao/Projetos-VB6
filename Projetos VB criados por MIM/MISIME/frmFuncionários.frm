VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFuncionários 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Funcionários"
   ClientHeight    =   6105
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11055
   Icon            =   "frmFuncionários.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   11055
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   5460
      ScaleHeight     =   270
      ScaleWidth      =   5220
      TabIndex        =   67
      Top             =   5025
      Width           =   5280
      Begin VB.TextBox foto 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Fotografia"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2325
         TabIndex        =   68
         Top             =   -15
         Width           =   2925
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Arquivo de Fotografia:"
         Height          =   210
         Index           =   29
         Left            =   0
         TabIndex        =   69
         Top             =   30
         Width           =   2295
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1365
      Top             =   4710
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   10935
      Top             =   3810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Procurar Foto de Funcionário"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Foto do Funcionário"
      Height          =   2100
      Left            =   9390
      TabIndex        =   65
      Top             =   1140
      Width           =   1635
      Begin VB.CommandButton Command1 
         Caption         =   "Procurar Foto"
         Height          =   330
         Left            =   195
         TabIndex        =   66
         Top             =   1710
         Width           =   1260
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1470
         Left            =   180
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Adicionar"
      Height          =   300
      Left            =   660
      TabIndex        =   61
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Deletar"
      Height          =   300
      Left            =   1725
      TabIndex        =   60
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Reamostrar"
      Height          =   300
      Left            =   2865
      TabIndex        =   59
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Caption         =   "Funcionários"
      Connect         =   "Access"
      DatabaseName    =   "E:\Gabriel\Meus Arquivos\Projetos do VB 6.0\MISIME\Misime.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Funcionários"
      Top             =   5760
      Width           =   11055
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LocalDoEscritório"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   30
      Left            =   7395
      MaxLength       =   20
      TabIndex        =   58
      Top             =   4605
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Observações"
      DataSource      =   "Data1"
      Height          =   310
      Index           =   29
      Left            =   7395
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   56
      Top             =   4275
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TelefoneDoContatoEmEmergência"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   28
      Left            =   7395
      MaxLength       =   30
      TabIndex        =   54
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ContactarEmEmergência"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   27
      Left            =   7395
      MaxLength       =   50
      TabIndex        =   52
      Top             =   3645
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NomeDoCônjuje"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   26
      Left            =   7395
      MaxLength       =   50
      TabIndex        =   50
      Top             =   3315
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CódigoDoSupervisor"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   25
      Left            =   7395
      TabIndex        =   48
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Deduções"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   24
      Left            =   7395
      TabIndex        =   46
      Top             =   2685
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TaxaDeCobrança"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   23
      Left            =   7395
      TabIndex        =   44
      Top             =   2355
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Salário"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   22
      Left            =   7395
      TabIndex        =   42
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DataDoAluguel"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   21
      Left            =   7395
      TabIndex        =   40
      Top             =   1725
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Data de Nascimento"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   20
      Left            =   7395
      TabIndex        =   38
      Top             =   1395
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CódigoDoDepartmamento"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   19
      Left            =   7395
      TabIndex        =   36
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "País"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   16
      Left            =   7395
      MaxLength       =   50
      TabIndex        =   32
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Região"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   14
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   29
      Top             =   4635
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EstadoOuProvíncia"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   13
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   27
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Cidade"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   12
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   25
      Top             =   4005
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Endereço"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   11
      Left            =   2040
      MaxLength       =   255
      TabIndex        =   23
      Top             =   3675
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ramal"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   10
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   21
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NomeNoCorreioEletrônico"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   9
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   19
      Top             =   3045
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Título"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   8
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   17
      Top             =   2715
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Sobrenome"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   7
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   15
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NomeDoMeio"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   6
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   13
      Top             =   2085
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PrimeiroNome"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   5
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   11
      Top             =   1755
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NúmeroDaCarteiraDeTrabalho"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   4
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   9
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NúmeroDoFuncionário"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   3
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   7
      Top             =   1125
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NúmeroDoINPS"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   2
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   5
      Top             =   795
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NomeDoDepartamento"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   1
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CódigoDoFuncionário"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   165
      Width           =   1935
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "CEP"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2040
      TabIndex        =   62
      Top             =   4950
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "00-000-000"
      Mask            =   "  -   -   "
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      DataField       =   "NºDoTelefone"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   7410
      TabIndex        =   63
      Top             =   450
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   15
      Format          =   "(0031)0000-0000"
      Mask            =   "(    )    -    "
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      DataField       =   "TelefoneComercial"
      DataSource      =   "Data1"
      Height          =   330
      Left            =   7410
      TabIndex        =   64
      Top             =   765
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   15
      Format          =   "(0031)0000-0000"
      Mask            =   "(    )    -    "
      PromptChar      =   "-"
   End
   Begin VB.Label lblLabels 
      Caption         =   "LocalDoEscritório:"
      Height          =   255
      Index           =   31
      Left            =   5475
      TabIndex        =   57
      Top             =   4620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Observações:"
      Height          =   255
      Index           =   30
      Left            =   5475
      TabIndex        =   55
      Top             =   4305
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TelefoneDoContatoEmEmergência:"
      Height          =   255
      Index           =   28
      Left            =   5475
      TabIndex        =   53
      Top             =   3975
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ContactarEmEmergência:"
      Height          =   255
      Index           =   27
      Left            =   5475
      TabIndex        =   51
      Top             =   3660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NomeDoCônjuje:"
      Height          =   255
      Index           =   26
      Left            =   5475
      TabIndex        =   49
      Top             =   3345
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CódigoDoSupervisor:"
      Height          =   255
      Index           =   25
      Left            =   5475
      TabIndex        =   47
      Top             =   3015
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Deduções:"
      Height          =   255
      Index           =   24
      Left            =   5475
      TabIndex        =   45
      Top             =   2700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TaxaDeCobrança:"
      Height          =   255
      Index           =   23
      Left            =   5475
      TabIndex        =   43
      Top             =   2385
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Salário:"
      Height          =   255
      Index           =   22
      Left            =   5475
      TabIndex        =   41
      Top             =   2055
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DataDoAluguel:"
      Height          =   255
      Index           =   21
      Left            =   5475
      TabIndex        =   39
      Top             =   1740
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Data de Nascimento:"
      Height          =   255
      Index           =   20
      Left            =   5475
      TabIndex        =   37
      Top             =   1425
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CódigoDoDepartmamento:"
      Height          =   255
      Index           =   19
      Left            =   5475
      TabIndex        =   35
      Top             =   1095
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TelefoneComercial:"
      Height          =   255
      Index           =   18
      Left            =   5475
      TabIndex        =   34
      Top             =   780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TelefoneResidencial:"
      Height          =   255
      Index           =   17
      Left            =   5475
      TabIndex        =   33
      Top             =   465
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "País:"
      Height          =   255
      Index           =   16
      Left            =   5475
      TabIndex        =   31
      Top             =   135
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CEP:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   30
      Top             =   4980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Região:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   28
      Top             =   4665
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EstadoOuProvíncia:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   26
      Top             =   4335
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cidade:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   4020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Endereço:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   3705
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ramal:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3375
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NomeNoCorreioEletrônico:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   3060
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Título:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2745
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sobrenome:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2415
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NomeDoMeio:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PrimeiroNome:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1785
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NúmeroDaCarteiraDeTrabalho:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1455
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NúmeroDoFuncionário:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NúmeroDoINPS:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   825
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NomeDoDepartamento:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   495
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CódigoDoFuncionário:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "frmFuncionários"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

On Error GoTo Err
  Data1.Refresh
If txtFields.Item(0).Text = "" Then

   Data1.Recordset.AddNew
     Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  cmdDelete.Enabled = True

Else
  
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  Data1.Recordset.AddNew
  End If
  Image1.Picture = Empty
Err:
  Exit Sub
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Err
Dim m As Integer
Dim a As String

If txtFields(1).Text = "" Then
a = "NÃO EXISTE REGISTRO"
Else
a = txtFields(1).Text
End If
m = MsgBox("Tem certeza que deseja excluir o registro: " & a & " ?", vbYesNo, "!ATENÇÃO OPERADOR!")
Select Case m
Case vbYes
If txtFields.Item(0).Text = "" Then
MsgBox "Não há nenhum registro a apagar, crie um novo registro e\ou preencha os registros em branco clicando na GRADE DE LANÇAMENTOS para que possa fazer novas operações", vbCritical, Me.Caption
cmdDelete.Enabled = False
cmdAdd.Enabled = False
Lista.Show
Unload Me

Else
If txtFields.Item(0).Text = "" Then
   MsgBox "Não há nenhum registro a apagar, crie um novo registro e\ou preencha os registros em branco clicando na GRADE DE LANÇAMENTOS para que possa fazer novas operações", vbCritical, Me.Caption
Else
   Data1.Recordset.Delete
Data1.Recordset.MoveLast
  End If
  End If
  'this may produce an error if you delete the last
  'record or the only record in the recordset
Case Else
End Select
Err:
Exit Sub
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
  'this is really only needed for multi user apps
  Data1.Refresh
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
Com.Filter = "Arquivos de Imagem|*.bmp;*.jpg;*.jpeg;*.gif;*.tif;*.tga"
Com.ShowOpen
foto.Text = Com.FileName
Image1.Picture = LoadPicture(foto.Text)
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0  'throw away the error
  
End Sub

Private Sub Data1_Reposition()

  On Error Resume Next
  'This will display the current record position
  'for dynasets and snapshots
  Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
  'for the table object you must set the index property when
  'the recordset gets created and use the following line
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
Image1.Picture = LoadPicture(foto.Text)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select

End Sub

Private Sub Form_Load()

Timer1.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Fundo.Show
Fundo.WindowState = 2
End Sub

Private Sub oleFields_DblClick(Index As Integer)

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Image1.Picture = LoadPicture(foto.Text)
Timer1.Enabled = True
End Sub
