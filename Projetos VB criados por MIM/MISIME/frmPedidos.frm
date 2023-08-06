VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos"
   ClientHeight    =   6795
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmPedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Adicionar"
      Height          =   300
      Left            =   120
      TabIndex        =   38
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Deletar"
      Height          =   300
      Left            =   1200
      TabIndex        =   37
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Atualizar"
      Height          =   300
      Left            =   2280
      TabIndex        =   36
      Top             =   6120
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Pedidos"
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
      RecordSource    =   "Pedidos"
      Top             =   6450
      Width           =   5520
   End
   Begin VB.TextBox txtFields 
      DataField       =   "AlíquotaDoImpostoSobreServiços"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   18
      Left            =   2040
      TabIndex        =   35
      Top             =   5800
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Frete"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   17
      Left            =   2040
      TabIndex        =   33
      Top             =   5480
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CódigoDoMétodoDeTransporte"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   16
      Left            =   2040
      TabIndex        =   31
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DataDeEnvio"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   15
      Left            =   2040
      TabIndex        =   29
      Top             =   4840
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PaísDestino"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   13
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   26
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EstadoOuProvínciaDeDestino"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   11
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   23
      Top             =   3560
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EstadoDestino"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   10
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   21
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CidadeDestino"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   9
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   19
      Top             =   2920
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EndereçoDestino"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   8
      Left            =   2040
      MaxLength       =   255
      TabIndex        =   17
      Top             =   2600
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NomeDoDestinatário"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   7
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   15
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PrometidoPorData"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1960
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "RequeridoPorData"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1640
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NúmeroDoPedidoDeCompra"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   4
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   9
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DataDoPedido"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1000
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CódigoDoFuncionário"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   680
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CódigoCliente"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CódigoDoPedido"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   40
      Width           =   1935
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "CEPDestino"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2040
      TabIndex        =   39
      Top             =   3885
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
      DataField       =   "TelefoneTransporte"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2040
      TabIndex        =   40
      Top             =   4530
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   15
      Format          =   "(0031)0000-0000"
      Mask            =   "(    )    -    "
      PromptChar      =   "-"
   End
   Begin VB.Label lblLabels 
      Caption         =   "AlíquotaDoImpostoSobreServiços:"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   34
      Top             =   5820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Frete:"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   32
      Top             =   5500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CódigoDoMétodoDeTransporte:"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   30
      Top             =   5180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DataDeEnvio:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   28
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TelefoneTransporte:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   27
      Top             =   4540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PaísDestino:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   25
      Top             =   4220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CEPDestino:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EstadoOuProvínciaDeDestino:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   3580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EstadoDestino:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CidadeDestino:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EndereçoDestino:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NomeDoDestinatário:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PrometidoPorData:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "RequeridoPorData:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NúmeroDoPedidoDeCompra:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DataDoPedido:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CódigoDoFuncionário:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CódigoCliente:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CódigoDoPedido:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmPedidos"
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

Private Sub Form_Unload(Cancel As Integer)
Fundo.Show
Fundo.WindowState = 2
End Sub
