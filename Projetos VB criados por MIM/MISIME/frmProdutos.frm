VERSION 5.00
Begin VB.Form frmProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtos"
   ClientHeight    =   4635
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmProdutos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Adicionar"
      Height          =   300
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Deletar"
      Height          =   300
      Left            =   1200
      TabIndex        =   25
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Reamostrar"
      Height          =   300
      Left            =   2280
      TabIndex        =   24
      Top             =   3960
      Width           =   1230
   End
   Begin VB.Data Data1 
      Caption         =   "Produtos"
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
      RecordSource    =   "Produtos"
      Top             =   4290
      Width           =   5520
   End
   Begin VB.TextBox txtFields 
      DataField       =   "HoraPrecedente"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   10
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   23
      Top             =   3560
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Descontinuado"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2040
      TabIndex        =   21
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Nível de Estoque"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   9
      Left            =   2040
      TabIndex        =   19
      Top             =   2920
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PreçoUnitário"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   2600
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "UnidadesPedidas"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "UnidadesEmEstoque"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1960
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NúmeroDeSérie"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   5
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   11
      Top             =   1640
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CódigoDoFornecedor"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CódigoDaCategoria"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1000
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DescriçãoDoProduto"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   2
      Left            =   2040
      MaxLength       =   255
      TabIndex        =   5
      Top             =   680
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NomeDoProduto"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   1
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CódigoDoProduto"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   40
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "HoraPrecedente:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   3580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descontinuado:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nível de Estoque:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PreçoUnitário:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "UnidadesPedidas:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "UnidadesEmEstoque:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NúmeroDeSérie:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CódigoDoFornecedor:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CódigoDaCategoria:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DescriçãoDoProduto:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NomeDoProduto:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CódigoDoProduto:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmProdutos"
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
On Error GoTo Err
  'this is really only needed for multi user apps
  Data1.Refresh
Err:
Exit Sub

End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next

End Sub

Private Sub cmdClose_Click()
On Error Resume Next
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

