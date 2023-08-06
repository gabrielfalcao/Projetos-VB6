VERSION 5.00
Begin VB.Form frmAgenda 
   Caption         =   "Agenda FoxPro Versão 2.0"
   ClientHeight    =   2985
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   5520
   Icon            =   "frmAgenda.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   4440
      TabIndex        =   14
      Top             =   2340
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   3360
      TabIndex        =   13
      Top             =   2340
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   300
      Left            =   2280
      TabIndex        =   12
      Top             =   2340
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   1200
      TabIndex        =   11
      Top             =   2340
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   2340
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "FoxPro 2.0;"
      DatabaseName    =   "c:\teste"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "AGENDAFOX"
      Top             =   2640
      Width           =   5520
   End
   Begin VB.TextBox txtFields 
      DataField       =   "OBSERVACOE"
      DataSource      =   "Data1"
      Height          =   975
      Index           =   4
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1310
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EMAIL"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1000
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TELEFONE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   5
      Top             =   680
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ENDERECO"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NOME"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   1
      Top             =   40
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Observações:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EMAIL:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Telefone:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Endereço:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim backed As String

On Error GoTo Err
  Data1.Refresh
If txtFields.Item(0).Text = "" Then
backed = " "
  Data1.Recordset.AddNew
  txtFields.Item(1).Text = backed
  txtFields.Item(0).Text = backed
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

Private Sub cmdFind_Click()
PopupMenu fnd
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo Err
  'this is really only needed for multi user apps
  Data1.Refresh
Err:
Exit Sub

End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Err
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Err:
Exit Sub

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

Private Sub Data1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Data1.Caption = "Registro atual: " & Data1.Recordset.RecordCount
End Sub

Private Sub Data1_Reposition()

  On Error Resume Next
  'This will display the current record position
  'for dynasets and snapshots
    Data1.Caption = "Total Cadastrado: " & Data1.Recordset.RecordCount
  'for the table object you must set the index property when
  'the recordset gets created and use the following line
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Select Case Action
    Case vbDataActionMoveFirst
Data1.Caption = "Total Cadastrado: " & Data1.Recordset.RecordCount
    Case vbDataActionMovePrevious
Data1.Caption = "Total Cadastrado: " & Data1.Recordset.RecordCount
    Case vbDataActionMoveNext
Data1.Caption = "Total Cadastrado: " & Data1.Recordset.RecordCount
    Case vbDataActionMoveLast
Data1.Caption = "Total Cadastrado: " & Data1.Recordset.RecordCount
    Case vbDataActionAddNew
Data1.Caption = "Total Cadastrado: " & Data1.Recordset.RecordCount
    Case vbDataActionUpdate
Data1.Caption = "Total Cadastrado: " & Data1.Recordset.RecordCount
    Case vbDataActionDelete
Data1.Caption = "Total Cadastrado: " & Data1.Recordset.RecordCount
    Case vbDataActionFind
Data1.Caption = "Total Cadastrado: " & Data1.Recordset.RecordCount
    Case vbDataActionBookmark
Data1.Caption = "Total Cadastrado: " & Data1.Recordset.RecordCount
    Case vbDataActionClose
Data1.Caption = "Total Cadastrado: " & Data1.Recordset.RecordCount
  End Select

End Sub

