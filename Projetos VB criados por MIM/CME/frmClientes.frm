VERSION 5.00
Begin VB.Form frmClientes 
   Caption         =   "Controle de Clientes"
   ClientHeight    =   7455
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11355
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   0
      ScaleHeight     =   4515
      ScaleWidth      =   11355
      TabIndex        =   5
      Top             =   0
      Width           =   11355
      Begin VB.TextBox txtFields 
         DataField       =   "Nome do Cliente"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   0
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   13
         Top             =   210
         Width           =   3480
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Endereço"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   12
         Top             =   525
         Width           =   3480
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Telefone Residencial"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   2
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   11
         Top             =   855
         Width           =   3480
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Telefone Comercial"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   3
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1170
         Width           =   3480
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CEP"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   4
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1485
         Width           =   3480
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Correio Eletrônico"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   5
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1800
         Width           =   3480
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "Paga bem?"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2100
         TabIndex        =   7
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Observações"
         DataSource      =   "Data1"
         Height          =   1650
         Index           =   6
         Left            =   2100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2475
         Width           =   3495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nome do Cliente:"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Endereço:"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   555
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Telefone Residencial:"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   870
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Telefone Comercial:"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   1185
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "CEP:"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   17
         Top             =   1515
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Correio Eletrônico:"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   16
         Top             =   1815
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Paga bem?:"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   15
         Top             =   2175
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Observações:"
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   14
         Top             =   2505
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   5955
      Width           =   11355
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFC0FF&
         Caption         =   "&Fechar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   7230
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Width           =   1770
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Último Registro - >|"
         Height          =   300
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         Width           =   1605
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Próximo Registro - >"
         Height          =   300
         Left            =   3945
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         Width           =   1635
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "< - Registro Anterior"
         Height          =   300
         Left            =   1290
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   1605
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "|< - 1º Registro"
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   1965
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Remover"
         Height          =   300
         Left            =   2925
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   315
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&ReVisualizar"
         Height          =   300
         Left            =   2895
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1050
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Salvar Atual"
         Height          =   300
         Left            =   3900
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   1020
      End
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "cme.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Controle de Clientes"
      Top             =   6630
      Width           =   11355
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
If txtFields.Item(0).Text = Empty Then
MsgBox "Campos vazios, preencha o registro antes de adicionar ou salvar um novo!", vbCritical, Me.Caption
Else
  'Data1.UpdateRecord
 ' Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  Data1.Recordset.AddNew
    Data1.Refresh
    End If
End Sub

Private Sub cmdDelete_Click()
  'this may produce an error if you delete the last
  'record or the only record in the recordset
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
End Sub

Private Sub cmdRefresh_Click()
  'this is really only needed for multi user apps
  Data1.Refresh
End Sub

Private Sub cmdUpdate_Click()
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    Data1.Refresh
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Data1.Recordset.MoveFirst
End Sub

Private Sub Command3_Click()
On Error GoTo err
Data1.Recordset.MovePrevious
err:
Exit Sub
End Sub

Private Sub Command4_Click()
On Error GoTo err
Data1.Recordset.MoveNext
err:
Exit Sub
End Sub

Private Sub Command5_Click()
Data1.Recordset.MoveLast
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
Data1.Caption = "Registro: " & Data1.Recordset.RecordCount
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

Private Sub Timer1_Timer()

End Sub

'Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
 ' MsgBox "Data error event hit err:" & Description
'End Sub

'Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
'  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
'End Sub

'Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
 ' Dim bCancel As Boolean

  'Select Case adReason
  'Case adRsnAddNew
  'Case adRsnClose
  'Case adRsnDelete
  'Case adRsnFirstChange
  'Case adRsnMove
  'Case adRsnRequery
  'Case adRsnResynch
  'Case adRsnUndoAddNew
  'Case adRsnUndoDelete
  'Case adRsnUndoUpdate
  'Case adRsnUpdate
  'End Select
'
'  If bCancel Then adStatus = adStatusCancel

Private Sub Form_Load()

End Sub
