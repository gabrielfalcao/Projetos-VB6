VERSION 5.00
Begin VB.Form frmFornecedores 
   Caption         =   "Fornecedores"
   ClientHeight    =   4815
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   8745
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8745
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   0
      ScaleHeight     =   1830
      ScaleWidth      =   8745
      TabIndex        =   9
      Top             =   0
      Width           =   8745
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
         Height          =   735
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   1770
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Nome do Fornecedor"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   0
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   12
         Top             =   105
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Endereço"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   11
         Top             =   420
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Telefone"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   2
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   10
         Top             =   750
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nome do Fornecedor:"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   15
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Endereço:"
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   14
         Top             =   450
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Telefone:"
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   13
         Top             =   765
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   8745
      TabIndex        =   0
      Top             =   3315
      Width           =   8745
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Último Registro - >|"
         Height          =   300
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1605
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Próximo Registro - >"
         Height          =   300
         Left            =   3945
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1635
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "< - Registro Anterior"
         Height          =   300
         Left            =   1290
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1605
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "|< - 1º Registro"
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Remover"
         Height          =   300
         Left            =   2910
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
         Left            =   3885
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
      Height          =   825
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Fornecedores"
      Top             =   3990
      Width           =   8745
   End
End
Attribute VB_Name = "frmFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  Data1.Recordset.AddNew
    Data1.Refresh
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



Private Sub Command2_Click()
Data1.Recordset.MoveFirst
End Sub

Private Sub Command3_Click()
Data1.Recordset.MovePrevious
End Sub

Private Sub Command4_Click()
Data1.Recordset.MoveNext
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


