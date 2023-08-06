VERSION 5.00
Begin VB.Form frmDRT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de DRTs"
   ClientHeight    =   1485
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5625
   ControlBox      =   0   'False
   Icon            =   "frmCodebank.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      DataField       =   "DRT"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   900
      MaxLength       =   50
      TabIndex        =   5
      Top             =   270
      Width           =   4575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Salvar"
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   825
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Atualizar"
      Height          =   315
      Left            =   2145
      TabIndex        =   2
      Top             =   825
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Deletar"
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Top             =   825
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Adicionar"
      Height          =   315
      Left            =   195
      TabIndex        =   0
      Top             =   825
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Text;"
      DatabaseName    =   "C:\BSC"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "drt#txt"
      Top             =   1140
      Width           =   5625
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DRT:"
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   6
      Top             =   300
      Width           =   390
   End
   Begin VB.Label cmdclose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SAIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4200
      TabIndex        =   4
      Top             =   825
      Width           =   1350
   End
End
Attribute VB_Name = "frmDRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error Resume Next
If Data1.EOFAction Then
  Data1.Recordset.AddNew
    Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    End If
    
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H80000005
End Sub

Private Sub cmdclose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &HFFFF00
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
  'this may produce an error if you delete the last
  'record or the only record in the recordset
  Data1.Recordset.Delete
  Data1.Recordset.MoveLast
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H80000005
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
  'this is really only needed for multi user apps
  Data1.Refresh
End Sub

Private Sub cmdRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H80000005
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
  Unload Me
End Sub

Private Sub cmdUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H80000005
End Sub

Private Sub Command1_Click()
On Error GoTo errfind
marcador = Data1.Recordset.Bookmark
Data1.Recordset.Index = "USUARIO"

criterio = InputBox$("Nome a localizar: ", "Localizar DRT")

If criterio <> Empty Then
Data1.Recordset.Seek "=", criterio
If Data1.Recordset.NoMatch Then
MsgBox "Associado não localizado, verifique se o nome digitado é válido ou se foi digitado corretamente!", vbCritical, "Localizar Pedido"
Data1.Recordset.Bookmark = marcador
End If
Else
Data1.Recordset.Bookmark = marcador
End If
errfind:
'MsgBox Err.Description & Err.Source
Exit Sub
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H80000005
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0  'throw away the error
End Sub

Private Sub Data1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H80000005
End Sub

Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position
  'for dynasets and snapshots
  Data1.Caption = "Registro: " & (Data1.Recordset.AbsolutePosition + 1)
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

Private Sub Form_Load()
Dim criterio As Long
Dim marcador As Variant
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H80000005

End Sub

Private Sub txtFields_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.BackColor = &H80000005
End Sub
