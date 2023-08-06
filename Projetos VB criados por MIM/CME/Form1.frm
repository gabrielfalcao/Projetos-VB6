VERSION 5.00
Begin VB.Form OS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordem de Serviço 1.00"
   ClientHeight    =   5985
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8385
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   750
      Left            =   6465
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2310
      Width           =   1755
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3750
      Picture         =   "Form1.frx":3989
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3180
      UseMaskColor    =   -1  'True
      Width           =   1800
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1868
      ScaleHeight     =   390
      ScaleWidth      =   2595
      TabIndex        =   24
      Top             =   90
      Width           =   2625
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Nº da Ordem"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1350
         TabIndex        =   25
         Text            =   "15"
         Top             =   -15
         Width           =   1260
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº da Ordem:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   26
         Top             =   90
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados do serviço"
      Height          =   1485
      Left            =   3255
      TabIndex        =   18
      Top             =   4035
      Width           =   4920
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEDEE&
         DataField       =   "Preço"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Top             =   1095
         Width           =   3135
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEDEE&
         DataField       =   "Serviço(s)"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Top             =   735
         Width           =   3135
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEDEE&
         DataField       =   "Áreas de aplicação"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   345
         Width           =   3135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço:"
         Height          =   195
         Left            =   1020
         TabIndex        =   29
         Top             =   1155
         Width           =   465
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serviços:"
         Height          =   195
         Left            =   900
         TabIndex        =   23
         Top             =   795
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Áreas de Aplicação:"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   405
         Width           =   1425
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data e Hora do Serviço"
      Height          =   1335
      Left            =   150
      TabIndex        =   13
      Top             =   3270
      Width           =   2835
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEDEE&
         DataField       =   "Data"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   870
         TabIndex        =   15
         Top             =   750
         Width           =   1410
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEDEE&
         DataField       =   "Horário"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   870
         TabIndex        =   14
         Top             =   390
         Width           =   1410
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   450
         TabIndex        =   17
         Top             =   795
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horário:"
         Height          =   195
         Left            =   285
         TabIndex        =   16
         Top             =   435
         Width           =   540
      End
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C8D3D7&
      Connect         =   "Access"
      DatabaseName    =   "C:\CME\os.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "os"
      Top             =   5670
      Width           =   8385
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Cliente"
      Height          =   2295
      Left            =   195
      TabIndex        =   0
      Top             =   675
      Width           =   5925
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEDEE&
         DataField       =   "Telefone"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2415
         TabIndex        =   6
         Top             =   1635
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEDEE&
         DataField       =   "Bairro"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3900
         TabIndex        =   5
         Top             =   1245
         Width           =   1650
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEDEE&
         DataField       =   "Complemento"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2100
         TabIndex        =   4
         Top             =   1245
         Width           =   1110
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEDEE&
         DataField       =   "Nº"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4785
         TabIndex        =   3
         Top             =   855
         Width           =   765
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEDEE&
         DataField       =   "Endereço"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1185
         TabIndex        =   2
         Top             =   855
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEDEE&
         DataField       =   "Cliente"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1035
         TabIndex        =   1
         Top             =   450
         Width           =   3135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   1620
         TabIndex        =   12
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   3330
         TabIndex        =   11
         Top             =   1305
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento:"
         Height          =   195
         Left            =   975
         TabIndex        =   10
         Top             =   1305
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº:"
         Height          =   195
         Left            =   4440
         TabIndex        =   9
         Top             =   915
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   330
         TabIndex        =   8
         Top             =   915
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   420
         TabIndex        =   7
         Top             =   495
         Width           =   525
      End
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ReAmostrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   270
      Left            =   6465
      MousePointer    =   1  'Arrow
      TabIndex        =   34
      Top             =   1485
      Width           =   1755
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Localizar registro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   270
      Left            =   6465
      MousePointer    =   1  'Arrow
      TabIndex        =   33
      Top             =   1755
      Width           =   1755
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Salvar registro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   6465
      MousePointer    =   1  'Arrow
      TabIndex        =   32
      Top             =   1215
      Width           =   1755
   End
   Begin VB.Label command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Novo registro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   6465
      MousePointer    =   1  'Arrow
      TabIndex        =   31
      Top             =   945
      Width           =   1755
   End
   Begin VB.Label command2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deletar registro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   270
      Left            =   6465
      MousePointer    =   1  'Arrow
      TabIndex        =   30
      Top             =   2025
      Width           =   1755
   End
   Begin VB.Menu teg 
      Caption         =   "Operações de Registro"
      Begin VB.Menu add 
         Caption         =   "Adicionar Registro"
      End
      Begin VB.Menu save 
         Caption         =   "Salvar Registro"
      End
      Begin VB.Menu refresh 
         Caption         =   "ReVisualizar Registros"
      End
      Begin VB.Menu killreg 
         Caption         =   "Deletar Registro"
      End
      Begin VB.Menu findreg 
         Caption         =   "Localizar Registro"
      End
   End
   Begin VB.Menu unloadform 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
Dim backed As String
On Error GoTo Err
  Data1.refresh
If Text12.Text = "" Then
  Data1.Recordset.AddNew
  Text12.Text = backed
  Text1.Text = backed
   Data1.Recordset.AddNew
     Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Else
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  Data1.Recordset.AddNew
  End If
Err:
  Exit Sub
End Sub

Private Sub byname_Click()
Dim criterio As Long
Dim marcador As Variant
marcador = Data1.Recordset.Bookmark
Data1.Recordset.Index = "Cliente"

criterio = InputBox$("Nome do cliente a localizar: ", "Localizar Pedido")

If criterio <> Empty Then
Data1.Recordset.Seek "=", criterio
If Data1.Recordset.NoMatch Then
MsgBox "Pedido não localizado, verifique se o nome é válido!", vbCritical, "Localizar Pedido"
Data1.Recordset.Bookmark = marcador
End If
Else
Data1.Recordset.Bookmark = marcador
End If
End Sub

Private Sub Command1_Click()
Dim backed As String
On Error GoTo Err
  Data1.refresh
If Text12.Text = "" Then
  Data1.Recordset.AddNew
  Text12.Text = backed
  Text1.Text = backed
   Data1.Recordset.AddNew
     Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Else
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  Data1.Recordset.AddNew
  End If
Err:
  Exit Sub
End Sub

Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.FontUnderline = True
command2.FontUnderline = False
Label13.FontUnderline = False
Label14.FontUnderline = False
Label15.FontUnderline = False
End Sub

Private Sub Command2_Click()
On Error GoTo Err
Dim m As Integer
Dim a As String

If Data1.Recordset.RecordCount <= 1 Then
a = "NÃO EXISTE REGISTRO"
Else
a = Text12.Text
End If
m = MsgBox("Tem certeza que deseja excluir o registro: " & a & " ?", vbYesNo, "!ATENÇÃO OPERADOR!")
Select Case m
Case vbYes
If Data1.Recordset.RecordCount <= 1 Then
MsgBox "Não há nenhum registro a apagar ou existe apenas 1 registro, verifique se há a quantidade mínima de 2 registros para que possa haver a remoção!", vbCritical, Me.Caption
cmdDelete.Enabled = False
cmdAdd.Enabled = False
Lista.Show
Unload Me

Else
If Data1.Recordset.RecordCount <= 1 Then
   MsgBox "Não há nenhum registro a apagar ou existe apenas 1 registro, verifique se há a quantidade mínima de 2 registros para que possa haver a remoção!", vbCritical, Me.Caption
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

Private Sub command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
command2.FontUnderline = True
Command1.FontUnderline = False
Label13.FontUnderline = False
Label14.FontUnderline = False
Label15.FontUnderline = False
End Sub

Private Sub Command3_Click()
Unload Me
frmPreview.Show
End Sub

Private Sub Command4_Click()
frmAbout.Show
End Sub

Private Sub Data1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Data1.Caption = "Registro atual: " & Data1.Recordset.RecordCount
End Sub

Private Sub Data1_Reposition()

  On Error Resume Next
  'This will display the current record position
  'for dynasets and snapshots
  Data1.Caption = "Total de registros: " & Data1.Recordset.RecordCount
  'for the table object you must set the index property when
  'the recordset gets created and use the following line
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, save As Integer)
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

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0  'throw away the error
End Sub

Private Sub findreg_Click()
On Error Resume Next
Dim criterio As Long
Dim marcador As Variant
marcador = Data1.Recordset.Bookmark
Data1.Recordset.Index = "PrimaryKey"

criterio = InputBox$("Nº da ordem de serviço a localizar: ", "Localizar Pedido")

If criterio <> Empty Then
Data1.Recordset.Seek "=", criterio
If Data1.Recordset.NoMatch Then
MsgBox "Pedido não localizado, verifique se o nº é válido!", vbCritical, "Localizar Pedido"
Data1.Recordset.Bookmark = marcador
End If
Else
Data1.Recordset.Bookmark = marcador
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.FontUnderline = False
Label13.FontUnderline = False
Label14.FontUnderline = False
Label15.FontUnderline = False
command2.FontUnderline = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Err


Dim mgs As Integer
mgs = MsgBox("Deseja salvar os registros atuais?", vbYesNo + vbQuestion, Me.Caption)
Select Case mgs
Case vbYes
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  Case Else
End Select
Err:
If Err.Number <> 0 Then
MsgBox "Não foi possível salvar, verifique a integridade do programa, do banco de dados, ou contate o distribuidor do programa!", vbCritical, "Erro de Gravação"
End If
Exit Sub
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.FontUnderline = False
Label13.FontUnderline = False
Label14.FontUnderline = False
Label15.FontUnderline = False
command2.FontUnderline = False

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.FontUnderline = False
Label13.FontUnderline = False
Label14.FontUnderline = False
Label15.FontUnderline = False
command2.FontUnderline = False

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.FontUnderline = False
Label13.FontUnderline = False
Label14.FontUnderline = False
Label15.FontUnderline = False
command2.FontUnderline = False

End Sub

Private Sub killreg_Click()
On Error GoTo Err
Dim m As Integer
Dim a As String

If Data1.Recordset.RecordCount <= 1 Then
a = "NÃO EXISTE REGISTRO"
Else
a = Text12.Text
End If
m = MsgBox("Tem certeza que deseja excluir o registro: " & a & " ?", vbYesNo, "!ATENÇÃO OPERADOR!")
Select Case m
Case vbYes
If Data1.Recordset.RecordCount <= 1 Then
MsgBox "Não há nenhum registro a apagar ou existe apenas 1 registro, verifique se há a quantidade mínima de 2 registros para que possa haver a remoção!", vbCritical, Me.Caption
cmdDelete.Enabled = False
cmdAdd.Enabled = False
Lista.Show
Unload Me

Else
If Data1.Recordset.RecordCount <= 1 Then
   MsgBox "Não há nenhum registro a apagar ou existe apenas 1 registro, verifique se há a quantidade mínima de 2 registros para que possa haver a remoção!", vbCritical, Me.Caption
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


Private Sub numberped_Click()

End Sub

Private Sub Label13_Click()
On Error GoTo Err
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Err:
If Error <> 0 Then
Exit Sub
End If
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.FontUnderline = True
Command1.FontUnderline = False
command2.FontUnderline = False
Label14.FontUnderline = False
Label15.FontUnderline = False
End Sub

Private Sub Label14_Click()
On Error Resume Next
Dim criterio As Long
Dim marcador As Variant
marcador = Data1.Recordset.Bookmark
Data1.Recordset.Index = "PrimaryKey"

criterio = InputBox$("Nº da ordem de serviço a localizar: ", "Localizar Pedido")

If criterio <> Empty Then
Data1.Recordset.Seek "=", criterio
If Data1.Recordset.NoMatch Then
MsgBox "Pedido não localizado, verifique se o nº é válido!", vbCritical, "Localizar Pedido"
Data1.Recordset.Bookmark = marcador
End If
Else
Data1.Recordset.Bookmark = marcador
End If

End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.FontUnderline = True
Command1.FontUnderline = False
Label13.FontUnderline = False
command2.FontUnderline = False
Label15.FontUnderline = False
End Sub

Private Sub Label15_Click()
On Error GoTo Err
  'this is really only needed for multi user apps
  Data1.refresh
Err:
Exit Sub
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.FontUnderline = True
Command1.FontUnderline = False
Label13.FontUnderline = False
Label14.FontUnderline = False
command2.FontUnderline = False
End Sub

Private Sub refresh_Click()
On Error GoTo Err
  'this is really only needed for multi user apps
  Data1.refresh
Err:
Exit Sub
End Sub

Private Sub save_Click()
On Error GoTo Err
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Err:
If Error <> 0 Then
Exit Sub
End If
End Sub

Private Sub unloadform_Click()
Unload Me
End Sub
