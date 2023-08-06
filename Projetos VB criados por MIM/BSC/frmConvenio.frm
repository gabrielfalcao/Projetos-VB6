VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmConvenio 
   BackColor       =   &H00CF8D45&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convênios BSC"
   ClientHeight    =   5895
   ClientLeft      =   420
   ClientTop       =   135
   ClientWidth     =   7455
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmConvenio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00CF8D45&
      Caption         =   "Texto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   60
      TabIndex        =   16
      Top             =   1830
      Width           =   2625
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFD1AD&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   105
         ScaleHeight     =   390
         ScaleWidth      =   2400
         TabIndex        =   17
         Top             =   225
         Width           =   2430
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Salvar"
            Height          =   300
            Left            =   780
            TabIndex        =   20
            Top             =   45
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Abrir"
            Height          =   300
            Left            =   45
            TabIndex        =   19
            Top             =   45
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Imprimir"
            Height          =   300
            Left            =   1515
            TabIndex        =   18
            Top             =   45
            Width           =   825
         End
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Adicionar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3420
      TabIndex        =   15
      Top             =   1665
      Width           =   1470
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Fechar Tela"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5970
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1245
      Width           =   1305
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFAEA&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5100
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   360
      Width           =   2085
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFAEA&
      Height          =   300
      Index           =   2
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   4
      Text            =   "0.00N"
      Top             =   1335
      Width           =   3210
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFAEA&
      Height          =   300
      Index           =   0
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   3
      Top             =   495
      Width           =   3210
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   2610
      Width           =   7335
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFAEA&
      Height          =   300
      Index           =   3
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   1
      Top             =   915
      Width           =   3210
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFAEA&
      DataField       =   "DRT"
      DataSource      =   "Data1"
      Height          =   300
      Index           =   1
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   0
      Top             =   75
      Width           =   3210
   End
   Begin RichTextLib.RichTextBox abre 
      Height          =   570
      Left            =   1245
      TabIndex        =   9
      Top             =   3825
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   1005
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmConvenio.frx":0ECA
   End
   Begin MSComDlg.CommonDialog cdopen 
      Left            =   1260
      Top             =   3390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir texto:"
      FileName        =   "CamBSC"
      Filter          =   "Arquivos de Texto|*.txt"
   End
   Begin MSComDlg.CommonDialog cdsave 
      Left            =   1260
      Top             =   3390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Salvar texto como:"
      FileName        =   "CamBSC"
      Filter          =   "Arquivos de Texto|*.txt"
   End
   Begin VB.CheckBox savecode1 
      Caption         =   "Check1"
      Height          =   225
      Left            =   1185
      TabIndex        =   12
      Top             =   3555
      Width           =   1575
   End
   Begin VB.CheckBox savecode2 
      Caption         =   "Check1"
      Height          =   225
      Left            =   1185
      TabIndex        =   13
      Top             =   3555
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total atual dos valores:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   5100
      TabIndex        =   10
      Top             =   75
      Width           =   2085
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código 2:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   75
      TabIndex        =   8
      Top             =   1410
      Width           =   765
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código 1:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   540
      Width           =   765
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor(sem R$):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   5
      Top             =   975
      Width           =   1260
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DRT do Associado:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   2
      Top             =   105
      Width           =   1545
   End
   Begin VB.Menu fnd 
      Caption         =   "Localizar"
      Visible         =   0   'False
      Begin VB.Menu byname 
         Caption         =   "por Nome do Associado"
      End
      Begin VB.Menu bydrt 
         Caption         =   "por DRT do associado"
      End
   End
End
Attribute VB_Name = "frmConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim salvo As Boolean
Private Sub cmdAdd_Click()
On Error GoTo Err
Dim valor1 As Currency
Dim valor2 As Currency
If txtFields(1).Text = Empty Or Len(txtFields(1).Text) = 0 Then
MsgBox "campo de DRT vazio ou incompleto, verifique e tente novamente"
txtFields(1).SetFocus
Exit Sub
End If
savecode1.Value = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD1")
savecode2.Value = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD2")
If savecode1.Value = "1" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE1", txtFields(0).Text
End If
If savecode2.Value = "1" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE2", txtFields(2).Text
End If
'If Text2.Text = Empty Then
'Text2.Text
'valor2 = txtFields(3).Text
'End If
valor1 = Text2.Text
valor2 = txtFields(3).Text
Dim val As String
If Len(txtFields(3)) = 7 Then
val = "        "
ElseIf Len(txtFields(3)) = 6 Then
val = "         "
ElseIf Len(txtFields(3)) = 5 Then
val = "          "
ElseIf Len(txtFields(3)) = 4 Then
val = "           "
Else
txtFields(3).Text = Empty
End If
If txtFields(3).Text <> Empty Then

If Len(txtFields(1)) = 8 Then
Text1.Text = Text1.Text & txtFields(1).Text & "         " & txtFields(0).Text & val & txtFields(3).Text & "           " & txtFields(2).Text & vbCrLf
Text2.Text = valor1 + valor2
ElseIf Len(txtFields(1)) = 6 Then
Text1.Text = Text1.Text & "00" & txtFields(1).Text & "         " & txtFields(0).Text & val & txtFields(3).Text & "           " & txtFields(2).Text & vbCrLf
Text2.Text = valor1 + valor2
End If
txtFields(3).Text = Empty
txtFields(1).Text = Empty
txtFields(1).SetFocus
Else
MsgBox "Não existe valor atual ou valor é menor que o mínimo de caracteres", vbCritical, Me.Caption
txtFields(3).SetFocus
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


Private Sub cmdUpdate_Click()

On Error GoTo Err

savecode1.Value = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD1")
savecode2.Value = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD2")
If savecode1.Value = "1" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE1", txtFields(0).Text
End If
If savecode2.Value = "1" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE2", txtFields(2).Text
End If
If savecode1 = "1" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE1", txtFields(0).Text
End If
If savecode2 = "1" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE2", txtFields(2).Text
End If
cdsave.ShowSave
Open cdsave.FileName For Output As #1
Print #1, Text1.Text
Close #1
salvo = True
Err:
Exit Sub

End Sub

Private Sub cmdClose_Click()
If MsgBox("Deseja realmente fechar o programa?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then Unload Me
End Sub

Private Sub Command1_Click()
On Error GoTo erro
If salvo = False Then
Dim Resposta  As Integer
Resposta = MsgBox("Você deseja salvar o texto atual?", vbYesNo, Me.Caption)
Select Case Resposta
Case vbYes
cmdUpdate_Click
Case vbNo
Open "C:\Backup BSC.txt" For Output As #1
Print #1, Text1.Text
Close #1
Case Else
End Select
End If
cdopen.ShowOpen
abre.LoadFile cdopen.FileName
Text1.Text = abre.Text
erro:
Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo erro
Dim Resposta
Resposta = MsgBox("Você deseja salvar o texto atual?", vbYesNoCancel + vbQuestion, Me.Caption)
If Resposta = vbYes Then
cmdUpdate_Click
If Resposta = vbNo Then
Open "C:\Backup BSC" & Date & Time & ".txt" For Output As #2
Print #2, Text1.Text
Close #2

  End If
  End If
Printer.Print Text1.Text
Printer.EndDoc
erro:
Exit Sub
End Sub


Private Sub Form_Load()
On Error GoTo erro
salvo = True
Dim havesaved As Boolean
savecode1.Value = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD1")
havesaved = False
txtFields(0).Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE1")
'txtFields(2).Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE2")
  drt.Text = "00" & txtFields(1).Text
erro:
'MsgBox "Ocorreu um erro: " & Err.Description & " - Nº: " & Err.Number & " - A ação selecionada foi cancelada e os dados podem não ser carregados!"
Exit Sub
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Resposta
If salvo = False Then
Resposta = MsgBox("Você ainda não salvou o arquivo deseja salvar?", vbYesNo + vbQuestion, Me.Caption)
If Resposta = vbYes Then
cmdUpdate_Click
If Resposta = vbNo Then
Open "C:\Backup BSC.txt" For Output As #1
Print #1, Text1.Text
Close #1

  End If
  End If
  End If
  
savecode1.Value = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD1")
savecode2.Value = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD2")
If savecode1.Value = "1" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE1", txtFields(0).Text
End If
If savecode2.Value = "1" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE2", txtFields(2).Text
End If
    Unload Me
End Sub

Private Sub Text1_Change()
On Error GoTo Err
If Len(Text1) <> 0 Then salvo = False
Text1.SelStart = Len(Text1.Text)
Err:
Exit Sub
End Sub

Private Sub Text1_Click()
cmdAdd.Default = False
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Err:
Exit Sub
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Err:
Exit Sub
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Err:
Exit Sub
End Sub

Private Sub txtFields_Change(Index As Integer)
If Len(txtFields(1).Text) = 8 Then txtFields(3).SetFocus
savecode1.Value = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD1")
savecode2.Value = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD2")
If savecode1.Value = "1" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE1", txtFields(0).Text
End If
If savecode2.Value = "1" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE2", txtFields(2).Text
End If
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
Select Case Index
Case 1
cmdAdd.Default = False
Case 3
cmdAdd.Default = True
End Select
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 1
If KeyAscii = 13 Then txtFields(3).SetFocus
End Select
End Sub
