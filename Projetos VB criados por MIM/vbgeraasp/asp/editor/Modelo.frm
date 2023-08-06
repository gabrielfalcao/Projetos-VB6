VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmModelo 
   ClientHeight    =   3360
   ClientLeft      =   720
   ClientTop       =   780
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   4800
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtbDocumento 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5953
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Modelo.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
rtbDocumento.AutoVerbMenu = True
'Margin 20 cm
'567 twip = 1 cm
rtbDocumento.RightMargin = 11340
End Sub

Private Sub Form_Resize()
On Error Resume Next
rtbDocumento.Width = Me.ScaleWidth - 50
rtbDocumento.Height = Me.ScaleHeight - 50
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim Resposta
Resposta = MsgBox("Want to save " & Me.Caption, vbYesNoCancel + vbQuestion, "Save?")
If Resposta = vbYes Then

frmEditor.cmmArquivo.DialogTitle = "Save"
frmEditor.cmmArquivo.Filter = "Rich Text Format (*.rtf)|*.rtf|Text Files (*.txt)|*.txt|Batch Files (*.bat)|*.bat|INI Files (*.ini)|*.ini|"
frmEditor.cmmArquivo.ShowSave
On Error GoTo Erro
If frmEditor.cmmArquivo.FilterIndex = 1 Then
Salvar = frmEditor.ActiveForm.rtbDocumento.SaveFile(frmEditor.cmmArquivo.FileName, 0)
frmEditor.ActiveForm.Caption = frmEditor.cmmArquivo.FileName
Else
Salvar = frmEditor.ActiveForm.rtbDocumento.SaveFile(frmEditor.cmmArquivo.FileName, 1)
frmEditor.ActiveForm.Caption = frmEditor.cmmArquivo.FileName
End If
Exit Sub
Erro:
MsgBox "Unable to save file." & Chr(13) & "Verify Disk Space or if exists document to save."
ElseIf Resposta = vbCancel Then
Cancel = True
End If
End Sub

Private Sub rtbDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If frmEditor.cboTamanhoDaFonte.Text <> frmEditor.ActiveForm.rtbDocumento.SelFontSize Then
frmEditor.cboTamanhoDaFonte.Text = frmEditor.ActiveForm.rtbDocumento.SelFontSize
End If
End Sub

Private Sub rtbDocumento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmEditor.cboTamanhoDaFonte.Text <> frmEditor.ActiveForm.rtbDocumento.SelFontSize Then
frmEditor.cboTamanhoDaFonte.Text = frmEditor.ActiveForm.rtbDocumento.SelFontSize
End If
End Sub
