VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFormatarFontes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formatar Fontes"
   ClientHeight    =   5970
   ClientLeft      =   3105
   ClientTop       =   990
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5055
   Begin VB.PictureBox Cor 
      Height          =   255
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   39
      Top             =   6480
      Width           =   375
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   3840
      TabIndex        =   33
      Top             =   5400
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cmmCor 
      Left            =   1440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   32
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   31
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Colors:"
      Height          =   1575
      Left            =   2520
      TabIndex        =   14
      Top             =   2640
      Width           =   2295
      Begin VB.PictureBox picPretoClaro 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   30
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picPretoEscuro 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   29
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picOutras 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   720
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   28
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picRoxoClaro 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   26
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picAzulClaro 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   25
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picVerdeClaro 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   24
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picAmareloClaro 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   23
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picVermelhoClaro 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   22
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picLaranjaClaro 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   21
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picRoxoEscuro 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   20
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picAzulEscuro 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   19
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picVerdeEscuro 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picAmareloEscuro 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picVermelhoEscuro 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   16
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picLaranjaEscuro 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Others:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sample:"
      Height          =   1455
      Left            =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   3495
      Begin VB.Label lblExemplo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sample"
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Effects:"
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   2175
      Begin VB.CheckBox chkRiscado 
         Caption         =   "&Riscado"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox chkSublinhado 
         Caption         =   "&Sublinhado"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.ListBox lstTamanhoDaFonte 
      Height          =   1425
      Left            =   4200
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtTamanhoDaFonte 
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Text            =   "10"
      Top             =   600
      Width           =   615
   End
   Begin VB.ListBox lstEstiloDaFonte 
      Height          =   1425
      ItemData        =   "FormatarFontes.frx":0000
      Left            =   2520
      List            =   "FormatarFontes.frx":0010
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtEstiloDaFonte 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Text            =   "Normal"
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox lstFontes 
      Height          =   1425
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtFontes 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Riscado 
      Caption         =   "Label5"
      Height          =   255
      Left            =   3480
      TabIndex        =   38
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Sublinhado 
      Caption         =   "Label5"
      Height          =   255
      Left            =   2040
      TabIndex        =   37
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label TamanhoDaFonte 
      Caption         =   "Label5"
      Height          =   255
      Left            =   2040
      TabIndex        =   36
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Estilo 
      Caption         =   "Label5"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Fonte 
      Caption         =   "Label5"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Size:"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Style:"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Font:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmFormatarFontes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkRiscado_Click()
If chkRiscado.Value = 1 Then
lblExemplo.FontStrikethru = True
Else
lblExemplo.FontStrikethru = False
End If
End Sub

Private Sub chkRiscado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "I checked, the text selected is StrikeThru"
End Sub

Private Sub chkSublinhado_Click()
If chkSublinhado.Value = 1 Then
lblExemplo.FontUnderline = True
Else
lblExemplo.FontUnderline = False
End If
End Sub

Private Sub chkSublinhado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "If checked, the text selected is underline"
End Sub

Private Sub cmdAplicar_Click()
On Error Resume Next
frmEditor.ActiveForm.rtbDocumento.SelFontName = lblExemplo.FontName
If lblExemplo.FontBold = True Then
frmEditor.ActiveForm.rtbDocumento.SelBold = True
Else
frmEditor.ActiveForm.rtbDocumento.SelBold = False
End If
If lblExemplo.FontItalic = True Then
frmEditor.ActiveForm.rtbDocumento.SelItalic = True
Else
frmEditor.ActiveForm.rtbDocumento.SelItalic = False
End If
If lblExemplo.FontUnderline = True Then
frmEditor.ActiveForm.rtbDocumento.SelUnderline = True
Else
frmEditor.ActiveForm.rtbDocumento.SelUnderline = False
End If
If lblExemplo.FontStrikethru = True Then
frmEditor.ActiveForm.rtbDocumento.SelStrikeThru = True
Else
frmEditor.ActiveForm.rtbDocumento.SelStrikeThru = False
End If
frmEditor.ActiveForm.rtbDocumento.SelFontSize = lblExemplo.FontSize
frmEditor.cboTamanhoDaFonte.Text = lblExemplo.FontSize
frmEditor.ActiveForm.rtbDocumento.SelColor = lblExemplo.ForeColor


End Sub

Private Sub cmdCancelar_Click()
txtFontes.Text = Fonte.Caption
txtEstiloDaFonte.Text = Estilo.Caption
txtTamanhoDaFonte.Text = Int(TamanhoDaFonte.Caption)
chkSublinhado.Value = Sublinhado.Caption
chkRiscado.Value = Riscado.Caption
lblExemplo.ForeColor = Cor.BackColor
cmdAplicar_Click
Unload Me
End Sub

Private Sub cmdOK_Click()
cmdAplicar_Click
Unload Me
End Sub

Private Sub Form_Load()

Dim Contador
For Contador = 1 To Screen.FontCount - 1
lstFontes.AddItem Screen.Fonts(Contador)
Next

Dim Tamanho
For Tamanho = 6 To 120
lstTamanhoDaFonte.AddItem (Tamanho)
Next

On Error Resume Next
txtFontes.Text = frmEditor.ActiveForm.rtbDocumento.SelFontName
lstFontes.Text = txtFontes.Text

If frmEditor.ActiveForm.rtbDocumento.SelItalic = True And frmEditor.ActiveForm.rtbDocumento.SelBold = True Then
txtEstiloDaFonte.Text = "Bold Italic"
lstEstiloDaFonte.Text = txtEstiloDaFonte.Text
ElseIf frmEditor.ActiveForm.rtbDocumento.SelItalic = True And frmEditor.ActiveForm.rtbDocumento.SelBold = False Then
txtEstiloDaFonte.Text = "Italic"
lstEstiloDaFonte.Text = txtEstiloDaFonte.Text
ElseIf frmEditor.ActiveForm.rtbDocumento.SelItalic = False And frmEditor.ActiveForm.rtbDocumento.SelBold = True Then
txtEstiloDaFonte.Text = "Bold"
lstEstiloDaFonte.Text = txtEstiloDaFonte.Text
Else
txtEstiloDaFonte.Text = "Normal"
lstEstiloDaFonte.Text = txtEstiloDaFonte.Text
End If

txtTamanhoDaFonte.Text = frmEditor.ActiveForm.rtbDocumento.SelFontSize

If frmEditor.ActiveForm.rtbDocumento.SelUnderline = True Then
chkSublinhado.Value = 1
End If

If frmEditor.ActiveForm.rtbDocumento.SelStrikeThru = True Then
chkRiscado.Value = 1
End If

picOutras.BackColor = frmEditor.ActiveForm.rtbDocumento.SelColor

Fonte.Caption = txtFontes.Text
Estilo.Caption = txtEstiloDaFonte.Text
TamanhoDaFonte.Caption = txtTamanhoDaFonte.Text
Sublinhado.Caption = chkSublinhado.Value
Riscado.Caption = chkRiscado.Value
Cor.BackColor = picOutras.BackColor
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Status"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmEditor.Enabled = True
frmEditor.ActiveForm.SetFocus
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Color of text selected"
End Sub

Private Sub lblExemplo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Sample of the text selected"
End Sub

Private Sub lstEstiloDaFonte_Click()
txtEstiloDaFonte.Text = lstEstiloDaFonte.Text
On Error Resume Next
If lstEstiloDaFonte.Text = "Normal" Then
lblExemplo.FontBold = False
lblExemplo.FontItalic = False
ElseIf lstEstiloDaFonte.Text = "Bold" Then
lblExemplo.FontBold = True
lblExemplo.FontItalic = False
ElseIf lstEstiloDaFonte.Text = "Italic" Then
lblExemplo.FontBold = False
lblExemplo.FontItalic = True
Else
lblExemplo.FontBold = True
lblExemplo.FontItalic = True
End If
End Sub

Private Sub lstEstiloDaFonte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Format the text selected is: " & lstEstiloDaFonte.Text
End Sub

Private Sub lstFontes_Click()
On Error Resume Next
txtFontes.Text = lstFontes.Text
lblExemplo.FontName = lstFontes.Text
End Sub



Private Sub lstFontes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "The font of the text selected is: " & lstFontes.Text
End Sub

Private Sub lstTamanhoDaFonte_Click()
txtTamanhoDaFonte.Text = lstTamanhoDaFonte.Text
On Error Resume Next
lblExemplo.FontSize = lstTamanhoDaFonte
End Sub

Private Sub lstTamanhoDaFonte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "The size of the font selected is: " & lstTamanhoDaFonte.Text
End Sub

Private Sub picAmareloClaro_Click()
lblExemplo.ForeColor = picAmareloClaro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 1
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0

End Sub

Private Sub picAmareloEscuro_Click()
lblExemplo.ForeColor = picAmareloEscuro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 1
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0

End Sub

Private Sub picAzulClaro_Click()
lblExemplo.ForeColor = picAzulClaro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 1
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0
End Sub

Private Sub picAzulEscuro_Click()
lblExemplo.ForeColor = picAzulEscuro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 1
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0
End Sub

Private Sub picLaranjaClaro_Click()
lblExemplo.ForeColor = picLaranjaClaro.BackColor
picLaranjaClaro.BorderStyle = 1
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0
End Sub

Private Sub picLaranjaEscuro_Click()
lblExemplo.ForeColor = picLaranjaEscuro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 1
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0
End Sub

Private Sub picOutras_Click()
cmmCor.ShowColor
picOutras.BackColor = cmmCor.Color
lblExemplo.ForeColor = picOutras.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 1
End Sub

Private Sub picPretoClaro_Click()
lblExemplo.ForeColor = picPretoClaro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 1
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0
End Sub

Private Sub picPretoEscuro_Click()
lblExemplo.ForeColor = picPretoEscuro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 1
picOutras.BorderStyle = 0
End Sub

Private Sub picRoxoClaro_Click()
lblExemplo.ForeColor = picRoxoClaro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 1
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0
End Sub

Private Sub picRoxoEscuro_Click()
lblExemplo.ForeColor = picRoxoEscuro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 1
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0
End Sub

Private Sub picVerdeClaro_Click()
lblExemplo.ForeColor = picVerdeClaro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 1
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0
End Sub

Private Sub picVerdeEscuro_Click()
lblExemplo.ForeColor = picVerdeEscuro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 1
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0
End Sub

Private Sub picVermelhoClaro_Click()
lblExemplo.ForeColor = picVermelhoClaro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 1
picVermelhoEscuro.BorderStyle = 0
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0
End Sub

Private Sub picVermelhoEscuro_Click()
lblExemplo.ForeColor = picVermelhoEscuro.BackColor
picLaranjaClaro.BorderStyle = 0
picLaranjaEscuro.BorderStyle = 0
picVermelhoClaro.BorderStyle = 0
picVermelhoEscuro.BorderStyle = 1
picAmareloClaro.BorderStyle = 0
picAmareloEscuro.BorderStyle = 0
picVerdeClaro.BorderStyle = 0
picVerdeEscuro.BorderStyle = 0
picAzulClaro.BorderStyle = 0
picAzulEscuro.BorderStyle = 0
picRoxoClaro.BorderStyle = 0
picRoxoEscuro.BorderStyle = 0
picPretoClaro.BorderStyle = 0
picPretoEscuro.BorderStyle = 0
picOutras.BorderStyle = 0

End Sub

Private Sub txtEstiloDaFonte_Change()
lstEstiloDaFonte.ListIndex = SendMessage(lstEstiloDaFonte.hwnd, LB_FINDSTRING, -1, ByVal CStr(txtEstiloDaFonte.Text))
End Sub

Private Sub txtFontes_Change()
lstFontes.ListIndex = SendMessage(lstFontes.hwnd, LB_FINDSTRING, -1, ByVal CStr(txtFontes.Text))
End Sub

Private Sub txtTamanhoDaFonte_Change()
lstTamanhoDaFonte.ListIndex = SendMessage(lstTamanhoDaFonte.hwnd, LB_FINDSTRING, -1, ByVal CStr(txtTamanhoDaFonte.Text))
End Sub
