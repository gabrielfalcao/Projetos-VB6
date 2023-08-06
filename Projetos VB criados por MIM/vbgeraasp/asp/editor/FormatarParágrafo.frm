VERSION 5.00
Begin VB.Form frmFormatarParágrafo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Format Paragraph"
   ClientHeight    =   3810
   ClientLeft      =   3750
   ClientTop       =   2310
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3630
   Begin VB.Frame Frame2 
      Caption         =   "Align of paragraph"
      Height          =   1335
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   3135
      Begin VB.OptionButton optDireita 
         Caption         =   "Align to right"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   960
         Width           =   1800
      End
      Begin VB.OptionButton optCentralizado 
         Caption         =   "Center"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   630
         Width           =   1575
      End
      Begin VB.OptionButton optEsquerda 
         Caption         =   "Align to left"
         Height          =   255
         Left            =   135
         TabIndex        =   11
         Top             =   285
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Margin of paragraph"
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3135
      Begin VB.TextBox txtMargemEsquerda 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   780
         Width           =   990
      End
      Begin VB.TextBox txtMargemDireita 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lblMargemDireita 
         Caption         =   "Margin right:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblMargemEsquerda 
         Caption         =   "Margin left:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "cm"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   780
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "cm"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Alinhamento 
      Caption         =   "Label2"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label MargemEsquerda 
      Caption         =   "Label2"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label MargemDireita 
      Caption         =   "Label2"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   3960
      Width           =   1095
   End
End
Attribute VB_Name = "frmFormatarParágrafo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAplicar_Click()
'1 twip = 567cm
On Error Resume Next
frmEditor.ActiveForm.rtbDocumento.RightMargin = txtLargura.Text * 567
frmEditor.ActiveForm.rtbDocumento.SelRightIndent = txtMargemDireita.Text * 567
frmEditor.ActiveForm.rtbDocumento.SelIndent = txtMargemEsquerda.Text * 567
If optEsquerda.Value = True Then
frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfLeft
ElseIf optCentralizado.Value = True Then
frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfCenter
ElseIf optDireita.Value = True Then
frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfRight
End If
End Sub

Private Sub cmdCancelar_Click()
txtMargemDireita.Text = MargemDireita.Caption
txtMargemEsquerda.Text = MargemEsquerda.Caption
If Alinhamento.Caption = "Left" Then
optEsquerda.Value = True
ElseIf Alinhamento.Caption = "Right" Then
optDireita.Value = True
Else
optCentralizado.Value = True
End If
cmdAplicar_Click
Unload Me
End Sub

Private Sub cmdOK_Click()
'1 twip = 567cm
On Error Resume Next
frmEditor.ActiveForm.rtbDocumento.SelRightIndent = txtMargemDireita.Text * 567
frmEditor.ActiveForm.rtbDocumento.SelIndent = txtMargemEsquerda.Text * 567
If optEsquerda.Value = True Then
frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfLeft
ElseIf optCentralizado.Value = True Then
frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfCenter
ElseIf optDireita.Value = True Then
frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfRight
End If
Unload Me
End Sub


Private Sub Form_Load()
'1 twip = 567cm
On Error Resume Next
txtMargemDireita.Text = frmEditor.ActiveForm.rtbDocumento.SelRightIndent / 567
txtMargemEsquerda.Text = frmEditor.ActiveForm.rtbDocumento.SelIndent / 567

If frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfLeft Then
optEsquerda.Value = True
ElseIf frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfCenter Then
optCentralizado.Value = True
ElseIf frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfRight Then
optDireita.Value = True
End If

MargemDireita.Caption = txtMargemDireita.Text
MargemEsquerda.Caption = txtMargemEsquerda.Text
If optEsquerda.Value = True Then
Alinhamento.Caption = "Left"
ElseIf optDireita.Value = True Then
Alinhamento.Caption = "Right"
Else
Alinhamento.Caption = "Center"
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Status"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmEditor.Enabled = True
frmEditor.ActiveForm.SetFocus
End Sub

Private Sub lblMargemDireita_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Right margin of the paragraph where is cursor"
End Sub

Private Sub lblMargemEsquerda_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Left margin of the paragraph where is cursor"
End Sub

Private Sub optCentralizado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Center paragraph where is cursor"
End Sub

Private Sub optDireita_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Align paragraph where is cursor at right"
End Sub

Private Sub optEsquerda_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Align paragraph where is cursor at left (default)"
End Sub

Private Sub txtMargemDireita_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Right margin of the paragraph where is cursor"
End Sub

Private Sub txtMargemEsquerda_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Left margin of the paragraph where is cursor"
End Sub
