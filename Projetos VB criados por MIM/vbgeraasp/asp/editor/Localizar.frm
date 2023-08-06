VERSION 5.00
Begin VB.Form frmLocalizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   2025
   ClientLeft      =   3315
   ClientTop       =   2745
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5925
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1095
      Left            =   360
      TabIndex        =   7
      Top             =   720
      Width           =   3255
      Begin VB.CheckBox chkPalavraInteira 
         Caption         =   "&Whole word only"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chkMaiúsculaMinúscula 
         Caption         =   "&Match Case"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdLocalizarNoTextoSelecionado 
      Caption         =   "Find in text selected"
      Default         =   -1  'True
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdLocalizarPrimeira 
      Caption         =   "&Find First"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtLocalizar 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Find:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmLocalizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkMaiúsculaMinúscula_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Find Case-sensitive"
End Sub

Private Sub chkPalavraInteira_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Find whole word only"
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdLocalizarNoTextoSelecionado_Click()
If chkMaiúsculaMinúscula.Value = 1 And chkPalavraInteira.Value = 1 Then
PosiçãoDaPalavraASerProcurada = frmEditor.ActiveForm.rtbDocumento.Find(txtLocalizar.Text, , , rtfWholeWord + rtfMatchCase)
    If PosiçãoDaPalavraASerProcurada = -1 Then
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word not found"
        MsgBox "Word not found", vbInformation, "Find"
    Else
        Unload Me
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word found"
    End If
ElseIf chkMaiúsculaMinúscula.Value = 1 Then
PosiçãoDaPalavraASerProcurada = frmEditor.ActiveForm.rtbDocumento.Find(txtLocalizar.Text, , , rtfMatchCase)
    If PosiçãoDaPalavraASerProcurada = -1 Then
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word not found"
        MsgBox "Palavra não encontrada", vbInformation, "Find"
    Else
        Unload Me
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word found"
    End If
ElseIf chkPalavraInteira.Value = 1 Then
PosiçãoDaPalavraASerProcurada = frmEditor.ActiveForm.rtbDocumento.Find(txtLocalizar.Text, , , rtfWholeWord)
    If PosiçãoDaPalavraASerProcurada = -1 Then
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word not found"
        MsgBox "Palavra não encontrada", vbInformation, "Find"
    Else
        Unload Me
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word found"
    End If
Else
PosiçãoDaPalavraASerProcurada = frmEditor.ActiveForm.rtbDocumento.Find(txtLocalizar.Text)
    If PosiçãoDaPalavraASerProcurada = -1 Then
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word not found"
        MsgBox "Palavra não encontrada", vbInformation, "Find"
    Else
        Unload Me
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word found"
    End If
End If

End Sub

Private Sub cmdLocalizarPrimeira_Click()
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Finding word..."
On Error Resume Next
Dim PosiçãoDaPalavraASerProcurada
If chkMaiúsculaMinúscula.Value = 1 And chkPalavraInteira.Value = 1 Then
PosiçãoDaPalavraASerProcurada = frmEditor.ActiveForm.rtbDocumento.Find(txtLocalizar.Text, 0, , rtfWholeWord + rtfMatchCase)
    If PosiçãoDaPalavraASerProcurada = -1 Then
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word not found"
        MsgBox "Palavra não encontrada", vbInformation, "Find"
    Else
        Unload Me
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word found"
    End If
ElseIf chkMaiúsculaMinúscula.Value = 1 Then
PosiçãoDaPalavraASerProcurada = frmEditor.ActiveForm.rtbDocumento.Find(txtLocalizar.Text, 0, , rtfMatchCase)
    If PosiçãoDaPalavraASerProcurada = -1 Then
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word not found"
        MsgBox "Palavra não encontrada", vbInformation, "Find"
    Else
        Unload Me
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word found"
    End If
ElseIf chkPalavraInteira.Value = 1 Then
PosiçãoDaPalavraASerProcurada = frmEditor.ActiveForm.rtbDocumento.Find(txtLocalizar.Text, 0, , rtfWholeWord)
    If PosiçãoDaPalavraASerProcurada = -1 Then
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word not found"
        MsgBox "Palavra não encontrada", vbInformation, "Find"
    Else
        Unload Me
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word found"
    End If
Else
PosiçãoDaPalavraASerProcurada = frmEditor.ActiveForm.rtbDocumento.Find(txtLocalizar.Text, 0)
    If PosiçãoDaPalavraASerProcurada = -1 Then
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word not found"
        MsgBox "Palavra não encontrada", vbInformation, "Find"
    Else
        frmEditor.BarraDeStatus.Panels.Item(1).Text = "Word found"
        Unload Me
    End If
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

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Status"
End Sub
