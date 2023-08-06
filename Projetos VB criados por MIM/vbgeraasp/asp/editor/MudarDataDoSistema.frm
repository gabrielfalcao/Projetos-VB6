VERSION 5.00
Begin VB.Form frmMudarDataDoSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Sistem Date"
   ClientHeight    =   1200
   ClientLeft      =   3540
   ClientTop       =   2970
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3345
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtNovaData 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "New date:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblDataAtual 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmMudarDataDoSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Erro
Date = txtNovaData.Text
Unload Me
Exit Sub
Erro:
MsgBox "Invalid Date. Enter a valid date or click Cancel", vbOKCancel, "Invalid date"
End Sub

Private Sub Form_Load()
lblDataAtual.Caption = Date
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "Status"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmEditor.Enabled = True
frmEditor.ActiveForm.SetFocus
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "System Date"
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "New date to operating system"
End Sub

Private Sub lblDataAtual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "System Date"
End Sub

Private Sub txtNovaData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmEditor.BarraDeStatus.Panels.Item(1).Text = "New date to operating system"
End Sub
