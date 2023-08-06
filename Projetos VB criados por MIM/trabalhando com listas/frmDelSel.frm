VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   6420
      TabIndex        =   4
      Top             =   690
      Width           =   3315
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   6420
      TabIndex        =   3
      Top             =   1065
      Width           =   3315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deletar Selecionados"
      Height          =   285
      Left            =   7230
      TabIndex        =   2
      Top             =   210
      Width           =   1770
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2340
      TabIndex        =   1
      Top             =   2295
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFD1AD&
      Height          =   3855
      ItemData        =   "frmDelSel.frx":0000
      Left            =   105
      List            =   "frmDelSel.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   75
      Width           =   6285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arq As String

Private Sub Command1_Click()
'List1.ListIndex = 0
'For i = 0 To List1.ListCount - 1
'If List1.Sorted = True Then Kill List1.Text
'List1.ListIndex = i
'Next i
'List1.ListIndex = 0
Kill List1.Tag
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
List1.Clear
arquivos

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Activate()
arquivos
End Sub

Private Sub Form_Load()
If Len(File1.Path) = 3 Then
arq = File1.Path & File1.FileName
Else
arq = File1.Path & "\" & File1.FileName
End If


End Sub
Private Sub charge()

If Len(File1.Path) = 3 Then
arq = File1.Path & File1.FileName
Else
arq = File1.Path & "\" & File1.FileName
End If


End Sub
Private Sub arquivos()

For i = 0 To File1.ListCount - 1
List1.AddItem arq
File1.ListIndex = i
charge
Next i
End Sub
