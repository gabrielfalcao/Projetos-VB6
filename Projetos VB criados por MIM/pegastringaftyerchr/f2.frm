VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "f2.frx":0000
      Left            =   45
      List            =   "f2.frx":0002
      TabIndex        =   1
      Top             =   1215
      Width           =   2580
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2355
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   570
      Left            =   3000
      TabIndex        =   0
      Top             =   1410
      Width           =   1290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function InStrRevVB5(ByVal StringCheck As String, ByVal StringMatch As String, Optional ByVal Start As Long = -1) As Long
    Dim lPos        As Long
    Dim lSavePos    As Long
    If Start = -1 Then Start = Len(StringCheck)
    lPos = InStr(1, StringCheck, StringMatch, vbBinaryCompare)
    While lPos > 0 And lPos < Start
        lSavePos = lPos
        lPos = InStr(lPos + 1, StringCheck, StringMatch, vbBinaryCompare)
    Wend
    InStrRevVB5 = lSavePos
End Function

Private Sub Command1_Click()
cd.DialogTitle = "ABRIR"
cd.Filter = "Arquivos de Vídeo|*.mpg;*.mpeg;*.avi"
cd.ShowOpen
List1.AddItem Mid$(cd.FileName, InStrRevVB5(cd.FileName, "\") + 1)
MsgBox "Adicionado " & Mid$(cd.FileName, InStrRevVB5(cd.FileName, "\") + 1)
End Sub

