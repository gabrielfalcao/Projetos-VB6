VERSION 5.00
Begin VB.Form frmInserirFigura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Picture"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInserir 
      Caption         =   "Insert"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   5280
      Width           =   2295
   End
   Begin VB.FileListBox filArquivos 
      Height          =   3210
      Left            =   3000
      Pattern         =   "*.bmp;*.gif;*.jpg"
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.DirListBox dirPastas 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox drvDrives 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image imgExemplo 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   120
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   2655
   End
End
Attribute VB_Name = "frmInserirFigura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInserir_Click()
Clipboard.Clear
Clipboard.SetData imgExemplo.Picture, vbCFDIB
frmEditor.Enabled = True
frmEditor.ActiveForm.rtbDocumento.SetFocus
SendKeys "^v"
Unload Me
End Sub

Private Sub dirPastas_Change()
filArquivos.Path = dirPastas.Path
End Sub

Private Sub drvDrives_Change()
dirPastas.Path = drvDrives.Drive
End Sub

Private Sub filArquivos_Click()
If Len(filArquivos.Path) > 3 Then
    imgExemplo.Picture = LoadPicture(filArquivos.Path & "\" & filArquivos.FileName)
Else
    imgExemplo.Picture = LoadPicture(filArquivos.Path & filArquivos.FileName)
End If
End Sub

Private Sub Form_Load()
dirPastas.Path = drvDrives.Drive
filArquivos.Path = dirPastas.Path
End Sub
