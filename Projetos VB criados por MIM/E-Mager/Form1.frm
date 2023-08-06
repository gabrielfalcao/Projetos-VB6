VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Mager"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdsave 
      Left            =   4215
      Top             =   2250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Salvar image como:"
      Filter          =   "IMAGENS|*.jpg"
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   7485
      Picture         =   "Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salvar imagem"
      Top             =   4035
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.HScrollBar VScroll1 
      Height          =   300
      Left            =   2835
      Max             =   4620
      TabIndex        =   5
      Top             =   4050
      Value           =   4620
      Width           =   4650
   End
   Begin VB.VScrollBar hScroll1 
      Height          =   3870
      Left            =   7485
      Max             =   3855
      TabIndex        =   4
      Top             =   165
      Value           =   3855
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecione o Arquivo:"
      Height          =   4260
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   2760
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   285
         Width           =   2385
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Width           =   2385
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2040
         Left            =   180
         Pattern         =   "*.bmp;*.jpg;*.jpeg;*.dat"
         TabIndex        =   1
         Top             =   2055
         Width           =   2385
      End
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   2850
      Stretch         =   -1  'True
      Top             =   180
      Width           =   4620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Not Image1.Picture = Empty Then
cdsave.ShowSave
SavePicture Image1.Picture, cdsave.FileName
Else
MsgBox "Escolha uma imagem para salvar!", vbCritical, Me.Caption & " - ERRO"
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()
On Error Resume Next
If Len(File1.Path) = 3 Then
Image1.Picture = LoadPicture(File1.Path & File1.FileName)
Else
Image1.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
End If
End Sub

Private Sub HScroll1_Change()
On Error Resume Next
Image1.Height = HScroll1.Value
End Sub

Private Sub VScroll1_Change()
On Error Resume Next
Image1.Width = VScroll1.Value
End Sub
