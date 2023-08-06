VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00D89970&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slideshow 1.0.0.1"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10125
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00AFAFF3&
      Caption         =   "Parar Slideshow"
      Height          =   375
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   1725
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2160
      Top             =   6240
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00A4E3AC&
      Caption         =   "Iniciar Slideshow"
      Height          =   375
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1725
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ESTICADA"
      Height          =   300
      Left            =   1290
      TabIndex        =   4
      Top             =   5790
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NORMAL"
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   5790
      Width           =   1155
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFAEA&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Top             =   975
      Width           =   2370
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFAEA&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   2340
      Left            =   90
      TabIndex        =   1
      Top             =   1335
      Width           =   2370
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFAEA&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   2040
      Left            =   90
      Pattern         =   "*.jpg;*.gif.*.bmp;*.jpeg;*.tif"
      TabIndex        =   0
      Top             =   3735
      Width           =   2400
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":08CA
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   420
      TabIndex        =   7
      Top             =   75
      Width           =   8460
   End
   Begin VB.Image Image1 
      Height          =   6195
      Left            =   2640
      Top             =   990
      Width           =   7080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i
Private Sub Command1_Click()
Image1.Stretch = False
Image1.Refresh
End Sub

Private Sub Command2_Click()
Image1.Height = 6120
Image1.Width = 6840
Image1.Stretch = True
Image1.Refresh
End Sub

Private Sub Command3_Click()
On Error Resume Next
Timer1.Enabled = True
File1.ListIndex = File1.ListIndex + 1
End Sub

Private Sub Command4_Click()
Timer1.Enabled = False
End Sub

Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1.Path
File1.ListIndex = 0

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Dim nome As String
If Len(File1.Path) = 3 Then
nome = File1.Path & File1.FileName
Else
nome = File1.Path & "\" & File1.FileName
End If
Image1.Picture = LoadPicture(nome)

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If File1.ListIndex < File1.ListCount - 1 Then
File1.ListIndex = File1.ListIndex + 1
End If
End Sub
