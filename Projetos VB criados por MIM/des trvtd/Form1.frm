VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desinstalador Trava Tudo v1.12"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Desinstalar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   1560
      TabIndex        =   3
      Top             =   705
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Height          =   330
      Left            =   30
      ScaleHeight     =   270
      ScaleWidth      =   4575
      TabIndex        =   2
      Top             =   1875
      Width           =   4635
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000F&
         FillColor       =   &H8000000F&
         FillStyle       =   3  'Vertical Line
         Height          =   270
         Left            =   0
         Top             =   0
         Width           =   30
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clique em ""Desinstalar"" para remover o Trava Tudo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   30
      TabIndex        =   1
      Top             =   2220
      Width           =   4635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desinstalador Trava Tudo v1.12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RunKey As String


Private Sub Command1_Click()
If Shape1.Width < 4575 Then
For i = Shape1.Width To Shape1.Width + 1
Label2.Caption = "Removendo arquivos..."
Shape1.Width = 2255
Dim file As String
file = App.Path & "\" & "Trava Tudo.exe"
If fileExist(file) = True Then
  Kill file
End If
Next i
For i = Shape1.Width To Shape1.Width + 1
Dim file2 As String
file2 = App.Path & "\" & "Vbrun60.exe"
If fileExist(file2) = True Then
  Kill file2
End If
Next i
Shape1.Width = 3595
Label2.Caption = "Limpando o registro..."
SetStringValue RunKey, "Trava Tudo", ""
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\trvtd", "CHKval", ""
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\trvtd", "PTAESSST", ""
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Trava Tudo", ""
Shape1.Width = 4595
For i = Shape1.Width To Shape1.Width + 1
Label2.Caption = "Desinstalação Concluída"
Command1.Enabled = False
Next i
End If
End Sub

Private Sub Form_Load()
RunKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
End Sub
