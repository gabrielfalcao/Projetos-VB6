VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form otim 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Otimizador de Performance"
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   Icon            =   "listapassando2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   -1005
      TabIndex        =   1
      Top             =   345
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Max             =   5
      Scrolling       =   1
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1590
      TabIndex        =   3
      Top             =   1650
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1432
      Top             =   1665
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   450
      Left            =   1155
      TabIndex        =   0
      Top             =   1170
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Otimizando computador para carregar o Teg Chat..."
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   30
      Width           =   6030
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   15
      TabIndex        =   2
      Top             =   315
      Width           =   6015
   End
End
Attribute VB_Name = "otim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Const GW_OWNER = 4
Private Const GWL_STYLE = (-16)
Private Const WS_DISABLED = &H8000000
Private Const WM_CANCELMODE = &H1F
Private Const WM_CLOSE = &H10
Function EndTask(TargetHwnd As Long) As Long
  Dim Tmp1 As Long
  
  If TargetHwnd = hwnd Or GetWindow(TargetHwnd, GW_OWNER) = hwnd Then
    End
  End If
  
  If IsWindow(TargetHwnd) = False Then GoTo EndTaskFail
  If (GetWindowLong(TargetHwnd, GWL_STYLE) And WS_DISABLED) Then GoTo EndTaskSucceed
    
  If IsWindow(TargetHwnd) Then
    If Not (GetWindowLong(TargetHwnd, GWL_STYLE) And WS_DISABLED) Then
      PostMessage TargetHwnd, WM_CANCELMODE, 0&, 0&
      PostMessage TargetHwnd, WM_CLOSE, 0&, 0&
      DoEvents
    End If
  End If
  
  GoTo EndTaskSucceed
    
EndTaskFail:
  Tmp1 = False
  GoTo EndTaskEndSub
EndTaskSucceed:
  Tmp1 = True
EndTaskEndSub:
  EndTask = Tmp1
End Function

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()


pb.Max = 5
Dim maximo As String
maximo = List1.ListIndex
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Select Case pb.Value + 1
Case 1
Label1.Caption = "Fechando Possíveis Invasores..."

On Error Resume Next
  Dim Tmp1 As Long
  Dim Tmp2 As Long
  Dim tmp3 As Long
  Dim tmp4 As Long
  Tmp1 = FindWindow(vbNullString, "WUpdate")
  tmp3 = FindWindow(vbNullString, "Windows Update")
    tmp4 = FindWindow(vbNullString, "WindowsUpdate")
  Tmp2 = EndTask(Tmp1)
  
pb.Value = 1

Case 2
Label1.Caption = "Deletando arquivos desnecessários..."
pb.Value = 2
Kill "C:\Windows\System32\Windows Update.exe"
Case Is = 3
Label1.Caption = "Limpando Registro do Sistema..."
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "Windows Update", ""
pb.Value = 3
Case Is = 4
Label1.Caption = "Tentando Acelerar o PC..."
pb.Value = 4

Case Is = 5
Label1.Caption = "Abrindo Teg Chat..."
pb.Value = 5
frmChatCliente.Show
Unload Me

Case Else
End Select

End Sub

