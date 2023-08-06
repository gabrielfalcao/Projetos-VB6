VERSION 5.00
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
      Height          =   1620
      Left            =   210
      TabIndex        =   2
      Top             =   1530
      Width           =   4245
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   945
      Width           =   2490
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Left            =   825
      TabIndex        =   0
      Top             =   240
      Width           =   2205
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_PATH& = 260

Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szexeFile As String * MAX_PATH
End Type

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Dim X(100), Y(100), Z(100) As Integer
Dim tmpX(100), tmpY(100), tmpZ(100) As Integer
Dim K As Integer
Dim Zoom As Integer
Dim Speed As Integer

Private Sub Command1_Click()
KillApp (Text1.Text)
End Sub
Public Function KillApp(myName As String) As Boolean

Const PROCESS_ALL_ACCESS = 0
Dim uProcess As PROCESSENTRY32
Dim rProcessFound As Long
Dim hSnapshot As Long
Dim szExename As String
Dim exitCode As Long
Dim myProcess As Long
Dim AppKill As Boolean
Dim appCount As Integer
Dim i As Integer
On Local Error GoTo Finish
appCount = 0

Const TH32CS_SNAPPROCESS As Long = 2&

uProcess.dwSize = Len(uProcess)
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
rProcessFound = ProcessFirst(hSnapshot, uProcess)
List1.Clear

Do While rProcessFound
i = InStr(1, uProcess.szexeFile, Chr(0))
szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
List1.AddItem (szExename)
If Right$(szExename, Len(myName)) = LCase$(myName) Then
KillApp = True
appCount = appCount + 1
myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
AppKill = TerminateProcess(myProcess, exitCode)
Call CloseHandle(myProcess)
End If


rProcessFound = ProcessNext(hSnapshot, uProcess)
Loop

Call CloseHandle(hSnapshot)
Finish:
End Function

Private Sub Form_Load()
KillApp ("none")
Command1.Caption = "Close Program"
Text1.Text = ""
End Sub

Private Sub List1_Click()
Text1.Text = List1.List(List1.ListIndex)
End Sub

Private Sub Text1_Change()
search$ = UCase$(Text1.Text)
searchlen = Len(search$)
If searchlen Then
For i = 0 To List1.ListCount - 1
If UCase$(Left$(List1.List(i), searchlen)) = search$ Then
List1.ListIndex = i
Text1.SelStart = Len(Text1.Text)
Exit For
End If
Next
End If

End Sub
