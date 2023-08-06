VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Programmer: Jason Hensley
'Email: elitecobra@hotmail.com
'Website: http://www16.brinkster.com/codesource
'Description: Designed to be simple and easy to use.
'You can easily make a complete mp3, wave, ,asf,
'midi player, ect. with this MM Module.
'
'Api to access the registry
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
 Const HKEY_CLASSES_ROOT = &H80000000
'Api to move form without titlebar
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long
'Api to send the commands to the mci device
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'Used to store the Filename
Public FileName As String
'Pause for a certain amount of time
Sub TimeOut(duration)
    StartTime = Timer


    Do While Timer - StartTime < duration
        x = DoEvents()
    Loop
End Sub

'Play a song
Public Sub Play()
On Error GoTo TrapIt
mciSendString "close " & FileName, 0, 0, 0
FileName = Chr$(34) + Trim(FileName) + Chr$(34)
 mciSendString "open " & FileName, 0, 0, 0
 mciSendString "play " & FileName, 0, 0, 0
Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Make your app the default app when they open a file
Public Sub MyAppDefault(ByVal sAppName As String, ByVal sEXE As String, ByVal sExt As String)
'Ex. MyAppDefault "My MP3 Player", App.Path & "\" & App.Exe & " %1", ".mp3"
Dim lRegKey As Long
Call RegCreateKey(HKEY_CLASSES_ROOT, sExt, lRegKey)
Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sAppName, Len(sAppName))
Call RegCloseKey(lRegKey)
Call RegCreateKey(HKEY_CLASSES_ROOT, sAppName & "\Shell\Open\Command", lRegKey) ' adds info into the shell open command
Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sEXE, Len(sEXE))
Call RegCloseKey(lRegKey)
End Sub
'Check if the song has ended. Good for repeat play, ect.
Public Function EndOfSong() As Boolean
On Error GoTo TrapIt
Dim ThePos As Long
Dim TheDur As Long
ThePos = PositioninSec
TheDur = DurationInSec
If ThePos = 0 Or TheDur = 0 Then Exit Function
If ThePos = TheDur Then
EndOfSong = True
Else
EndOfSong = False
End If
Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Move a borderless form
Public Sub FormMove(Frm As Form)
    ReleaseCapture
    Call SendMessage(Frm.hWnd, &HA1, 2, 0&)
End Sub
'Open and load a .m3u playlist
Public Sub OpenPlaylist(TheList As String, Listbox As Control)
On Error GoTo TrapIt
Dim test As String
If TheList = "" Then Exit Sub
Open TheList For Input As 1
While Not EOF(1)
Line Input #1, test
Listbox.AddItem RTrim(test)
Wend
Close 1
Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Get the time remaining in milli-seconds
Public Function TimeRemaininginMS() As Long
On Error GoTo TrapIt
TimeRemaininginMS = DurationInMS - PositioninMS
Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Get the time remaining in seconds
Public Function TimeRemaininginSec() As Long
On Error GoTo TrapIt
TimeRemaininginSec = DurationInSec - PositioninSec
Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Get the formatted time remaining
Public Function FormatTimeRemaining() As String
On Error GoTo TrapIt
   FormatTimeRemaining = GetThisTime(TimeRemaininginMS)
Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Used to find the last backslash of the file location
Private Function GetLastBackSlash(text As String) As String
On Error GoTo TrapIt
    Dim i, pos As Integer
    Dim lastslash As Integer
    For i = 1 To Len(text)
        pos = InStr(i, text, "\", vbTextCompare)
        If pos <> 0 Then lastslash = pos
    Next i
    GetLastBackSlash = Right(text, Len(text) - lastslash)
Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
''Used for getting the Right/Left of a string
Private Function RightLeft(source As String, token As String) As String
On Error GoTo TrapIt
Dim i As Long
RightLeft = ""
For i = Len(source) To 1 Step -1
If Mid(source, i, 1) = token Then
RightLeft = Left(source, i - 1)
Exit Function
End If
Next i
Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Strip the path and ext. from a list
Public Sub ListNoChar(List1 As Listbox, List2 As Control)
On Error GoTo TrapIt
Dim x As Long
Dim NoChar As String
Dim NoEnd As String
For x = 0 To List2.ListCount - 1
NoChar = GetLastBackSlash(List2.List(x))
NoEnd = RightLeft(NoChar, ".")
List1.AddItem NoEnd
Next x
Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Strip the path and ext. from a single file
Public Sub ListSingleNoChar(List1, List2)
On Error GoTo TrapIt
Dim x As String
Dim NoChar As String
Dim NoEnd As String
x = List2.ListCount - 1
List2.ListIndex = x
x = List2.text
NoChar = GetLastBackSlash(x)
NoEnd = RightLeft(NoChar, ".")
List1.AddItem NoEnd
MsgBox x
Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Take the .mp3 off the end of file
Private Function NoEndChar(List1 As Listbox, List2 As Listbox) As String
On Error GoTo TrapIt
Dim n As Long
For n = 0 To List2.ListCount - 1
NoEndChar = Left(List2.List(n), 1)
Next n
Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Save a playlist in .m3u format
Public Sub SavePlaylist(TheList As String, Listbox As Control)
On Error GoTo TrapIt
Dim i As Integer
Dim a As String
Open TheList For Output As #1
For i = 0 To Listbox.ListCount - 1
a$ = Listbox.List(i)
Print #1, a$
Next
Close 1
Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Get the position in MilliSeconds
Public Function PositioninMS() As Long
On Error GoTo TrapIt
Static s As String * 30
mciSendString "set " & FileName & " time format milliseconds", 0, 0, 0
mciSendString "status " & FileName & " position", s, Len(s), 0
PositioninMS = Val(s)
Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Get the Position in Seconds
Public Function PositioninSec() As Long
On Error GoTo TrapIt
PositioninSec = Val(PositioninMS \ 1000)
Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Get the Formatted Position
Public Function FormatPosition() As String
On Error GoTo TrapIt
Dim Sec As Long
Dim mins As Long
Static s As String * 30
    mciSendString "set " & FileName & " time format milliseconds", 0, 0, 0
    mciSendString "status " & FileName & " position", s, Len(s), 0
    Sec = PositioninSec
    If Sec < 60 Then FormatPosition = "0:" & format(Sec, "00")
    If Sec > 59 Then
        mins = Int(Sec / 60)
        Sec = Sec - (mins * 60)
        FormatPosition = format(mins, "0") & ":" & format(Sec, "00")
    End If
Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Change the position of the song
Public Sub ChangePosition(TheSecond As Long)
On Error GoTo TrapIt
TheSecond = TheSecond * 1000
If IsPlaying = True Then mciSendString "play " & FileName & " from " & TheSecond, 0, 0, 0
If IsPlaying = False Then mciSendString "seek " & FileName & " to " & TheSecond, 0, 0, 0
Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Stop playing the song
Public Sub StopPlay()
On Error GoTo TrapIt
mciSendString "close " & FileName, 0, 0, 0
'mciSendString "stop " & FileName, 0, 0, 0
Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Pause the song
Public Sub Pause()
On Error GoTo TrapIt
mciSendString "stop " & FileName, 0, 0, 0
Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Get duration in Milliseconds
Public Function DurationInMS() As Long
On Error GoTo TrapIt
Dim TotalTime As String * 128
    mciSendString "status " & FileName & " length", TotalTime, 128, 0&
    DurationInMS = Val(TotalTime)
Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Get the formatted duration
Public Function FormatDuration() As String
On Error GoTo TrapIt
Dim TotalTime As String * 128
Dim T As String
Dim lTotalTime As Long

    mciSendString "set " & FileName & " time format ms", TotalTime, 128, 0&
    mciSendString "status " & FileName & " length", TotalTime, 128, 0&

    mciSendString "set " & FileName & " time format frames", 0&, 0&, 0&
    
    lTotalTime = Val(TotalTime)
   T = GetThisTime(lTotalTime)
    FormatDuration = T
Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Used to format the position and duration
Private Function GetThisTime(ByVal timein As Long) As String
    On Error GoTo TrapIt
    Dim conH As Integer
    Dim conM As Integer
    Dim conS As Integer
    Dim remTime As Long
    Dim strRetTime As String
    remTime = timein / 1000
    conH = Int(remTime / 3600)
    remTime = remTime Mod 3600
    conM = Int(remTime / 60)
    remTime = remTime Mod 60
    conS = remTime
    
    If conH > 0 Then
        strRetTime = Trim(STR(conH)) & ":"
    Else
        strRetTime = ""
    End If
    
    If conM >= 10 Then
        strRetTime = strRetTime & Trim(STR(conM))
    ElseIf conM > 0 Then
        strRetTime = strRetTime & Trim(STR(conM))
    Else
        strRetTime = strRetTime & "0"
    End If
    
    strRetTime = strRetTime & ":"
    
    If conS >= 10 Then
        strRetTime = strRetTime & Trim(STR(conS))
    ElseIf conS > 0 Then
        strRetTime = strRetTime & "0" & Trim(STR(conS))
    Else
        strRetTime = strRetTime & "00"
    End If
    
    GetThisTime = strRetTime
Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Get and save setting to the registry
Private Sub GetandSaveSettings()
'You can make these function in the module yourself if you want them.
'These function are built into vb(VB 6.0 i've tested with)
'These are good for saving and loading the location of a playlist or something. For a autoload feature, ect.
'This is how you can save/get a setting to/from the registry
'a = SaveSetting("Your Player Here", "What the setting is for", "Location", Form1.C.FileName)
'b = GetSetting("FV8 MP3 Player", "AutoLoad", "True or False")
End Sub


'Get the Duration in Seconds
Public Function DurationInSec() As Long
On Error GoTo TrapIt
DurationInSec = DurationInMS \ 1000
Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Resume a paused song
Public Sub ResumePlay()
On Error GoTo TrapIt
mciSendString "play " & FileName, 0, 0, 0
Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Check to see if a song is playing
Public Function IsPlaying() As Boolean
On Error GoTo TrapIt
Dim a As Long
a = mciSendString("status " & FileName & " mode", 0, 0, 0)
If a = "0" Then
IsPlaying = True
Else
IsPlaying = False
End If
Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function

