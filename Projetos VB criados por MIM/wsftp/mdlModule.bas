Attribute VB_Name = "mdlModule"
'WinFTP, created by the KPD-Team 2000
'This file can be downloaded from http://www.allapi.net/
'For questions or comments, contact us at KPDTeam@Allapi.net

' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.

Public Const FOP_UPLOAD = &H1
Public Const FOP_DOWNLOAD = &H2
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Type tFO
    sName As String
    sPath As String
    bProcedure As Byte
    bCompleted As Boolean
    nFileSize As Long
End Type
Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long
Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetTickCount& Lib "kernel32" ()
Public bFOBusy As Boolean
Public foFiles() As tFO, foItems As Long, TotalFileSize As Long, SentBytes As Long, OldSpeed As Single
Public ActiveFileBytesSent As Long, ActiveFileBytesTotal As Long, UploadFlag As Long
Public ActiveFile As String, ActiveIndex As Long, ActiveProcedure As Byte, StartT As Long
Public Sub StartFO()
    Dim Ret As Long, FF As Integer
    bFOBusy = True
    If foItems <> 0 Then
        Ret = GetNextFile
        StartT = GetTickCount
        While Ret <> -1
            OldSpeed = 0
            ActiveIndex = Ret
            ActiveFile = foFiles(Ret).sName
            ActiveProcedure = foFiles(Ret).bProcedure
            ActiveFileBytesTotal = foFiles(Ret).nFileSize
            ActiveFileBytesSent = 0
            If foFiles(Ret).bProcedure = FOP_UPLOAD Then
                frmMain.rfFile.RemoteFile = foFiles(Ret).sName
                frmMain.rfFile.UploadFile frmMain.rfConnection, foFiles(Ret).sPath + foFiles(Ret).sName
                foFiles(Ret).bCompleted = True
            ElseIf foFiles(Ret).bProcedure = FOP_DOWNLOAD Then
                frmMain.rfFile.RemoteFile = foFiles(Ret).sName
                frmMain.rfFile.GetFile frmMain.rfConnection
                FF = FreeFile
                Open foFiles(Ret).sPath + foFiles(Ret).sName For Binary As #FF
                    Put #FF, , frmMain.rfFile.FileData
                Close #FF
                foFiles(Ret).bCompleted = True
            End If
            GetStatus
            Ret = GetNextFile
        Wend
        foItems = 0
        ReDim foFiles(1 To 1) As tFO
        TotalFileSize = 0
        SentBytes = 0
        frmMain.FillRemoteListView
        frmMain.FillLocalListView frmMain.sCurPath
        ActiveFile = ""
        ActiveFileBytesSent = 0
        ActiveFileBytesTotal = 1
        ActiveProcedure = 0
        TotalFileSize = 1
        SentBytes = 0
        frmProgress.UpdateProgress
        NotifyWhenComplete
    End If
    bFOBusy = False
End Sub
Public Function GetNextFile() As Long
    Dim Cnt As Long
    GetNextFile = -1
    For Cnt = 1 To foItems
        If foFiles(Cnt).bCompleted = False Then
            GetNextFile = Cnt
            Exit For
        End If
    Next Cnt
End Function
Public Function NotifyWhenComplete()
    frmMain.WindowState = vbNormal
    SetForegroundWindow frmMain.hWnd
    Beep
End Function
Public Sub AddToCollection(bProcedure As Byte, sFile As String, sPath As String, nFileSize As Long)
    Dim Cnt As Long, bOk As Boolean
    For Cnt = 1 To foItems
        If foFiles(Cnt).sName = sFile Then
            bOk = True
            Exit For
        End If
    Next Cnt
    If bOk = False Then
        foItems = foItems + 1
        ReDim Preserve foFiles(1 To foItems) As tFO
        foFiles(foItems).bProcedure = bProcedure
        foFiles(foItems).nFileSize = nFileSize
        foFiles(foItems).sName = sFile
        foFiles(foItems).sPath = sPath
        TotalFileSize = TotalFileSize + nFileSize
    End If
End Sub
Sub GetStatus()
    frmMain.txtStatus.Text = frmMain.txtStatus.Text + frmMain.rfConnection.GetLastResponseInfo
    frmMain.txtStatus.SelStart = Len(frmMain.txtStatus.Text)
End Sub
