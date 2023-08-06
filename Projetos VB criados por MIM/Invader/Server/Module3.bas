Attribute VB_Name = "Module3"
Const MAX_PATH As Long = 260
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLenght As Long, ByVal lpBuffer As String) As Long
Public Function getWinDir() As String
  Dim strFolder As String
  Dim lngResult As Long
  strFolder = String(MAX_PATH, 0)
  lngResult = GetWindowsDirectory(strFolder, MAX_PATH)
  If lngResult <> 0 Then
    If Right(Left(strFolder, lngResult), 1) = "\" Then
      getWinDir = Left(strFolder, lngResult)
    Else
      getWinDir = Left(strFolder, lngResult) & "\"
    End If
  Else
    GetWinPath = ""
  End If
End Function
Public Function GetSysDir() As String
  Dim strFolder As String
  Dim lngResult As Long
  strFolder = String(MAX_PATH, 0)
  lngResult = GetSystemDirectory(strFolder, MAX_PATH)
  If lngResult <> 0 Then
    If Right(Left(strFolder, lngResult), 1) = "\" Then
      GetSysDir = Left(strFolder, lngResult)
    Else
      GetSysDir = Left(strFolder, lngResult) & "\"
    End If
  Else
    GetSysDir = ""
  End If
End Function
Public Function GetTempDir() As String
  Dim strFolder As String
  Dim lngResult As Long
  strFolder = String(MAX_PATH, 0)
  lngResult = GetTempPath(MAX_PATH, strFolder)
  If lngResult <> 0 Then
    If Right(Left(strFolder, lngResult), 1) = "\" Then
      GetTempDir = Left(strFolder, lngResult)
    Else
      GetTempDir = Left(strFolder, lngResult) & "\"
    End If
  Else
    GetTempDir = ""
  End If
End Function

