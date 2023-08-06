VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UUFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Event UUFileFound(ByVal UUFileName As String, ByVal UUFilePath As String, ByVal UUFileSize As Long)
Public Event UUFindComplete(ByVal UUTotalFolders As String, ByVal UUTotalFiles As String, ByVal UUTotalSize As Long)

'Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Dim mDirCol As UUFindList

'Private Type FILETIME
'
'    dwLowDateTime As Long
'    dwHighDateTime As Long
'
'End Type
'
'Private Type WIN32_FIND_DATA
'
'    dwFileAttributes As Long
'    ftCreationTime As FILETIME
'    ftLastAccessTime As FILETIME
'    ftLastWriteTime As FILETIME
'    nFileSizeHigh As Long
'    nFileSizeLow As Long
'    dwReserved0 As Long
'    dwReserved1 As Long
'    cFileName As String * MAX_PATH
'    cAlternate As String * 14
'
'End Type

Private Function StripNulls(OriginalStr As String) As String
    
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
    
End Function
Public Function FindFilesAPI(Path As String, SearchStr As String) As UUFindList
    Dim mPath As String
    Dim mSearchStr As String
    Dim mFileCount As Integer
    Dim mDirCount As Integer
    Dim mTotalSize As Long
    
    Set mDirCol = New UUFindList
    
    mPath = Path
    mSearchStr = SearchStr
    mTotalSize = FindFilesAPIX(mPath, mSearchStr, mFileCount, mDirCount)
    Set FindFilesAPI = mDirCol
    RaiseEvent UUFindComplete(mDirCount + 1, mFileCount, mTotalSize)
End Function

Private Function FindFilesAPIX(Path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
    Dim fName As String
    Dim dName As String
    Dim dNames() As String
    Dim nDir As Integer
    Dim i As Integer
    Dim hSearch As Long
    Dim WFD As WIN32_FIND_DATA
    
    
    Dim Cont As Integer
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    nDir = 0
    ReDim dNames(nDir)
    Cont = True
    hSearch = FindFirstFile(Path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        dName = StripNulls(WFD.cFileName)
        If (dName <> ".") And (dName <> "..") Then
            If GetFileAttributes(Path & dName) And FILE_ATTRIBUTE_DIRECTORY Then
                dNames(nDir) = dName
                DirCount = DirCount + 1
                nDir = nDir + 1
                ReDim Preserve dNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD)
        Loop
        Cont = FindClose(hSearch)
    End If
    hSearch = FindFirstFile(Path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            fName = StripNulls(WFD.cFileName)
            If (fName <> ".") And (fName <> "..") Then
                FindFilesAPIX = FindFilesAPIX + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                FileCount = FileCount + 1
                RaiseEvent UUFileFound(fName, Path, (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow)
                mDirCol.Add WFD.dwFileAttributes, Win32ToVbTime(WFD.ftCreationTime), Win32ToVbTime(WFD.ftLastAccessTime), Win32ToVbTime(WFD.ftLastWriteTime), WFD.nFileSizeLow, fName, Path
                DoEvents
            End If
            Cont = FindNextFile(hSearch, WFD)
        Wend
        Cont = FindClose(hSearch)
    End If
    If nDir > 0 Then
        For i = 0 To nDir - 1
            FindFilesAPIX = FindFilesAPIX + FindFilesAPIX(Path & dNames(i) & "\", SearchStr, FileCount, DirCount)
        Next i
    End If
End Function


