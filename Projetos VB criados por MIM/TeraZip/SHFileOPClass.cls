VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHFileOPClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_ALLOWUNDO = False
Private Const DEFAULT_CONFIRMMAKEDIR = True
Private Const DEFAULT_CONFIRMOPERATION = True
Private Const DEFAULT_CUSTOMTEXT = ""
Private Const DEFAULT_INCLUDEDIRECTORIES = True
Private Const DEFAULT_PARENTWND = 0
Private Const DEFAULT_RENAMEONCOLLISION = False
Private Const DEFAULT_SILENTMODE = False

Private Const FO_MOVE = 1
Private Const FO_COPY = 2
Private Const FO_DELETE = 3
Private Const FO_RENAME = 4

Private Const FOF_MULTIDESTFILES = &H1
Private Const FOF_SILENT = &H4
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_WANTMAPPINGHANDLE = &H20
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_FILESONLY = &H80
Private Const FOF_SIMPLEPROGRESS = &H100
Private Const FOF_NOCONFIRMMKDIR = &H200

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type


Private m_bAllowUndo As Boolean
Private m_bConfirmMakeDir As Boolean
Private m_bConfirmOperation As Boolean
Private m_sCustomText As String
Private m_bIncludeDirectories As Boolean
Private m_hParentWnd As Long
Private m_bRenameOnCollision As Boolean
Private m_bSilentMode As Boolean
Private m_SourceFiles As New Collection
Private m_DestFiles As New Collection

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, pFrom As Any, ByVal lCount As Long)

Public Property Get AllowUndo() As Boolean
    AllowUndo = m_bAllowUndo
End Property
Public Property Let AllowUndo(bAllowUndo As Boolean)
    m_bAllowUndo = bAllowUndo
End Property

Public Property Get ConfirmMakeDir() As Boolean
    ConfirmMakeDir = m_bConfirmMakeDir
End Property
Public Property Let ConfirmMakeDir(bConfirmMakeDir As Boolean)
    m_bConfirmMakeDir = bConfirmMakeDir
End Property

Public Property Get ConfirmOperation() As Boolean
    ConfirmOperation = m_bConfirmOperation
End Property
Public Property Let ConfirmOperation(bConfirmOperation As Boolean)
    m_bConfirmOperation = bConfirmOperation
End Property

Public Property Get CustomText() As String
    CustomText = m_sCustomText
End Property
Public Property Let CustomText(sCustomText As String)
    m_sCustomText = sCustomText
End Property

Public Property Get IncludeDirectories() As Boolean
    IncludeDirectories = m_bIncludeDirectories
End Property
Public Property Let IncludeDirectories(bIncludeDirectories As Boolean)
    m_bIncludeDirectories = bIncludeDirectories
End Property

Public Property Get ParentWnd() As Long
    ParentWnd = m_hParentWnd
End Property
Public Property Let ParentWnd(hParentWnd As Long)
    m_hParentWnd = hParentWnd
End Property

Public Property Get RenameOnCollision() As Boolean
    RenameOnCollision = m_bRenameOnCollision
End Property
Public Property Let RenameOnCollision(bRenameOnCollision As Boolean)
    m_bRenameOnCollision = bRenameOnCollision
End Property

Public Property Get SilentMode() As Boolean
    SilentMode = m_bSilentMode
End Property
Public Property Let SilentMode(bSilentMode As Boolean)
    m_bSilentMode = bSilentMode
End Property

Public Function MoveFiles() As Boolean
    MoveFiles = DoOperation(FO_MOVE)
End Function

Public Function CopyFiles() As Boolean
    CopyFiles = DoOperation(FO_COPY)
End Function

Public Function DeleteFiles() As Boolean
    DeleteFiles = DoOperation(FO_DELETE)
End Function

Public Function RenameFiles() As Boolean
    RenameFiles = DoOperation(FO_RENAME)
End Function

Public Sub ClearSourceFiles()
    Set m_SourceFiles = Nothing
End Sub

Public Sub ClearDestFiles()
    Set m_DestFiles = Nothing
End Sub

Public Sub AddSourceFile(sFilename As String)
    m_SourceFiles.Add sFilename
End Sub

Public Sub AddDestFile(sFilename As String)
    m_DestFiles.Add sFilename
End Sub

Private Sub Class_Initialize()
    m_bAllowUndo = DEFAULT_ALLOWUNDO
    m_bConfirmMakeDir = DEFAULT_CONFIRMMAKEDIR
    m_bConfirmOperation = DEFAULT_CONFIRMOPERATION
    m_sCustomText = DEFAULT_CUSTOMTEXT
    m_bIncludeDirectories = DEFAULT_INCLUDEDIRECTORIES
    m_hParentWnd = DEFAULT_PARENTWND
    m_bRenameOnCollision = DEFAULT_RENAMEONCOLLISION
    m_bSilentMode = DEFAULT_SILENTMODE
End Sub

Private Function DoOperation(wFunc As Integer) As Boolean
    Dim i As Long, ptr As Long
    Dim shfo As SHFILEOPSTRUCT
    Dim ByteArray() As Byte
    Dim buff1() As Byte, buff2() As Byte, buff3() As Byte

    With shfo
      .hwnd = m_hParentWnd
      .wFunc = wFunc
      .fFlags = 0
      If m_bAllowUndo Then .fFlags = .fFlags Or FOF_ALLOWUNDO
      If m_bSilentMode Then .fFlags = .fFlags Or FOF_SILENT
      If m_bRenameOnCollision Then .fFlags = .fFlags Or FOF_RENAMEONCOLLISION
      If Not m_bConfirmOperation Then .fFlags = .fFlags Or FOF_NOCONFIRMATION
      If Not m_bConfirmMakeDir Then .fFlags = .fFlags Or FOF_NOCONFIRMMKDIR
      If Not m_bIncludeDirectories Then .fFlags = .fFlags Or FOF_FILESONLY
      
      If Len(m_sCustomText) > 0 Then
          .lpszProgressTitle = m_sCustomText
          .fFlags = shfo.fFlags Or FOF_SIMPLEPROGRESS
      End If
   End With

    If m_SourceFiles.Count = 0 Then
        Err.Raise vbObjectError + 1000, , "No source files specified file operation"
    End If
    
    For i = 1 To m_SourceFiles.Count
        shfo.pFrom = shfo.pFrom & m_SourceFiles(i) & Chr$(0)
    Next i

    For i = 1 To m_DestFiles.Count
        shfo.pTo = shfo.pTo & m_DestFiles(i) & Chr$(0)
    Next i

    If m_DestFiles.Count > 1 Then
        shfo.fFlags = shfo.fFlags Or FOF_MULTIDESTFILES
    End If
    
    ReDim ByteArray(LenB(shfo) - 2)
    CopyMemory ByteArray(0), shfo.hwnd, Len(shfo.hwnd)
    CopyMemory ByteArray(4), shfo.wFunc, Len(shfo.wFunc)
    buff1 = StrConv(shfo.pFrom & Chr$(0), vbFromUnicode)
    ptr = VarPtr(buff1(0))
    CopyMemory ByteArray(8), ptr, LenB(ptr)
    buff2 = StrConv(shfo.pTo & Chr$(0), vbFromUnicode)
    ptr = VarPtr(buff2(0))
    CopyMemory ByteArray(12), ptr, LenB(ptr)
    CopyMemory ByteArray(16), shfo.fFlags, Len(shfo.fFlags)
    CopyMemory ByteArray(18), shfo.fAnyOperationsAborted, Len(shfo.fAnyOperationsAborted)
    CopyMemory ByteArray(22), shfo.hNameMappings, Len(shfo.hNameMappings)
    buff3 = StrConv(shfo.lpszProgressTitle & Chr$(0), vbFromUnicode)
    ptr = VarPtr(buff3(0))
    CopyMemory ByteArray(26), ptr, LenB(ptr)
    i = SHFileOperation(ByteArray(0))
    CopyMemory shfo.fAnyOperationsAborted, ByteArray(18), Len(shfo.fAnyOperationsAborted)
    DoOperation = Not CBool(i Or shfo.fAnyOperationsAborted)
End Function

