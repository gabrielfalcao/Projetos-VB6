VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3v1Tag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal dwSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal dwOverlapped As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, dwOverlapped As Long) As Long
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_BEGIN As Long = 0
Private Const FILE_CURRENT As Long = 1
Private Const FILE_END As Long = 2
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const OPEN_EXISTING As Long = 3
Private Const TAGSIZE As Long = &H80
Private Type ID3v1Tag
    sID As String
    sTitle As String
    sArtist As String
    sAlbum As String
    sYear As String
    sComment As String
    bTrack As Byte
    bNull As Byte
    bGenre As Byte
End Type
Private IDTag As ID3v1Tag
Private sFile As String
Private Sub Class_Initialize()
    sFile = ""
    Init
End Sub
Public Property Get Filename() As String
    Filename = sFile
End Property
Public Property Let Filename(ByVal pFilename As String)
    sFile = pFilename
End Property
Public Property Get id() As String
    id = IDTag.sID
End Property
Public Property Get title() As String
    title = VBString(IDTag.sTitle)
End Property
Public Property Let title(ByVal pTitle As String)
    IDTag.sTitle = pTitle
End Property
Public Property Get artist() As String
    artist = VBString(IDTag.sArtist)
End Property
Public Property Let artist(ByVal pArtist As String)
    IDTag.sArtist = pArtist
End Property
Public Property Get album() As String
    album = VBString(IDTag.sAlbum)
End Property
Public Property Let album(ByVal pAlbum As String)
    IDTag.sAlbum = pAlbum
End Property
Public Property Get year() As String
    year = VBString(IDTag.sYear)
End Property
Public Property Let year(ByVal pYear As String)
    IDTag.sYear = pYear
End Property
Public Property Get Comment() As String
    Comment = VBString(IDTag.sComment)
End Property
Public Property Let Comment(ByVal pComment As String)
    IDTag.sComment = pComment
End Property
Public Property Get track() As Byte
    track = IDTag.bTrack
End Property
Public Property Let track(ByVal pTrack As Byte)
    IDTag.bTrack = pTrack
End Property
Public Property Get genre() As Byte
    genre = IDTag.bGenre
End Property
Public Property Let genre(ByVal pGenre As Byte)
    IDTag.bGenre = pGenre
End Property
Public Function ReadTag(Optional ByRef Version As Single) As Boolean
    Dim b() As Byte, t As String
    Dim h As Long, c As Long
    
    ReDim b(TAGSIZE - 1) As Byte
    
    h = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ, CLng(0), OPEN_EXISTING, 0, 0)
    If h = INVALID_HANDLE_VALUE Then GoTo Failure
    
    c = SetFilePointer(h, -TAGSIZE, ByVal 0, FILE_END)
    If c < 0 Then GoTo Failure
    
    If HasTag(h) = False Then GoTo Failure
    
    Call ReadFile(h, b(0), UBound(b) + 1, 0, CLng(0))
    Call CloseHandle(h)
    
    t = ""
    For c = 0 To UBound(b)
        t = t & Chr$(b(c))
    Next c
    
    Erase b
    
    With IDTag
        .sID = Mid$(t, 1, 3)
        .sTitle = Mid$(t, 4, 30)
        .sArtist = Mid$(t, 34, 30)
        .sAlbum = Mid$(t, 64, 30)
        .sYear = Mid$(t, 94, 4)
        .bNull = CByte(Asc(Mid$(t, 126, 1)))
        If .bNull = 0 Then
            .sComment = Mid$(t, 98, 28)
            .bTrack = CByte(Asc(Mid$(t, 127, 1)))
            Version = 1.1
        Else
            .sComment = Mid$(t, 98, 30)
            .bTrack = CByte(0)
            Version = 1
        End If
        .bGenre = CByte(Asc(Mid$(t, 128, 1)))
    End With
    
    ReadTag = True
    
    Exit Function
    
Failure:
    Erase b
    Call CloseHandle(h)
    
    Init
    ReadTag = False
End Function

Public Function WriteTag(Optional ByVal Version As Single = 1.1) As Boolean
    Dim t As String
    Dim h As Long
    
    h = CreateFile(sFile, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE, CLng(0), OPEN_EXISTING, 0, 0)
    If h = INVALID_HANDLE_VALUE Then GoTo Failure
    
    If HasTag(h) = False Then Call CreateTagBuffer(h)
    
    IDTag.sID = "TAG"
    t = IDTag.sID & String_(IDTag.sTitle, 30) & String_(IDTag.sArtist, 30) & _
        String_(IDTag.sAlbum, 30) & String_(IDTag.sYear, 4)
    If Version = 1.1 Then
        IDTag.bNull = 0
        t = t & String_(IDTag.sComment, 28) & Chr$(IDTag.bNull) & Chr$(IDTag.bTrack)
    Else
        t = t & String_(IDTag.sComment, 30)
    End If
    t = t & Chr$(IDTag.bGenre)
    
    Call SetFilePointer(h, -TAGSIZE, ByVal 0, FILE_END)
    Call WriteFile(h, ByVal t, TAGSIZE, 0, ByVal CLng(0))
    Call CloseHandle(h)
    
    WriteTag = True
    
    Exit Function
    
Failure:
    Call CloseHandle(h)
    
    Init
    WriteTag = False
End Function

Public Function RemoveTag() As Boolean
    Dim h As Long, c As Long
    
    h = CreateFile(sFile, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE, CLng(0), OPEN_EXISTING, 0, 0)
    If h = INVALID_HANDLE_VALUE Then GoTo Failure
    
    c = SetFilePointer(h, -TAGSIZE, ByVal 0, FILE_END)
    If c < 0 Then GoTo Failure
    
    If HasTag(h) = False Then GoTo Failure
    
    Call SetEndOfFile(h)
    Call CloseHandle(h)
    
    Init
    RemoveTag = True
    
    Exit Function
    
Failure:
    Call CloseHandle(h)
    
    Init
    RemoveTag = False
End Function

Private Sub Init()
    With IDTag
        .sTitle = ""
        .sArtist = ""
        .sAlbum = ""
        .sYear = ""
        .sComment = ""
        .bNull = 0
        .bTrack = 0
        .bGenre = 255
    End With
End Sub

Private Function HasTag(ByVal file As Long) As Boolean
    Dim b() As Byte, t As String
    Dim c As Long, p As Long
    
    ReDim b(2) As Byte
    
    p = SetFilePointer(file, 0, ByVal 0, FILE_CURRENT)
    
    Call SetFilePointer(file, -TAGSIZE, ByVal 0, FILE_END)
    Call ReadFile(file, b(0), UBound(b) + 1, 0, CLng(0))
    
    t = ""
    For c = 0 To UBound(b)
        t = t & Chr$(b(c))
    Next c
    
    Erase b
    Call SetFilePointer(file, p, ByVal 0, FILE_BEGIN)
    
    HasTag = (UCase$(t) = "TAG")
End Function

Private Sub CreateTagBuffer(ByVal file As Long)
    Call SetFilePointer(file, TAGSIZE, ByVal 0, FILE_END)
    Call SetEndOfFile(file)
End Sub

Private Function String_(ByVal CString As String, ByVal Length As Long) As String
    Dim s As String
    
    If Len(CString) >= Length Then
        s = left$(CString, Length)
    Else
        s = CString & String$(Length - Len(CString), 0)
    End If
    
    String_ = s
End Function

Private Function VBString(ByVal CString As String) As String
    Dim i As Long
    
    i = InStr(1, CString, vbNullChar)
    
    If i = 0 Then
        VBString = CString
    Else
        VBString = left$(CString, i - 1)
    End If
End Function
