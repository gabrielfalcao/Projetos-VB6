VERSION 5.00
Begin VB.UserControl Zipit 
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Zipit.ctx":0000
   PropertyPages   =   "Zipit.ctx":030A
   ScaleHeight     =   510
   ScaleWidth      =   510
   ToolboxBitmap   =   "Zipit.ctx":0320
End
Attribute VB_Name = "Zipit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Private Const MAX_PATH = 260
Private Const LocalFileHeaderSig = &H4034B50
Private Const CentralFileHeaderSig = &H2014B50
Private Const EndCentralDirSig = &H6054B50
Private Const APP_TITLE = "Richsoft Zipit 1.0"

Private Type ZipFile
  Version As Integer
  Flag As Integer
  CompressionMethod As Integer
  Time As Integer
  Date As Integer
  CRC32 As Long
  CompressedSize As Long
  UncompressedSize As Long
  FileNameLength As Integer
  ExtraFieldLength As Integer
  FileName As String
End Type

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Public Enum ZipLevel
  zipStore = 0
  zipLevel1 = 1
  zipSuperFast = 2
  zipFast = 3
  zipLevel4 = 4
  zipNormal = 5
  zipLevel6 = 6
  zipLevel7 = 7
  zipLevel8 = 8
  zipMax = 9
End Enum

Public Enum ZipAction
  zipDefault = 1
  zipFreshen = 2
  zipUpdate = 3
End Enum

Private Declare Function AddFile Lib "zipit.dll" (ByVal ZipFilename As String, ByVal FileName As String, ByVal StoreDirInfo As Boolean, ByVal DOS83 As Boolean, ByVal Action As Integer, ByVal CompressionLevel As Integer) As Boolean
Private Declare Function ExtractFile Lib "zipit.dll" (ByVal ZipFilename As String, ByVal FileName As String, ByVal ExtrDir As String, ByVal UseDirInfo As Boolean, ByVal Overwrite As Boolean, ByVal Action As Integer) As Boolean
Private Declare Function DeleteFile Lib "zipit.dll" (ByVal ZipFilename As String, ByVal FileName As String) As Boolean
Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Public ZipFiles As New Collection

Private ZipFilename As String
Private CompLevel As ZipLevel
Private ExtrDir As String
Private UseDirInfo As Boolean
Private AddFileAction As ZipAction
Private ExtrFileAction As ZipAction
Private OverwriteFiles As Boolean
Private DOS83Format As Boolean
Private RecurseSubs As Boolean
Private IncludeSysFiles As Boolean

Public Event Change()
Attribute Change.VB_MemberFlags = "200"
Public Event DeleteProgress(Percentage As Integer, FileName As String)
Public Event DeleteComplete(Successful As Long)
Public Event UnzipComplete(Successful As Long)
Public Event UnzipProgress(Percentage As Integer, FileName As String)
Public Event ZipComplete(Successful As Long)
Public Event ZipProgress(Percentage As Integer, FileName As String)

Private Function ConvertWildcards(Files As Collection, ByVal IncludeSysFiles As Boolean, ByVal Recurse As Boolean) As Collection
    Dim i As Long
    Dim R As String
    Dim ret As New Collection
    Dim Path As String
    Dim Buffer As String * MAX_PATH
    Dim Attributes As Integer
    Attributes = vbNormal Or vbReadOnly
    
    If IncludeSysFiles = True Then
        Attributes = Attributes Or vbSystem Or vbHidden
    End If
    
    For i = 1 To Files.Count
        Path = ParsePath(Files(i))
        R = Dir$(Files(i), Attributes)
        Do Until R = ""
            ret.Add Path & R
            R = Dir$()
        Loop
    Next i
    Set ConvertWildcards = ret
End Function

Public Function Delete(Filenames As Collection) As Long
    Dim R As Long
    Dim i As Long
    Dim ZipFile As String
    Dim FileName As String
    On Error Resume Next
    Delete = 0
    
    ZipFile = ZipFilename
    If ZipFile = "" Then
        Delete = 0
        Exit Function
    End If
    
    For i = 1 To Filenames.Count
        Read
        If ZipFiles.Count <> 1 Then
            R = DeleteFile(ZipFile, Filenames(i))
            If R = True Then
                Delete = Delete + 1
            End If
        Else
            Kill ZipFile
            Delete = Delete + 1
        End If
        RaiseEvent DeleteProgress((i / Filenames.Count) * 100, Filenames(i))
        DoEvents
    Next i
    RaiseEvent DeleteComplete(Delete)
    Read
    RaiseEvent Change
End Function

Public Function Add(ByVal Filenames As Collection) As Long
Attribute Add.VB_UserMemId = 0
    Dim FileName As String
    Dim locZipFile As String
    Dim locUseDirInfo As String
    Dim locDOS83Format As Boolean
    Dim locAction As ZipAction
    Dim locCompLevel As ZipLevel
    Dim locRecurse As Boolean
    Dim locIncludeSysFiles As Boolean
    Dim i As Long
    Dim R As Boolean
    
    Add = 0
    locZipFile = ZipFilename
    locUseDirInfo = UseDirInfo
    locDOS83Format = DOS83Format
    locAction = AddFileAction
    locCompLevel = CompLevel
    locRecurse = RecurseSubs
    locIncludeSysFiles = True
    
    If locZipFile = "" Then
        Add = 0
        Exit Function
    End If
    
    If ZipFiles.Count = 0 Then Kill locZipFile
    
    Set Filenames = ConvertWildcards(Filenames, locIncludeSysFiles, locRecurse)
    
    For i = 1 To Filenames.Count
        R = AddFile(locZipFile, Filenames(i), locUseDirInfo, locDOS83Format, locAction, locCompLevel)
        If R = True Then Add = Add + 1
        RaiseEvent ZipProgress((i / Filenames.Count) * 100, Filenames(i))
        DoEvents
    Next i
    RaiseEvent ZipComplete(Add)
    Read
    RaiseEvent Change
End Function

Private Sub AddEntry(zFile As ZipFile)
    Dim xFile As New ZipFileEntry
    With xFile
     .Version = zFile.Version
     .Flag = zFile.Flag
     .CompressionMethod = zFile.CompressionMethod
     .CRC32 = zFile.CRC32
     .FileDateTime = GetDateTime(zFile.Date, zFile.Time)
     .CompressedSize = zFile.CompressedSize
     .UncompressedSize = zFile.UncompressedSize
     .FileNameLength = zFile.FileNameLength
     .FileName = zFile.FileName
     .ExtraFieldLength = zFile.ExtraFieldLength
    End With
    ZipFiles.Add xFile
End Sub

Public Function Extract(Filenames As Collection) As Long
    Dim R As Boolean
    Dim i As Long
    Dim FileName As String
    Dim locZipFile As String
    Dim locUseDirInfo As Boolean
    Dim locOverwrite As Boolean
    Dim locAction As ZipAction
    Dim locExtrDir As String
    
    Extract = 0
    locZipFile = ZipFilename
    locUseDirInfo = UseDirInfo
    locOverwrite = OverwriteFiles
    locAction = ExtrFileAction
    locExtrDir = ExtractDir
    
    If locZipFile = "" Then
        Extract = 0
        Exit Function
    End If
        
    For i = 1 To Filenames.Count
        R = ExtractFile(locZipFile, Filenames(i), locExtrDir, locUseDirInfo, locOverwrite, locAction)
        If R = True Then
            Extract = Extract + 1
        End If
        RaiseEvent UnzipProgress((i / Filenames.Count) * 100, Filenames(i))
        DoEvents
    Next i
    RaiseEvent UnzipComplete(Extract)
End Function

Public Property Get ExtractDir() As String
Attribute ExtractDir.VB_ProcData.VB_Invoke_Property = "ZipitProperties"
    ExtractDir = ExtrDir
End Property
Public Property Let ExtractDir(New_Extractdir As String)
    ExtrDir = New_Extractdir
    PropertyChanged "ExtractDir"
End Property

Private Function ParsePath(Path As String)
Dim A As Integer
    For A = Len(Path) To 1 Step -1
        If Mid$(Path, A, 1) = "\" Then
            ParsePath = Left$(Path, A - 1) & "\"
            Exit Function
        End If
    Next A
End Function

Public Property Get UseDirectoryInfo() As Boolean
    UseDirectoryInfo = UseDirInfo
End Property
Public Property Let UseDirectoryInfo(New_UseDirectoryInfo As Boolean)
    UseDirInfo = New_UseDirectoryInfo
    PropertyChanged "UseDirectoryInfo"
End Property

Public Property Get CompressionLevel() As ZipLevel
    CompressionLevel = CompLevel
End Property
Public Property Let CompressionLevel(New_CompressionLevel As ZipLevel)
    CompLevel = New_CompressionLevel
    PropertyChanged "CompressionLevel"
End Property

Public Property Get FileName() As String
    FileName = ZipFilename
End Property
Public Property Let FileName(New_Filename As String)
Attribute FileName.VB_ProcData.VB_Invoke_PropertyPut = "ZipitProperties"
    Dim R As Long
    Dim i As Long
    ZipFilename = New_Filename
    PropertyChanged "Filename"
    R = Read
    RaiseEvent Change
End Property

Private Function GetDateTime(ZipDate As Integer, ZipTime As Integer) As Date
    Dim R As Long
    Dim FTime As FILETIME
    Dim Sys As SYSTEMTIME
    Dim ZipDateStr As String
    Dim ZipTimeStr As String
    
    R = DosDateTimeToFileTime(CLng(ZipDate), CLng(ZipTime), FTime)
    R = FileTimeToSystemTime(FTime, Sys)
    ZipDateStr = Sys.wDay & "/" & Sys.wMonth & "/" & Sys.wYear
    ZipTimeStr = Sys.wHour & ":" & Sys.wMinute & ":" & Sys.wSecond
    GetDateTime = ZipDateStr & " " & ZipTimeStr
End Function

Public Function Read() As Long
    Dim Sig As Long
    Dim ZipStream As Integer
    Dim Res As Long
    Dim zFile As ZipFile
    Dim Name As String
    Dim i As Integer
    If ZipFilename = "" Then
        Read = 0
        For i = ZipFiles.Count To 1 Step -1
            ZipFiles.Remove i
        Next i
        Exit Function
    End If
    
    For i = ZipFiles.Count To 1 Step -1
        ZipFiles.Remove i
    Next i
    
    ZipStream = FreeFile
    Open ZipFilename For Binary As ZipStream
    Do While True
        Get ZipStream, , Sig
              If Sig = LocalFileHeaderSig Then
                    Get ZipStream, , zFile.Version
                    Get ZipStream, , zFile.Flag
                    Get ZipStream, , zFile.CompressionMethod
                    Get ZipStream, , zFile.Time
                    Get ZipStream, , zFile.Date
                    Get ZipStream, , zFile.CRC32
                    Get ZipStream, , zFile.CompressedSize
                    Get ZipStream, , zFile.UncompressedSize
                    Get ZipStream, , zFile.FileNameLength
                    Get ZipStream, , zFile.ExtraFieldLength
                    
                    Name = String$(zFile.FileNameLength, " ")
                    Get ZipStream, , Name
                    zFile.FileName = Mid$(Name, 1, zFile.FileNameLength)
                    Seek ZipStream, (Seek(ZipStream) + zFile.ExtraFieldLength)
                    Seek ZipStream, (Seek(ZipStream) + zFile.CompressedSize)
                    AddEntry zFile
              Else
                If Sig = CentralFileHeaderSig Or Sig = 0 Then
                    Exit Do
                Else
                  If Sig = EndCentralDirSig Then
                      Exit Do
                  End If
                End If
            End If
        Loop
        Close ZipStream
        Read = ZipFiles.Count
    RaiseEvent Change
End Function

Public Property Get UseDOS83Format() As Boolean
    UseDOS83Format = DOS83Format
End Property
Public Property Let UseDOS83Format(New_UseDOS83Format As Boolean)
    DOS83Format = New_UseDOS83Format
    PropertyChanged "UseDOS83Format"
End Property

Public Property Get RecurseSubFolders() As Boolean
    RecurseSubFolders = RecurseSubs
End Property
Public Property Let RecurseSubFolders(New_RecurseSubFolders As Boolean)
    RecurseSubs = New_RecurseSubFolders
    PropertyChanged "RecurseSubFolders"
End Property

Public Property Get AddAction() As ZipAction
    AddAction = AddFileAction
End Property
Public Property Let AddAction(New_AddAction As ZipAction)
    AddFileAction = New_AddAction
    PropertyChanged "AddAction"
End Property

Public Property Get Overwrite() As Boolean
    Overwrite = OverwriteFiles
End Property
Public Property Let Overwrite(New_Overwrite As Boolean)
    OverwriteFiles = New_Overwrite
    PropertyChanged "Overwrite"
End Property

Public Property Get ExtrAction() As ZipAction
    ExtrAction = ExtrFileAction
End Property
Public Property Let ExtrAction(New_ExtrAction As ZipAction)
    ExtrFileAction = New_ExtrAction
    PropertyChanged "ExtrAction"
End Property


Private Sub UserControl_InitProperties()
 AddAction = zipDefault
 ExtrAction = zipDefault
 CompressionLevel = zipMax
 UseDirInfo = False
 Overwrite = True
 ExtrDir = "c:\"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ExtrDir = PropBag.ReadProperty("ExtractDir", "")
    CompLevel = PropBag.ReadProperty("CompressionLevel", zipMax)
    ZipFilename = PropBag.ReadProperty("Filename", "")
    UseDirInfo = PropBag.ReadProperty("UseDirectoryInfo", True)
    OverwriteFiles = PropBag.ReadProperty("Overwrite", False)
    DOS83Format = PropBag.ReadProperty("UseDOS83Format", False)
    AddFileAction = PropBag.ReadProperty("AddAction", zipDefault)
    ExtrFileAction = PropBag.ReadProperty("ExtrAction", zipDefault)
    RecurseSubs = PropBag.ReadProperty("RecurseSubFolders", False)
    IncludeSysFiles = PropBag.ReadProperty("IncludeSystemFiles", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Filename", ZipFilename, ""
    PropBag.WriteProperty "ExtractDir", ExtrDir, ""
    PropBag.WriteProperty "UseDirectoryInfo", UseDirInfo, True
    PropBag.WriteProperty "CompressionLevel", CompLevel, zipMax
    PropBag.WriteProperty "Overwrite", OverwriteFiles, False
    PropBag.WriteProperty "UseDOS83Format", DOS83Format, False
    PropBag.WriteProperty "AddAction", AddFileAction, zipDefault
    PropBag.WriteProperty "ExtrAction", ExtrFileAction, zipDefault
    PropBag.WriteProperty "RecurseSubFolders", RecurseSubs, False
    PropBag.WriteProperty "IncludeSystemFiles", IncludeSysFiles, True
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 560, 560
End Sub

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
    ShellAbout UserControl.hWnd, APP_TITLE, "Required files: Zipit.dll, Zipdll.dll and Unzdll.dll in SYSTEM FOLDER", UserControl.Picture
End Sub
