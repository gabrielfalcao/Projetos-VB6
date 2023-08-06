Attribute VB_Name = "ModDrives"
Declare Function GetLogicalDriveStrings Lib "kernel32" Alias _
  "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal _
  lpBuffer As String) As Long
Declare Function GetDriveType Lib "kernel32" Alias _
  "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

