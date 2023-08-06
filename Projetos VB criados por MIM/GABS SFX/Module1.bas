Attribute VB_Name = "Module1"
Option Explicit

Public Type BrowseInfo
    hwndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib "Shell32.dll" (bBrowse As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "Shell32.dll" (ByVal lItem As Long, ByVal sDir As String) As Long

' Let the user browse for a directory. Return the
' selected directory. Return an empty string if
' the user cancels.
Public Function BrowseForDirectory() As String
Dim browse_info As BrowseInfo
Dim item As Long
Dim dir_name As String
   
   browse_info.hwndOwner = Form2.hWnd
   browse_info.pidlRoot = 0
   browse_info.sDisplayName = Space$(260)
   browse_info.sTitle = "Select Directory"
   browse_info.ulFlags = 1 ' Return directory name.
   browse_info.lpfn = 0
   browse_info.lParam = 0
   browse_info.iImage = 0
   
   item = SHBrowseForFolder(browse_info)
   If item Then
       dir_name = Space$(260)
       If SHGetPathFromIDList(item, dir_name) Then
           BrowseForDirectory = Left(dir_name, InStr(dir_name, Chr$(0)) - 1)
       Else
           BrowseForDirectory = ""
       End If
   End If
End Function



