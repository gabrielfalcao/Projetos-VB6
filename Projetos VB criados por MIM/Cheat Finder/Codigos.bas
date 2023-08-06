Attribute VB_Name = "Codigos"
'*************************************
'Copyright © 2001 by Alexander Anikin
''e-mail: aka@i.com.ua
'http://www.hotmix.kiev.ua
'*************************************
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
"GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Public FileOP As String
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4

Public OFN As OPENFILENAME
Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Public Function fileExist(fileName As String) As Boolean
    Dim l As Long
    
    On Error Resume Next
    
    l = FileLen(fileName)
    
    fileExist = Not (Err.Number > 0)
    
    On Error GoTo 0
End Function



