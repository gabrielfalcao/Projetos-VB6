Attribute VB_Name = "modIni"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158

Global INI As String
Global winINI As String

Global strMainX     As String
Global strMainY     As String
Global strRateX     As String
Global strRateY     As String
Global bSPLS        As Boolean
Global bSRS         As Boolean
Global bAlwaysTray  As Boolean
Global bMinTray     As Boolean
Global strLOpenPath As String
Global strLSavePath As String
Global intLIndex    As Integer


Function GetINI(strINI As String, strSection As String, strSetting As String, strDefault As String)
    Dim lngReturn As Long, strReturn As String, lngSize As Long
    lngSize = 255
    strReturn = String(lngSize, 0)
    lngReturn = GetPrivateProfileString(strSection, strSetting, strDefault, strReturn, lngSize, winINI)
    If strReturn = "" Then
        GetINI = strDefault
        PutINI strINI, strSection, strSetting, strDefault
    Else
        GetINI = LeftOf(strReturn, Chr(0))
    End If

End Function

Sub PutINI(strINI As String, strSection As String, strLValue As String, strRValue As String)
    Dim lngReturn As Long
    lngReturn = WritePrivateProfileString(strSection, strLValue, strRValue, strINI)
    'MsgBox lngReturn & "..ini"
End Sub

Public Function TF(bVal As Boolean) As Integer
    If bVal Then
        TF = 1
    Else
        TF = 0
    End If
End Function

Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function

Function ReadINI(strSection As String, strSetting As String, strDefault As String)
    Dim lngReturn As Long, strReturn As String, lngSize As Long
    lngSize = 255
    strReturn = String(lngSize, 0)
    lngReturn = GetPrivateProfileString(strSection, strSetting, strDefault, strReturn, lngSize, INI)
    If strReturn = "" Then
        ReadINI = strDefault
        WriteINI strSection, strSetting, strDefault
    Else
        ReadINI = LeftOf(strReturn, Chr(0))
    End If
End Function

Sub WriteINI(strSection As String, strLValue As String, strRValue As String)
    Dim lngReturn As Long
    lngReturn = WritePrivateProfileString(strSection, strLValue, strRValue, INI)
    'MsgBox lngReturn & "..ini"
End Sub

Function LeftOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    
    intPos = InStr(strData, strDelim)
    If intPos Then
        LeftOf = Left(strData, intPos - 1)
    Else
        LeftOf = strData
    End If
End Function

