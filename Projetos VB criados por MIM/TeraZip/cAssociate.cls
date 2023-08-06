VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAssociate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Title As String
Public Class As String
Public Extension As String
Public Command As String
Public DefaultIcon As String


Private Declare Function RegCreateKey& Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey&, ByVal lpszSubKey$, hKey&)
Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)
Private Declare Function RegSetValueEx& Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal Reserved&, ByVal dwType&, ByVal lpData$, ByVal cbData&)
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey&, ByVal lpSubKey$)
Private Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Private Declare Function RegEnumKeyEx& Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey&, ByVal dwIndex&, ByVal lpName As String, lpcbName&, lpReserved&, ByVal lpClass$, lpcbClass&, lpftLastWriteTime As Any)



Private Const ERROR_SUCCESS = 0&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 256&
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2

Private sKey As String
Private KeyVal As String
Private RC As Long
Private hKey As Long

Public Sub Associate(Optional SubKey As String = "shell\open\command")
    If Title = "" Then
        Err.Raise 898 + 24, "Associate Class", "Associate Class - Application title not specified"
        Exit Sub
    ElseIf Class = "" Then
        Err.Raise 898 + 28, "Associate Class", "Associate Class - Application Class name not specified"
        Exit Sub
    ElseIf Command = "" Then
        Err.Raise 898 + 32, "Associate Class", "Associate Class - Command line of the program not specified"
        Exit Sub
    End If
    
    sKey = Class
    KeyVal = Title
    RC& = RegCreateKey&(HKEY_CLASSES_ROOT, sKey, hKey&)
    RC& = RegSetValue&(hKey&, "", REG_SZ, KeyVal, 0&)
    
    KeyVal = Class
    RC& = RegCreateKey&(HKEY_CLASSES_ROOT, Extension, hKey&)
    RC& = RegSetValue&(hKey&, "", REG_SZ, KeyVal, 0&)
       
    KeyVal = Command & " %1"
    RC& = RegCreateKey&(HKEY_CLASSES_ROOT, sKey, hKey&)
    RC& = RegSetValue&(hKey&, SubKey, REG_SZ, KeyVal, MAX_PATH)
    RC& = RegSetValue&(hKey&, "DefaultIcon", REG_SZ, DefaultIcon, Len(DefaultIcon))
    
End Sub

