Attribute VB_Name = "modFAssociation"
Public Type mnuCommands
 Captions As New Collection
 Commands As New Collection
End Type

Public Type filetype
 Commands As mnuCommands
 extension As String
 ProperName As String
 FullName As String
 ContentType As String
 IconPath As String
 IconIndex As Integer
End Type

Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000

Public Declare Function RegCloseKey Lib _
"advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib _
"advapi32" Alias "RegCreateKeyA" (ByVal _
hKey As Long, ByVal lpszSubKey As String, _
phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib _
"advapi32" Alias "RegSetValueExA" (ByVal _
hKey As Long, ByVal lpszValueName As String, _
ByVal dwReserved As Long, ByVal fdwType As _
Long, lpbData As Any, ByVal cbData As Long) As Long

Public Sub CreateExtension(newfiletype As filetype)

Dim IconString As String
Dim Result As Long, Result2 As Long, ResultX As Long
Dim ReturnValue As Long, HKeyX As Long
Dim cmdloop As Integer

IconString = newfiletype.IconPath & "," & _
newfiletype.IconIndex

If Left$(newfiletype.extension, 1) <> "." Then _
newfiletype.extension = "." & newfiletype.extension

RegCreateKey HKEY_CLASSES_ROOT, _
newfiletype.extension, Result
ReturnValue = RegSetValueEx(Result, "", 0, REG_SZ, _
ByVal newfiletype.ProperName, _
LenB(StrConv(newfiletype.ProperName, vbFromUnicode)))

'Set up content type
If newfiletype.ContentType <> "" Then
ReturnValue = RegSetValueEx(Result, _
"Content Type", 0, REG_SZ, ByVal _
CStr(newfiletype.ContentType), _
LenB(StrConv(newfiletype.ContentType, vbFromUnicode)))
End If

RegCreateKey HKEY_CLASSES_ROOT, _
newfiletype.ProperName, Result

If Not IconString = ",0" Then
RegCreateKey Result, "DefaultIcon", _
Result2 'Create The Key of "ProperNameDefaultIcon"
ReturnValue = RegSetValueEx(Result2, _
"", 0, REG_SZ, ByVal IconString, _
LenB(StrConv(IconString, vbFromUnicode)))
'Set The Default Value for the Key
End If

ReturnValue = RegSetValueEx(Result, _
"", 0, REG_SZ, ByVal newfiletype.FullName, _
LenB(StrConv(newfiletype.FullName, vbFromUnicode)))
RegCreateKey Result, ByVal "Shell", ResultX

'Create neccessary subkeys for each command
For cmdloop = 1 To newfiletype.Commands.Captions.Count
RegCreateKey ResultX, ByVal _
newfiletype.Commands.Captions(cmdloop), Result
RegCreateKey Result, ByVal "Command", Result2
Dim CurrentCommand$
CurrentCommand = newfiletype.Commands.Commands(cmdloop)
ReturnValue = RegSetValueEx(Result2, _
"", 0, REG_SZ, ByVal CurrentCommand$, _
LenB(StrConv(CurrentCommand$, vbFromUnicode)))
RegCloseKey Result
RegCloseKey Result2
Next

RegCloseKey Result2
End Sub



