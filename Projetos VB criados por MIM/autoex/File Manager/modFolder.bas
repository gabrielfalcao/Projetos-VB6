Attribute VB_Name = "modFolder"
'----------------------------------------------------------------------------------
'Module     : ArielBrowseFolder
'Description: Browse For Folder code module
'Version    : V1.00 Sep 2000
'Release    : VB6
'Copyright  : © Tom De Lange, 2000
'e-mail     : tomdl@attglobal.net
'----------------------------------------------------------------------------------
'V1.00    Sep 00 Original version
'----------------------------------------------------------------------------------
'Credits:
'All code obtained from www.planet-source-code.com
'Per Andersson, FireStorm@GoToMy.com, www.FireStormEntertainment.cjb.net
'Roman Blachman, romaz@inter.net.il
'Stephen Fonnesbeck, steev@xmission.com, http://www.xmission.com/~steev
'Max Raskin, www.planet-source-code.com
'----------------------------------------------------------------------------------
'Notes:
'----------------------------------------------------------------------------------
Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O

Private mProper As Boolean

'---------------------------------------------
'Api Structures
'---------------------------------------------
'BrowseInfo used by SHBrowseForFolder API call
'---------------------------------------------
'hWndOwner      Handle to the owner window for the dialog box
'pIdlRoot       Address of an ITEMIDLIST structure specifying the
'               location of the root folder from which to browse.
'               Only the specified folder and its subfolders appear
'               in the dialog box. This member can be NULL; in that case,
'               the namespace root (the desktop folder) is used.
'pszDisplayName Address of a buffer to receive the display name of the
'               folder selected by the user. The size of this buffer
'               is assumed to be MAX_PATH bytes.
'               Some developers declares this parm as String
'lpszTitle      Address of a null-terminated string that is displayed
'               above the tree view control in the dialog box. This
'               string can be used to specify instructions to the user.
'               Normally declared as a Long. String is easier.
'ulFlags        Flags specifying the options for the dialog box.
'               This member can include zero or a combination of the
'               CSIDL type flags
'lpfnCallback   Address of an application-defined function that the
'               dialog box calls when an event occurs. For more information,
'               see BrowseCallbackProc(). This member can be NULL.
'lParam         Application-defined value that the dialog box passes
'               to the callback function, if one is specified.
'iImage         Variable to receive the image associated with the
'               selected folder. The image is specified as an index
'               to the system image list. Other contributers have used
'               a Long declaration here (Why?). Shell32 specifies Integer
Private Type BROWSEINFO
  hWndOwner       As Long
  pidlRoot        As Long
  pszDisplayName  As Long
  lpszTitle       As String
  ulFlags         As Long
  lpfnCallback    As Long
  lParam          As Long
  iImage          As Long
End Type

Private Type OPENFILENAME
  lStructSize As Long           'Length of structure, in bytes
  hWndOwner As Long             'Window that owns the dialog, or NULL
  hInstance As Long             'Handle of mem object containing template (not used)
  lpstrFilter As String         'File types/descriptions, delimited with chr(0), ends with 2xchr(0)
  lpstrCustomFilter As String   'Filters typed in by user
  nMaxCustFilter As Long        'Length of CustomFilter, min 40x chars
  nFilterIndex As Long          'Filter Index to use (1,2,etc) or 0 for custom
  lpstrFile As String           'Initial file/returned file(s), delimited with chr(0) for multi files
  nMaxFile As Long              'Size of Initial File string, min 256
  lpstrFileTitle As String      'File.ext excluding path
  nMaxFileTitle As Long         'Length of FileTitle
  lpstrInitialDir As String     'Initial file dir, null for current dir
  lpstrTitle As String          'Title bar of dialog
  flags As Long                 'See OFN_Flags
  nFileOffset As Integer        'Offset to file name in full path, 0-based
  nFileExtension As Integer     'Offset to file ext in full path, 0-based (excl '.')
  lpstrDefExt As String         'Default ext appended, excl '.', max 3 chars
  lCustData As Long             'Appl defined data for lpfnHook
  lpfnHook As Long              'Pointer to hook procedure
  lpTemplateName As String      'Template Name (not used)
End Type

'These structures are also used by the BrowseInfo (Shell32) function
Private Type SHItemID
  cb      As Long
  abID    As Byte
End Type

Private Type ItemIDList
  mkid  As SHItemID
End Type

'---------------------------------------------------
'Api Function Declarations
'---------------------------------------------------
'FolderBrowsing
'SHBrowseForFolder - Executes Browse For Folder Dialog
'SHGetPathFromIDList - Converts ID List (pidl) to String (Path)
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" _
        (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" _
        (ByVal pIdList As Long, _
        ByVal pszPath As String) As Long
Private Declare Function SHGetFolderLocation Lib "shell32" _
        (hWnd As Long, _
        nFolder As Long, _
        hToken As Long, _
        dwReserved As Long, _
        ppidl As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" _
        (ByVal hWndOwner As Long, _
        ByVal nFolder As Long, _
        pidl As ItemIDList) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" _
        (ByVal szPath As String) As Long
        
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'The following API function is not used, but provided for reference purposes
'Note that is requires the shfolder.dll, distributed with IE4/5/Win98
'Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" _
'        (ByVal hWndOwner As Long, _
'        ByVal nFolder As Long, _
'        ByVal hToken As Long, _
'        ByVal dwFlags As Long, _
'        ByVal pszPath As String) As Long

'-----------------------------------------
'String & Memory handling
'-----------------------------------------
'lstrcat API function appends a string to another
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
'-----------------------------------------
'Windows Messaging for CallbackProc
'-----------------------------------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, _
          ByVal wMsg As Long, _
          ByVal wParam As Long, _
          lParam As Any) As Long
          'ByVal lParam As String) As Long


'-------------------------------------------------------
'API Constants
'-------------------------------------------------------
'Max Length of long file name
Private Const MAX_PATH = 260

'Flag Constants of the BrowseForFolder API function
Private Const BIF_RETURNONLYFSDIRS = &H1       'Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed.
Private Const BIF_DONTGOBELOWDOMAIN = &H2      'Do not include network folders below the domain level in the tree view control.
Private Const BIF_STATUSTEXT = &H4             'Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box.
Private Const BIF_RETURNFSANCESTORS As Long = &H8     'Only return file system ancestors. If the user selects anything other than a file system ancestor, the OK button is grayed.
Private Const BIF_EDITBOX As Long = &H10              'Version 4.71. The browse dialog includes an edit control in which the user can type the name of an item.
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000  'Only return computers. If the user selects anything other than a computer, the OK button is grayed.
Private Const BIF_BROWSEFORPRINTER As Long = &H2000   'Only return printers. If the user selects anything other than a printer, the OK button is grayed.
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000 'The browse dialog will display files as well as folders.
'private const BIF_VALIDATE             Version 4.71. If the user types an invalid name into the edit box, the browse dialog will call the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message. This flag is ignored if BIF_EDITBOX is not specified.

'Constants used by the CallBack Procedure
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

'Constants used by the CheckFolder procedure
'Not used in this implementation
Private Const CSIDL_LOCAL_APPDATA = &H1C&
Private Const CSIDL_FLAG_CREATE = &H8000&
Private Const E_INVALIDARG = &H80070057 ' Invalid CSIDL Value
Private Const C_OK = &H0                ' Success
Private Const C_FALSE = &H1             ' The Folder is valid, but does not exist
Private Const SHGFP_TYPE_CURRENT = 0
Private Const SHGFP_TYPE_DEFAULT = 1

Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_EXPLORER = &H80000
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000

Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

'For determining which part of the border to draw
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_SOFT = &H1000
Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)

'For drawing borders with the DrawEdge() function
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_BUMP = BDR_RAISEDOUTER Or BDR_SUNKENINNER
Private Const EDGE_ETCHED = BDR_SUNKENOUTER Or BDR_RAISEDINNER
Private Const EDGE_RAISED = BDR_RAISEDOUTER Or BDR_RAISEDINNER
Private Const EDGE_SUNKEN = BDR_SUNKENOUTER Or BDR_SUNKENINNER

Private Function IsLetter(Txt As String) As Boolean
'-----------------------------------------
'Tests if first character of string is a letter
'-----------------------------------------
Dim c As String
c = Left(Txt, 1)
Select Case Asc(c)
Case 65 To 90, 97 To 122
  IsLetter = True
Case Else
  IsLetter = False
End Select

End Function
Public Function StringToProper(ByVal Txt As String, Optional ByVal Sep As String = ":\ ") As String
'-------------------------------------------------------
'Returns a proper string with first letters capitalized
'Capitalises 1st letter after space, period or start
'-------------------------------------------------------
Dim n, i, s As String, bSp As Boolean, c As String
Dim aSt() As String, St As String, ni, j

'First determine if any lower case letters are present
'If so, then don't make any changes as the folder is
'obviously already a long file name
'This prevents conversion of e.g. Heretic II folder

aSt = Split(Txt, "\")
ni = UBound(aSt)
For j = 0 To ni
  St = aSt(j)
  If UCase(St) = St Then
    s = ""
    n = Len(St)
    bSp = True
    For i = 1 To n
      c = Mid(St, i, 1)
      If InStr(Sep, c) > 0 Then
        bSp = True
      Else
        If IsLetter(c) Then
          If bSp Then
            c = UCase(c)
            bSp = False
          Else
            c = LCase(c)
          End If
        End If
      End If
      s = s + c
    Next
    aSt(j) = s
  End If
Next
StringToProper = Join(aSt, "\")

End Function

Public Function BrowseFolder(ByVal hWnd As Long, lFlags As Long, Optional ByVal sTitle As String, Optional vSelPath As Variant, Optional vTopFolder As Variant, Optional ByVal bProper As Boolean = False) As String
'------------------------------------------------------------------------
'Activate the folder browser dialog
'Main routine calling the API function SHBrowseForFolder()
'hWnd       OwnerWindow.hWnd handle   Long
'Title      Instructions to the user  String
'SelPath    Preselected folder path   String or Enum SpecialFolders
'TopFolder  Topmost folder in dialog  String or Enum SpecialFolders
'Return     Full path of Selected folder, or "" if unsuccesfull
'------------------------------------------------------------------------
Dim lpIdList As Long            'pointer to item identifier list of selected path
Dim szTitle As String
Dim sPath As String
Dim uBrowseInfo As BROWSEINFO
Dim lRet As Long
Dim sRet As String

'Anderson
Dim lItemIDList As ItemIDList

mProper = bProper
'mCurrentFolder = CStr(vSelPath) & vbNullChar

szTitle = sTitle
 
With uBrowseInfo
  'The desktop/form or userctrl will own the dialog
  .hWndOwner = hWnd
  'This will be the dialog's root folder. If missing, use Desktop
  If IsMissing(vTopFolder) Then
    vTopFolder = &H0
  End If
  If Not IsNumeric(vTopFolder) Then
    'String Path passed in vTopFolder, convert to long CSID
    .pidlRoot = SHSimpleIDListFromPath(CStr(vTopFolder))
  Else
    'Long CSIDL Special Folder Constant passed
    lRet = SHGetSpecialFolderLocation(ByVal hWnd, ByVal vTopFolder, lItemIDList)
    .pidlRoot = lItemIDList.mkid.cb
  End If
  'Set the dialog's prompt string (title)
  .lpszTitle = sTitle
  '.lpszTitle = lstrcat(szTitle, Chr(0)) 'Append title string to a long value
  '.lpszTitle = lstrcat(szTitle, "")      'Append title string to a null char
  
  'Set the Flags
  .ulFlags = lFlags
  '.ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
  '.ulFlags = BIF_RETURNONLYFSDIRS + BIF_RETURNFSANCESTORS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
  '.ulFlags = BIF_STATUSTEXT
  
  'Obtain and set the address of the callback function
  .lpfnCallback = FunctionPointer(AddressOf BrowseCallbackProc)
  'Obtain and set the pidl of the pre-selected folder
  If IsMissing(vSelPath) Then
    'Nothing passed in, set to the topmost folder
    .lParam = .pidlRoot
  ElseIf Not IsNumeric(vSelPath) Then
    'String Path passed, convert the path to the ID
    'If empty string, the root is used
    If vSelPath = "" Then
      .lParam = .pidlRoot
    Else
      'Code used by Blachman:-
      'Dim selectedPathPointer As Long
      'SelectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1) ' Allocate a string
      'CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1 ' Copy the path to the string
      '.lParam = SelectedPathPointer ' The folder to preselect
      .lParam = SHSimpleIDListFromPath(CStr(vSelPath))
    End If
  Else
    'Long CSIDL Special Folder Constant passed in
    lRet = SHGetSpecialFolderLocation(ByVal hWnd, ByVal CLng(vSelPath), lItemIDList)
    .pidlRoot = lItemIDList.mkid.cb

    
    
    
    
    ''First check if the special folder exists
    'sRet = CheckSpecialFolder((vSelPath))
    'If sRet <> "" Then
    '  .lParam = SHSimpleIDListFromPath(sRet)

      '.lParam= vSelPath  'If there is any valid ID use it
      'lRet = SHGetSpecialFolderLocation(ByVal hWnd, ByVal CLng(vSelPath), lItemIDList)
      '.lParam = lItemIDList.mkid.cb
    'Else
    '  .lParam = .pidlRoot     'Make same as root
    'End If
  End If
End With
 
'Finally, show the FolderBrowse dialog.
'Control doesn't return until the dialog is closed.
'The BrowseCallbackProc will receive all browse dialog specific messages while
'the dialog is open. pidlRet will contain the pidl of the selected folder
'if the dialog is not cancelled.
lpIdList = SHBrowseForFolder(uBrowseInfo)
 
If (lpIdList) Then
  'Get the path from the selected folder's pidl returned
  'from the SHBrowseForFolder call. Returns True on success,
  'sPath must be pre-allocated!
  sPath = Space(MAX_PATH)
  If SHGetPathFromIDList(lpIdList, sPath) Then
    'Return the path
    sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
  End If
  'Free the memory the shell allocated for the pIdList
  Call CoTaskMemFree(lpIdList)
End If

'Free the memory the shell allocated for the pre-selected folder.
Call CoTaskMemFree(uBrowseInfo.lParam)
'Code used by Blachman
'Call LocalFree(SelectedPathPointer) ' Free the string from the memory
 
 
'If Cancel button was clicked, an error occured or a Non File System Folder
'was selected. Then Path = ""
BrowseFolder = sPath

End Function
Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------------
'This function is used by the Browse Folder dialog to call back for instructions
'
'-----------------------------------------------------------------------------------------------------------------------------
Dim lpIdList As Long
Dim lRet As Long
Dim sBuffer As String

On Error Resume Next  'Sugested by MS to prevent an error from
                      'propagating back into the calling process.
Select Case uMsg
Case BFFM_INITIALIZED   '&H1
  'Anderson
  Call SendMessage(hWnd, BFFM_SETSELECTIONA, False, ByVal lpData)
  'Blachman
  'Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
  'Fonnesbeck
  'Sent the CurrentPath directly from here
  'Call SendMessage(hWnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
  'Ariel:Fonnesbeck
  'Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal mCurrentFolder)

Case BFFM_SELCHANGED
  'Fonnesbeck added this part.
  'It displays the current selection below the title string
  sBuffer = Space(MAX_PATH)
  lRet = SHGetPathFromIDList(lParam, sBuffer)
  If lRet = 1 Then
    If mProper Then
      sBuffer = StringToProper(sBuffer)
    End If
    Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, ByVal sBuffer)
  End If
End Select
BrowseCallbackProc = 0      'Fonnesbeck

End Function

Private Function FunctionPointer(FunctionAddress As Long) As Long
'---------------------------------------------------
'Allows the address of a function to be conveyed
'as a pointer, as required
'by the BrowseFolder dialog
'Returns a pointer to the address of the function
'using AddressOf operator (VB5/6 only!)
'---------------------------------------------------
FunctionPointer = FunctionAddress

End Function

Function ValidateFolder(ByVal Folder As String) As String
'--------------------------------------------------------
'Validates the folder name, by removing the last "\"
'if present
'Network folder: \\Tom\data\
'Normal folder : C:\Program Files\
'Root folder   : C:\
'--------------------------------------------------------
If Right(Folder, 1) = "\" Then
  Folder = Left(Folder, Len(Folder) - 1)
  If Right(Folder, 1) = ":" Then
    Folder = Folder & "\"
  End If
End If
ValidateFolder = Folder

End Function

Public Function ShowSaveDialog(sFilter As String, ByRef nFilterIndex As Long, sTitle As String, ByRef sPath As String, ByRef sFileName As String) As String
'--------------------------------------------------------------------------------
'Show the Save File dialog
'Filter       : Filter List i.e. "All Files (*.*)|*.*"
'FilterIndex  : Element no in Filter list, starting from 1
'               Updated on exit
'Title        : Dialog Title
'InitialPath  : Default Folder
'FileName     : On entry contains the initial file name excl path
'               On exit contains file name excl path, i.e. "command.com"
'--------------------------------------------------------------------------------
Dim ofn As OPENFILENAME
Dim i

ofn.lStructSize = Len(ofn)
ofn.hWndOwner = frmPack.hWnd
ofn.hInstance = App.hInstance
If Right(sFilter, 1) <> "|" Then sFilter = sFilter + "|"
For i = 1 To Len(sFilter)
  If Mid(sFilter, i, 1) = "|" Then
    Mid(sFilter, i, 1) = Chr$(0)
  End If
Next
ofn.lpstrFilter = sFilter
ofn.nFilterIndex = nFilterIndex
ofn.lpstrFile = Left(sFileName & Space(254), 254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space(254) 'on exit contains the filename
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = sPath
ofn.lpstrTitle = sTitle
ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
i = GetSaveFileName(ofn)

If (i) Then
  ShowSaveDialog = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
  sPath = ValidateFolder(Left(ofn.lpstrFile, ofn.nFileOffset))
  sFileName = Left(ofn.lpstrFileTitle, InStr(ofn.lpstrFileTitle, vbNullChar) - 1)
  nFilterIndex = ofn.nFilterIndex
Else
  ShowSaveDialog = ""
End If

End Function

Public Function ShowOpenDialog(ByVal sFilter As String, ByRef nFilterIndex As Long, ByVal sTitle As String, ByRef sPath As String, ByRef sFileName As String) As String
'------------------------------------------------------------------------------------
'Displays the Open File Dialog
'Filter       : Filter List i.e. "All Files (*.*)|*.*"
'FilterIndex  : Element no in Filter list, starting from 1
'               Updated on exit
'Title        : Dialog Title
'Path         : Default Folder on entry, selected folder on exit
'FileName     : On entry contains the initial file name excl path
'               On exit contains file name excl path, i.e. "command.com"
'------------------------------------------------------------------------------------
Dim ofn As OPENFILENAME
Dim i

ofn.lStructSize = Len(ofn)
ofn.hWndOwner = frmPack.hWnd
ofn.hInstance = App.hInstance

'Ensure | character added to end
If Right(sFilter, 1) <> "|" Then
  sFilter = sFilter + "|"
End If
'Replace the | character with chr(0)
For i = 1 To Len(sFilter)
  If Mid(sFilter, i, 1) = "|" Then
    Mid(sFilter, i, 1) = Chr(0)
  End If
Next
ofn.lpstrFilter = sFilter
ofn.nFilterIndex = nFilterIndex
'ofn.lpstrCustomFilter      'Not Used
'ofn.nMaxCustomFilter       'Not Used
ofn.lpstrFile = Left(sFileName & Space(254), 254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space(254) 'on exit contains the filename
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = sPath
ofn.lpstrTitle = sTitle
ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST 'Or OFN_EXPLORER Or OFN_LONGNAMES
i = GetOpenFileName(ofn)

If (i) Then
  ShowOpenDialog = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
  sPath = ValidateFolder(Left(ofn.lpstrFile, ofn.nFileOffset))
  sFileName = Left(ofn.lpstrFileTitle, InStr(ofn.lpstrFileTitle, vbNullChar) - 1)
  nFilterIndex = ofn.nFilterIndex
Else
  ShowOpenDialog = ""
End If

End Function


