VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type OPENFILENAME
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
Private Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Private Const OFN_EXPLORER = &H80000 ' new look commdlg
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0


Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long


Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long


'Function SaveDialog(Form1 As Form, Filter As String, title As String, InitDir As String) As String
    
 '   Dim ofn As OPENFILENAME
  '  Dim a As Long
  '  ofn.lStructSize = Len(ofn)
  '  ofn.hwndOwner = Form1.hWnd
  '  ofn.hInstance = App.hInstance
  '  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"


    'For a = 1 To Len(Filter)
     '   If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    'Next
    'ofn.lpstrFilter = Filter
    'ofn.lpstrFile = Space$(254)
    'ofn.nMaxFile = 255
    'ofn.lpstrFileTitle = Space$(254)
    'ofn.nMaxFileTitle = 255
    'ofn.lpstrInitialDir = InitDir
    'ofn.lpstrTitle = title
    'ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
    'a = GetSaveFileName(ofn)


    'If (a) Then
     '   SaveDialog = Trim$(ofn.lpstrFile)
    'Else
    '    SaveDialog = ""
    'End If
'End Function


'Function OpenDialog(Form1 As Form, Filter As String, title As String, InitDir As String) As String
    
 '   Dim ofn As OPENFILENAME
  '  Dim a As Long
   ' ofn.lStructSize = Len(ofn)
    'ofn.hwndOwner = Form1.hWnd
    'ofn.hInstance = App.hInstance
    'If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"


    'For a = 1 To Len(Filter)
    '    If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    'Next
    'ofn.lpstrFilter = Filter
    'ofn.lpstrFile = Space$(254)
    'ofn.nMaxFile = 255
    'ofn.lpstrFileTitle = Space$(254)
    'ofn.nMaxFileTitle = 255
    'ofn.lpstrInitialDir = InitDir
    'ofn.lpstrTitle = title
    'ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    'a = GetOpenFileName(ofn)


    'If (a) Then
    '    OpenDialog = Trim$(ofn.lpstrFile)
    'Else
     '   OpenDialog = ""
    'End If
'End Function

