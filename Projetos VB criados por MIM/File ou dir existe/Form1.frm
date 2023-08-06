VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   1260
      TabIndex        =   1
      Top             =   1545
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1035
      TabIndex        =   0
      Text            =   "C:\MUSICAO.zip"
      Top             =   780
      Width           =   1905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------
' FUNCTION: FileExists
'
' Determines whether the specified File name exists.
'
' IN: [Path$] - name of File to check for(with full path)
'
' Returns: True if the File exists, False otherwise
'-----------------------------------------------------------
Function FileExist(path$) As Integer
    Dim x
    x = FreeFile
    On Error Resume Next
    Open path$ For Input As x
    FileExist = IIf(Err = 0, True, False)
    Close x
    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: DirExists
'
' Determines whether the specified directory name exists.
'
' IN: [strDirName] - name of directory to check for
'
' Returns: True if the directory exists, False otherwise
'-----------------------------------------------------------
'
Public Function DirExists(ByVal strDirName As String) As Boolean
    Const gstrNULL$ = ""
    Dim strDummy As String

    strDummy = Dir$(strDirName, vbDirectory)
    If strDummy = gstrNULL$ Then
        DirExists = False
    Else
        DirExists = True
    End If
End Function

'---------------------------------------------------------------------------------
'Function: To extract the path of the any file
'Process:
'   if input Filename="c:\windows\desktop\sample.txt"
'   returns"c:\windows\desktop"
'---------------------------------------------------------------------------------
Public Function ExtractPath(Filename As String) As String
Dim l As Integer
Dim tempchar As String
l = Len(Filename)
While l > 0
    tempchar = Mid(Filename, l, 1)  'trapping the last '\' char to retrieve only the path of the setup file
    If tempchar = "\" Then
        ExtractPath = Mid(Filename, 1, l - 1)
        Exit Function
    End If
    l = l - 1
Wend
End Function

'-------------------------------------------------------------------------------
'Function:- To find whether a particular type of file exist
'   Process : for example here when we pass fileext as '.mdb' it checks for access database files only
'   when atleast one file exist of particular type returns true else false
' set reference to Microsoft scripting run time under
' project-> Reference
'---------------------------------------------------------------------------------
Public Function SpecificFileExists(filepath As String, FileExt As String) As Boolean

Dim fso As New Scripting.FileSystemObject
Dim folder As folder
Dim Filename As File
Dim path As String
path = filepath
If Len(Trim(path)) <> 0 Then
    Set folder = fso.GetFolder(path)    'seting folder to variable for easy manipulation
    For Each Filename In folder.Files   'retrieving files one by one
         If StrComp(Right(Filename.Name, 4), FileExt, vbTextCompare) = 0 Then
            SpecificFileExists = True
            Exit Function
        End If
    Next
End If
    SpecificFileExists = False 'when no mdb is found in the folder
End Function



Private Sub Command1_Click()
If FileExist(Text1.Text) = True Then
MsgBox "arquivo existe!"
Else
MsgBox "arquivo não existe!"
End If
End Sub

Private Sub Form_Load()

End Sub
