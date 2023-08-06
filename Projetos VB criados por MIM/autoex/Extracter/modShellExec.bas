Attribute VB_Name = "modShellExec"
      Private Declare Function ShellExecute Lib "shell32.dll" Alias _
      "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As _
      String, ByVal lpszFile As String, ByVal lpszParams As String, _
      ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

      Public Declare Function GetDesktopWindow Lib "user32" () As Long

' hWnd = Window handle to a parent window.
' This window receives any message boxes that an application produces.
'
' lpszOp = Address of a null-terminated string that specifies the operation to perform.
' The following operation strings are valid:
' open, print, explore
' This parameter can be NULL. In that case, the function opens the file specified
' by lpFile.
'
' lpFile = Address of a null-terminated string that specifies the file to open
' or print or the folder to open or explore. The function can open an executable
' file or a document file. The function can print a document file.
'
' lpParameters = If the lpFile parameter specifies an executable file,
' lpParameters is an address to a null-terminated string that specifies
' the parameters to be passed to the application.
' If lpFile specifies a document file, lpParameters should be NULL.
'
' lpDirectory = Address of a null-terminated string that specifies
' the default directory.
'
' nShowCmd = If lpFile specifies an executable file, nShowCmd specifies
' how the application is to be shown when it is opened.
' This parameter can be one of the following values:

    Const SW_HIDE = 0                '  Hides the window and activates another window.
    Const SW_MAXIMIZE = 3            '  Maximizes the specified window.
    Const SW_MINIMIZE = 6            '  Minimizes the specified window and activates the next top-level window in the z-order.
    Const SW_RESTORE = 9             '  Activates and displays the window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    Const SW_SHOW = 5                '  Activates the window and displays it in its current size and position.
    Const SW_SHOWDEFAULT = 10        '  Sets the show state based on the SW_ flag specified in theSTARTUPINFO structure passed to theCreateProcess function by the program that started the application. An application should callShowWindow with this flag to set the initial show state of its main window.
    Const SW_SHOWMAXIMIZED = 3       '  Activates the window and displays it as a maximized window.
    Const SW_SHOWMINIMIZED = 2       '  Activates the window and displays it as a minimized window.
    Const SW_SHOWMINNOACTIVE = 7     '  Displays the window as a minimized window. The active window remains active.
    Const SW_SHOWNA = 8              '  Displays the window in its current state. The active window remains active.
    Const SW_SHOWNOACTIVATE = 4      '  Displays a window in its most recent size and position. The active window remains active.
    Const SW_SHOWNORMAL = 1          '  Activates and displays a window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when displaying the window for the first time.

' If lpFile specifies a document file, nShowCmd should be zero.
' You can use this function to open or explore a shell folder.
' To open a folder, use either
'
' ShellExecute(handle, NULL, path_to_folder, NULL, NULL, SW_SHOWNORMAL);
' or
' ShellExecute(handle, "open", path_to_folder, NULL, NULL, SW_SHOWNORMAL);
'
' To explore a folder, use the following call:
'
' ShellExecute(handle, "explore", path_to_folder, NULL, NULL, SW_SHOWNORMAL);
'
' If lpOperation is NULL, the function opens the file specified by lpFile. If lpOperation is "open" or "explore", the function will attempt to open or explore the folder.
'
' To obtain information about the application that is launched as a result of calling
'
' Returns a value greater than 32 if successful, or an error value that is less
' than or equal to 32 otherwise. The following table lists the error values.
' The return value is cast as an HINSTANCE for backward compatibility with 16-bit
' Microsoft® Windows® applications. It is not a true HINSTANCE, however.
' The only thing that can be done with the returned HINSTANCE is to cast it to an
' integer and compare it with the value 32 or one of the error codes below.

    Const SE_ERR_FNF = 2                 '  File not found
    Const SE_ERR_PNF = 3                 '  Path not found
    Const SE_ERR_ACCESSDENIED = 5        '  Access denied
    Const SE_ERR_OOM = 8                 '  Out of memory
    Const SE_ERR_DLLNOTFOUND = 32        '  DLL not found
    Const SE_ERR_SHARE = 26              '  A sharing violation occurred
    Const SE_ERR_ASSOCINCOMPLETE = 27    '  Incomplete or invalid file association
    Const SE_ERR_DDETIMEOUT = 28         '  DDE Time out
    Const SE_ERR_DDEFAIL = 29            '  DDE transaction failed
    Const SE_ERR_DDEBUSY = 30            '  DDE busy
    Const SE_ERR_NOASSOC = 31            '  No association for file extension
    Const ERROR_BAD_FORMAT = 11&         '  Invalid EXE file or error in EXE image
    Const ERROR_FILE_NOT_FOUND = 2&      '  The specified file was not found.
    Const ERROR_PATH_NOT_FOUND = 3&      '  The specified path was not found.
    Const ERROR_BAD_EXE_FORMAT = 193&    '  The .exe file is invalid (non-Win32® .exe or error in .exe image).


Public Function ShellExecLaunchFile(ByVal strPathFile As String, ByVal strOpenInPath As String, ByVal strArguments As String) As Long
On Error Resume Next
    Dim Scr_hDC As Long
    
    'Get the Desktop handle
    Scr_hDC = GetDesktopWindow()
    
    'Launch File
    ShellExecLaunchFile = ShellExecute(Scr_hDC, "", strPathFile, "", strOpenInPath, SW_SHOWNORMAL)

End Function


Public Function ShellExecLaunchErr(ByVal lngErrorNumber As Long, ByVal blnRaiseMsg As Boolean) As String
    On Error Resume Next
    Dim msg As VbMsgBoxResult
    Dim strErrorMessage As String
    
    If lngErrorNumber < 33 Then
        'There was an error
        Select Case lngErrorNumber
            Case SE_ERR_FNF
                strErrorMessage = "File not found."
            Case SE_ERR_PNF
                strErrorMessage = "Path not found."
            Case SE_ERR_ACCESSDENIED
                strErrorMessage = "Access denied."
            Case SE_ERR_OOM
                strErrorMessage = "Out of memory."
            Case SE_ERR_DLLNOTFOUND
                strErrorMessage = "DLL not found."
            Case SE_ERR_SHARE
                strErrorMessage = "A sharing violation occurred."
            Case SE_ERR_ASSOCINCOMPLETE
                strErrorMessage = "Incomplete or invalid file association."
            Case SE_ERR_DDETIMEOUT
                strErrorMessage = "DDE Time out."
            Case SE_ERR_DDEFAIL
                strErrorMessage = "DDE transaction failed."
            Case SE_ERR_DDEBUSY
                strErrorMessage = "DDE busy."
            Case SE_ERR_NOASSOC
                strErrorMessage = "No association for file extension."
            Case ERROR_BAD_FORMAT
                strErrorMessage = "Invalid EXE file or error in EXE image."
            Case ERROR_FILE_NOT_FOUND
                strErrorMessage = "The specified file was not found."
            Case ERROR_PATH_NOT_FOUND
                strErrorMessage = "The specified path was not found."
            Case ERROR_BAD_EXE_FORMAT
                strErrorMessage = "The .exe file is invalid (non-Win32® .exe or error in .exe image)."
            Case Else
                strErrorMessage = "Unknown error."
        End Select
        
        'If the blnRaiseMsg = True then raise a MsgBox with error
        If blnRaiseMsg = True Then msg = MsgBox(strErrorMessage, vbCritical, "File Error")
        
        'Return Error string
        ShellExecLaunchErr = blnRaiseMsg
    
    End If
    
End Function


' So the way to use all this is:
'
'    Dim lngReturnNumber As Long
'
'    lngReturnNumber = ShellExecLaunchFile(txtPathFile.Text, txtStartPath.Text, txtArguments.Text)
'    If lngReturnNumber < 33 Then
'        Call ShellExecLaunchErr(lngReturnNumber, True)
'        Exit Sub
'    End If
'
'==================================================================================
'
'
'Use the following runRegEntry Function to Silently Run .reg Files
'
Public Function runRegEntry(strPathFile As String) As Boolean
 On Error Resume Next
    
    Dim dblTemp As Double
    dblTemp = Shell("regedit.exe /s " & strPathFile, vbHide)
    
    runRegEntry = True
    
    Exit Function
    
Command1Err:
    Dim msg As VbMsgBoxResult

    msg = MsgBox("Error # " & CStr(Err.Number) & " " & Err.Description & vbNewLine & "With: " & strPathFile, vbCritical, "Error:")
    Err.Clear    ' Clear the error.
    runRegEntry = False
    
End Function
