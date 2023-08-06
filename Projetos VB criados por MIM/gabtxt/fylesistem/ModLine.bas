Attribute VB_Name = "ModLine"
 Option Explicit

Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Any) As Long

Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETLINE = &HC4
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINEINDEX = &HBB

Public Function LineText(TextBox As TextBox, LineNumber As Long) _
       As String

    Dim nRet As Long
    
    nRet = SendMessage(TextBox.hwnd, EM_LINEINDEX, LineNumber - 1, _
           ByVal 0)
           
    LineText = Space(SendMessage(TextBox.hwnd, EM_LINELENGTH, _
               nRet, ByVal 0) + 2)
               
    Mid(LineText, 1, 1) = Chr(Len(LineText) Mod 256)
    Mid(LineText, 2, 1) = Chr(Len(LineText) \ 256)
    
    nRet = SendMessage(TextBox.hwnd, EM_GETLINE, LineNumber - 1, _
           ByVal LineText)
           
    Debug.Print Len(LineText), nRet
           
    LineText = Mid(LineText, 1, nRet)

End Function

Public Function CurrentLine(TextBox As TextBox) As Long

    CurrentLine = SendMessage(TextBox.hwnd, EM_LINEFROMCHAR, -1, _
                  ByVal 0) + 1

End Function

Public Function LineCount(TextBox As TextBox) As Long

    LineCount = SendMessage(TextBox.hwnd, EM_GETLINECOUNT, 0, ByVal 0)

End Function

