Attribute VB_Name = "FileExists"
Option Explicit
Public Function FileExist(path$) As Integer
    Dim x
    x = FreeFile
    On Error Resume Next
    Open path$ For Input As x
    FileExist = IIf(Err = 0, True, False)
    Close x
    Err = 0
End Function

