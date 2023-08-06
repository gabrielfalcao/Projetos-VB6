Attribute VB_Name = "modBuscaDePalavrasPelaInicial"
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Public Const LB_FINDSTRING = &H18F

Public strPathEdit As String
