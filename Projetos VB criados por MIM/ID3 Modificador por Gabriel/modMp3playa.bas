Attribute VB_Name = "modMp3playa"
Public Const AliasName = "audiomp3"
Global Path As String
Public strFilePath As String
Public strCommand As String

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209

Sub Main()
  If App.PrevInstance Then
    Dim SaveTitle As String
    If Command$ <> "" Then SaveSetting App.Title, "Config", "Command", Command$
    SaveTitle = App.Title
    App.Title = ""
    AppActivate SaveTitle
    End
  End If

  Path = App.Path
  If Right$(Path, 1) <> "\" Then Path = Path & "\"
  
  INI = Path & "MPlayer3.ini"
  If Not FileExists(INI) Then
    Open INI For Output As #1
      Print #1, ""
    Close #1
  End If
      
  strMainX = ReadINI("MPlayer3", "MainX", "2000")
  strMainY = ReadINI("MPlayer3", "MainY", "1500")
  strRateX = ReadINI("MPlayer3", "RateX", "5000")
  strRateY = ReadINI("MPlayer3", "RateY", "3000")
  bSPLS = CBool(ReadINI("MPlayer3", "SPLS", "True"))
  bSRS = CBool(ReadINI("MPlayer3", "SRS", "False"))
  bAlwaysTray = CBool(ReadINI("MPlayer3", "AlwaysTray", "False"))
  bMinTray = CBool(ReadINI("MPlayer3", "MinTray", "False"))
  strLOpenPath = ReadINI("MPlayer3", "LOpenPath", App.Path)
  strLSavePath = ReadINI("MPlayer3", "LSavePath", App.Path)
  intLIndex = CInt(ReadINI("MPlayer3", "LIndex", "0"))
      
  Load frmMain
End Sub

Public Sub ReadMP3Header(sPassFileName As String)
Dim z, i
Dim BinaryString As String
Dim byteArray(4) As Byte    'array that store first four bytes
Dim bin As String           'string that store binary number converted from readed bytes
Dim BinString As String     'containing binary string
Dim DecString As Integer  'containing decimal extracted from BinString
'''''''''''''''end of declarations'''''''

Open sPassFileName For Binary Access Read As #1  'open file #1 for read
   For z = 1 To 4                           'step through four bytes
   Get #1, z, byteArray(z)                  'store every(z)byte  in array position z
   Next z                                   'back for next byte
 Close #1                                   'close file
 bin = ""                                   'reset and build the desired binary number in this string
   For z = 1 To 4                           'convert all bytes to binary
     For i = 0 To 7 Step 1                  'Here comes the decimal=>binary conversion
         If byteArray(z) And (2 ^ i) Then   'Use the logical "AND" operator.
            bin = bin + "1"
            Else
            bin = bin + "0"
         End If
         Next i                             'End of binary conversion
Next z
BinaryString = bin
'''''''''check MP3HeaderInfo.Frequency''''
DecString = 0
BinString = Mid(bin, 19, 2)         'take 19 to 21
For i = 1 To Len(BinString)         'convert to decimal
  If Mid(BinString, i, 1) = 1 Then
    DecString = DecString + 2 ^ (Len(BinString) - i)
  End If
Next i
Select Case DecString
  Case 0
    frmMain.lblKhz = 44
  Case 1
    frmMain.lblKhz = 32
  Case 2
    frmMain.lblKhz = 48
  Case 3
End Select
''''check MP3HeaderInfo.Mode''''
DecString = 0
BinString = Mid(bin, 31, 2)
For i = 1 To Len(BinString)
  If Mid(BinString, i, 1) = 1 Then
    DecString = DecString + 2 ^ (Len(BinString) - i)
  End If
Next i
Select Case DecString
  Case 0
    frmMain.lblmode = "stereo"
  Case 1
    frmMain.lblmode = "stereo"
  Case 2
    frmMain.lblmode = "stereo"
  Case 3
    frmMain.lblmode = "mono"
End Select
'''''check MP3HeaderInfo.Bitrate''''
DecString = 0
BinString = Mid(bin, 21, 4)
For i = 1 To Len(BinString)
  If Mid(BinString, i, 1) = 1 Then
    DecString = DecString + 2 ^ (Len(BinString) - i)
  End If
Next i
Select Case DecString
  Case 0
    frmMain.lblbitrate = 0
  Case 1
    frmMain.lblbitrate = 112
  Case 2
    frmMain.lblbitrate = 56
  Case 3
    frmMain.lblbitrate = 224
  Case 4
    frmMain.lblbitrate = 40
  Case 5
    frmMain.lblbitrate = 160
  Case 6
    frmMain.lblbitrate = 80
  Case 7
    frmMain.lblbitrate = 320
  Case 8
    frmMain.lblbitrate = 32
  Case 9
    frmMain.lblbitrate = 128
  Case 10
    frmMain.lblbitrate = 64
  Case 11
    frmMain.lblbitrate = 256
  Case 12
    frmMain.lblbitrate = 48
  Case 13
    frmMain.lblbitrate = 192
  Case 14
    frmMain.lblbitrate = 96
  Case 15
    frmMain.lblbitrate = 0
End Select
End Sub

