VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAVI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class       : CAVI
' Description : This class plays AVI files using the MCI (media control
'               interface) API
' Source      : Total VB SourceBook 6
'

Private Type AVIRECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function GetClientRect _
  Lib "user32" _
  (ByVal hWnd As Long, _
    lpRect As AVIRECT) _
  As Long
  
Private Declare Function mciSendCommand _
  Lib "winmm.dll" _
  Alias "mciSendCommandA" _
  (ByVal wDeviceID As Long, _
    ByVal uMessage As Long, _
    ByVal dwParam1 As Long, _
    ByRef dwParam2 As Any) _
  As Long
  
Private Declare Function mciGetErrorString _
  Lib "winmm.dll" _
  Alias "mciGetErrorStringA" _
  (ByVal dwError As Long, _
    ByVal lpstrBuffer As String, _
    ByVal uLength As Long) _
  As Long
  
Private Declare Function MoveWindow _
  Lib "user32" _
  (ByVal hWnd As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal bRepaint As Long) _
  As Long

Private Type MCI_DGV_RECT_PARMS
  dwCallback As Long
  rc As AVIRECT
End Type

Private Type MCI_DGV_OPEN_PARMS
  dwCallback As Long
  wDeviceID As Long
  lpstrDeviceType As String
  lpstrElementName As String
  lpstrAlias As String
  dwStyle As Long
  hWndParent As Long
End Type

Private Type MCI_DGV_WINDOW_PARMS
  dwCallback As Long
  hWnd As Long
  nCmdShow As Long
  lpstrText As String
    
End Type

Private Type MCI_DGV_STATUS_PARMS
  dwCallback As Long
  dwReturn As Long
  dwItem As Long
  dwTrack As Long
  lpstrDrive As String
  dwReference As Long
End Type

Private Type MCI_OPEN_PARMS
  dwCallback As Long
  wDeviceID As Long
  lpstrDeviceType As String
  lpstrElementName As String
  lpstrAlias As String
End Type

Private Type MCI_GENERIC_PARMS
  dwCallback As Long
End Type

Private Type MCI_SET_PARMS
  dwCallback As Long
  dwTimeFormat As Long
  dwAudio As Long
End Type

Private Type MCI_PLAY_PARMS
  dwCallback As Long
  dwFrom As Long
  dwTo As Long
End Type

Private Type MCI_STATUS_PARMS
  dwCallback As Long
  dwReturn As Long
  dwItem As Long
  dwTrack As Integer
End Type

Private Type MCI_SEEK_PARMS
  dwCallback As Long
  dwTo As Long
End Type

Private Type MCI_RECORD_PARMS
  dwCallback As Long
  dwFrom As Long
  dwTo As Long
End Type

Private Type MCI_SAVE_PARMS
  dwCallback As Long
  lpFileName As String
End Type

Private Const WS_CHILD = &H40000000
Private Const SW_SHOW = 5

Private Const MCI_CLOSE = &H804&
Private Const MCI_WAIT = &H2&
Private Const MCI_OPEN = &H803&
Private Const MCI_FORMAT_MILLISECONDS = 0
Private Const MCI_SET = &H80D&
Private Const MCI_OPEN_ELEMENT = &H200&
Private Const MCI_SET_TIME_FORMAT = &H400&
Private Const MCI_STOP = &H808&
Private Const MCI_SEEK = &H807&
Private Const MCI_SEEK_TO_START = &H100&
Private Const MCI_PLAY = &H806&
Private Const MCI_NOTIFY = &H1&
Private Const MCI_STATUS_POSITION = &H2&
Private Const MCI_STATUS = &H814&
Private Const MCI_STATUS_ITEM = &H100&
Private Const MCI_STATUS_LENGTH = &H1&
Private Const MCI_TO = &H8&
Private Const MCI_OPEN_TYPE = &H2000&
Private Const MCI_RECORD = &H80F&
Private Const MCI_SAVE = &H813&
Private Const MCI_SAVE_FILE = &H100&
Private Const MCI_MCIAVI_PLAY_WINDOW = &H1000000
Private Const MCI_PAUSE = &H809&
Private Const MCI_DGV_OPEN_PARENT = &H20000
Private Const MCI_DGV_OPEN_WS = &H10000
Private Const MCI_WINDOW = &H841&
Private Const MCI_DGV_WINDOW_STATE = &H40000
Private Const MCI_DGV_STATUS_HWND = &H4001&
Private Const MCI_WHERE = &H843
Private Const MCI_DGV_WHERE_SOURCE = &H20000

Private m_strFileName As String
Private m_lnghWndParent As Long
Private m_fStretch As Boolean

'The device id of the mci device
Private mlngDeviceID As Long
Private m_lnghWnd As Long

'position
Private m_lngLeft As Long
Private m_lngTop As Long
Private m_lngWidth As Long
Private m_lngHeight As Long

Private Sub Class_Terminate()
  ' Make sure the AVI file is closed
  ' Source: Total VB SourceBook 6

  CloseAVI

End Sub

Public Property Get FileName() As String
  ' Returns: the filename of the AVI file to play
  ' Source: Total VB SourceBook 6

  FileName = m_strFileName

End Property

Public Property Let FileName(ByVal strValue As String)
  ' strValue: Set the filename of the AVI file to play
  ' Source: Total VB SourceBook 6

  m_strFileName = strValue

End Property

Public Property Get Height() As Long
  ' Returns: The Height of the playback window
  ' Source: Total VB SourceBook 6

  Height = m_lngHeight

End Property

Public Property Let Height(ByVal lngValue As Long)
  ' lngValue: Set the Height of the playback window. This is measured in pixels
  ' Source: Total VB SourceBook 6

  m_lngHeight = lngValue

End Property

Public Property Get hWnd() As Long
  ' Returns: The handle of the AVI window
  ' Source: Total VB SourceBook 6

  hWnd = m_lnghWnd

End Property

Public Property Get hWndParent() As Long
  ' Returns: The handle of the parent window
  ' Source: Total VB SourceBook 6

  hWndParent = m_lnghWndParent

End Property

Public Property Let hWndParent(ByVal lngValue As Long)
  ' lngValue: Set the handle of the parent window
  ' Source: Total VB SourceBook 6

  m_lnghWndParent = lngValue

End Property

Public Property Get Left() As Long
  ' Returns: The left position of the playback window
  ' Source: Total VB SourceBook 6

  Left = m_lngLeft

End Property

Public Property Let Left(ByVal lngValue As Long)
  ' lngValue: Set the left position of the playback window. This is measured in
  ' pixels, and is relative to the parent window
  ' Source: Total VB SourceBook 6

  m_lngLeft = lngValue

End Property

Public Property Get Length() As Long
  ' Comments  : Returns the length of the AVI movie
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR
  
  Dim mciStatusParms As MCI_STATUS_PARMS

  ' indicate that we want to retrieve the length
  mciStatusParms.dwItem = MCI_STATUS_LENGTH
  ' get the length
  mciSendCommand mlngDeviceID, MCI_STATUS, MCI_STATUS_ITEM, mciStatusParms

  ' return the length
  Length = mciStatusParms.dwReturn
  
PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Length"
  Resume PROC_EXIT
  
End Property

Public Property Get Position() As Long
  ' Returns: The current playback position
  ' Source: Total VB SourceBook 6

  Dim mciStatusParms As MCI_STATUS_PARMS
  Dim lngResult As Long

  On Error GoTo PROC_ERR
  
  mciStatusParms.dwItem = MCI_STATUS_POSITION
  ' get the position
  lngResult = mciSendCommand(mlngDeviceID, MCI_STATUS, MCI_STATUS_ITEM, _
    mciStatusParms)
  ' If there wasn't an error
  If lngResult = 0 Then
    ' If the return value is invalid, return 0
    If mciStatusParms.dwReturn < 0 Or mciStatusParms.dwReturn > Length Then
      Position = 0
    Else
      ' return the position
      Position = mciStatusParms.dwReturn
    End If
  Else
    ' If there was an error, return 0
    Position = 0
  End If

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Position"
  Resume PROC_EXIT

End Property

Public Property Let Position(ByVal lngValue As Long)
  ' lngValue: Set the current position of playback
  ' Source: Total VB SourceBook 6
  
  Dim mciSeekParms As MCI_SEEK_PARMS
  
  On Error GoTo PROC_ERR

  mciSeekParms.dwTo = lngValue

  mciSendCommand mlngDeviceID, MCI_SEEK, MCI_TO, mciSeekParms

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Position"
  Resume PROC_EXIT

End Property

Public Property Get Stretch() As Boolean
  ' Returns: The stretch flag
  ' Source: Total VB SourceBook 6

  Stretch = m_fStretch

End Property

Public Property Let Stretch(ByVal fValue As Boolean)
  ' fValue: Set the stretch flag
  ' Source: Total VB SourceBook 6

  m_fStretch = fValue
  PositionMovieWindow

End Property

Public Property Get Top() As Long
  ' Returns: The top position of the playback window
  ' Source: Total VB SourceBook 6

  Top = m_lngTop

End Property

Public Property Let Top(ByVal lngValue As Long)
  ' lngValue: Set the Top position of the playback window. This is measured in
  ' pixels, and is relative to the parent window
  ' Source: Total VB SourceBook 6

  m_lngTop = lngValue

End Property

Public Property Get Width() As Long
  ' Returns: The Width of the playback window
  ' Source: Total VB SourceBook 6

  Width = m_lngWidth

End Property

Public Property Let Width(ByVal lngValue As Long)
  ' lngValue: Set the Width of the playback window. This is measured in pixels
  ' Source: Total VB SourceBook 6

  m_lngWidth = lngValue

End Property

Public Sub CloseAVI()
  ' Comments  : Close the AVI file
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim mciGenericParms As MCI_GENERIC_PARMS
  
  On Error GoTo PROC_ERR
  
  ' if a movie is open then close it
  If mlngDeviceID <> 0 Then
    ' tell the device to close
    mciSendCommand mlngDeviceID, MCI_CLOSE, 0, mciGenericParms
    mlngDeviceID = 0
  End If

  m_lnghWnd = 0

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "CloseAVI"
  Resume PROC_EXIT

End Sub

Public Sub OpenAVI()
  ' Comments  : Open the AVI file
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim mciDGVOpenParms As MCI_DGV_OPEN_PARMS
  Dim mciDGVWindowParms As MCI_DGV_WINDOW_PARMS
  Dim mciDGVStatusParms As MCI_DGV_STATUS_PARMS
  Dim lngResult As Long
  Dim strErrorDescription As String
  
  On Error GoTo PROC_ERR
  
  ' close any old movies before opening the new one.
  CloseAVI
  
  ' set up the open parameters
  mciDGVOpenParms.dwCallback = 0
  mciDGVOpenParms.wDeviceID = 0
  mciDGVOpenParms.lpstrDeviceType = 0
  mciDGVOpenParms.lpstrElementName = m_strFileName
  mciDGVOpenParms.lpstrAlias = vbNullString
  mciDGVOpenParms.dwStyle = WS_CHILD
  mciDGVOpenParms.hWndParent = m_lnghWndParent
  
  ' try to open the file
  lngResult = mciSendCommand(0, MCI_OPEN, MCI_OPEN_ELEMENT Or _
    MCI_DGV_OPEN_PARENT Or MCI_DGV_OPEN_WS, mciDGVOpenParms)
  If lngResult = 0 Then
    ' we opened the file o.k., now set up to
    ' play it.
    
    ' save ID
    mlngDeviceID = mciDGVOpenParms.wDeviceID
    
    ' show the playback window
    mciDGVWindowParms.dwCallback = 0
    mciDGVWindowParms.hWnd = 0
    mciDGVWindowParms.nCmdShow = SW_SHOW
    mciDGVWindowParms.lpstrText = vbNullString
    mciSendCommand mlngDeviceID, MCI_WINDOW, MCI_DGV_WINDOW_STATE, _
      mciDGVWindowParms
    
    ' get the window handle
    mciDGVStatusParms.dwItem = MCI_DGV_STATUS_HWND
    mciSendCommand mlngDeviceID, MCI_STATUS, MCI_STATUS_ITEM, _
      mciDGVStatusParms
    m_lnghWnd = mciDGVStatusParms.dwReturn
    
    ' now position the movie window
    PositionMovieWindow
  
  Else
    strErrorDescription = Space$(100)
    mciGetErrorString lngResult, strErrorDescription, 100
    Err.Raise lngResult, , strErrorDescription
  End If
  
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "OpenAVI"
  Resume PROC_EXIT

End Sub

Public Sub Pause()
  ' Comments  : Pause the playback
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim mciGenericParms As MCI_GENERIC_PARMS

  On Error GoTo PROC_ERR

  ' tell MCI to pause
  mciSendCommand mlngDeviceID, MCI_PAUSE, 0, mciGenericParms

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Pause"
  Resume PROC_EXIT

End Sub

Public Sub Play()
  ' Comments  : Play the AVI file
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim mciPlayParms As MCI_PLAY_PARMS
    
  On Error GoTo PROC_ERR
  
  ' tell MCI to start playing
  mciSendCommand mlngDeviceID, MCI_PLAY, MCI_MCIAVI_PLAY_WINDOW, mciPlayParms

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Play"
  Resume PROC_EXIT

End Sub

Public Sub Rewind()
  ' Comments  : Rewinds the movie
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  
  On Error GoTo PROC_ERR
    
  ' tell MCI to rewind the file
  lngResult = mciSendCommand(mlngDeviceID, MCI_SEEK, MCI_WAIT Or _
    MCI_SEEK_TO_START, 0)
  
  ' set position to 0
  Position = 0

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Rewind"
  Resume PROC_EXIT

End Sub

Public Sub StopPlay()
  ' Comments  : Stop the AVI play
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  ' tell MCI to pause
  mciSendCommand mlngDeviceID, MCI_STOP, 0, 0
  
  ' stop also needs to reset the play to the beginning
  Rewind

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "StopPlay"
  Resume PROC_EXIT

End Sub

Private Sub PositionMovieWindow()
  ' Comments  : Position the movie playback window. If stretch is True then
  '             make playback window the size indicated in the width and
  '             height properties. Otherwise size it to the size of the movie
  '             starting at the left and top properties All measurements are
  '             in pixels
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim rcMovie As AVIRECT
  Dim mciDGVRectParms As MCI_DGV_RECT_PARMS
  
  On Error GoTo PROC_ERR
  
  If m_fStretch Then
  
    ' reposition the playback window
    MoveWindow m_lnghWnd, m_lngLeft, m_lngTop, m_lngWidth, m_lngHeight, 1
  
  Else
  
    ' get the original size of the movie
    mciSendCommand mlngDeviceID, MCI_WHERE, MCI_DGV_WHERE_SOURCE, _
      mciDGVRectParms
    
    rcMovie = mciDGVRectParms.rc
    rcMovie.Left = m_lngLeft
    rcMovie.Top = m_lngTop
    
    ' reposition the playback window
    MoveWindow m_lnghWnd, rcMovie.Left, rcMovie.Top, rcMovie.Right, _
      rcMovie.Bottom, 1
  End If

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "PositionMovieWindow"
  Resume PROC_EXIT

End Sub


