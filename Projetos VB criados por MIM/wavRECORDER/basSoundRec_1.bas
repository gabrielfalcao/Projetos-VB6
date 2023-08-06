Attribute VB_Name = "basSoundRec_1"
Option Explicit

Public Const CALLBACK_FUNCTION = &H30000
Public Const CALLBACK_WINDOW = &H10000      '  dwCallback is a HWND
Public Const MM_WIM_DATA = &H3C0
Public Const WHDR_DONE = &H1         '  done bit
Public Const WIM_DATA = MM_WIM_DATA
Public Const GMEM_FIXED = &H0         ' Global Memory Flag used by GlobalAlloc functin
Public Const NUM_BUFFERS = 10
Public BUFFER_SIZE As Long  '= 8192
Public Const DEVICEID = -1
Public Const GWL_WNDPROC = -4

Type WAVEHDR
   lpData As Long          ' Address of the waveform buffer.
   dwBufferLength As Long  ' Length, in bytes, of the buffer.
   dwBytesRecorded As Long ' When the header is used in input, this member specifies how much
                           ' data is in the buffer.

   dwUser As Long          ' User data.
   dwFlags As Long         ' Flags supplying information about the buffer. Set equal to zero.
   dwLoops As Long         ' Number of times to play the loop. Set equal to zero.
   lpNext As Long          ' Not used
   reserved As Long        ' Not used
End Type

Type WAVEFORMAT
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long


Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Sub CopyStringFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal a As String, p As Any, ByVal cb As Long)
Public Declare Sub CopyStructFromPtr Lib "kernel32" _
                     Alias "RtlMoveMemory" _
                     (struct As Any, _
                     ByVal ptr As Long, _
                     ByVal cb As Long)

Public Declare Sub CopyPtrFromStruct Lib "kernel32" _
                     Alias "RtlMoveMemory" _
                     (ByVal ptr As Long, _
                     struct As Any, _
                     ByVal cb As Long)
                     
Public Declare Function CallWindowProc Lib "user32" Alias _
"CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal msg As Long, _
    ByVal wParam As Long, ByRef lParam As WAVEHDR) As Long

Public Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public i As Integer
Public j As Integer
Public rc As Long

Public msg As String * 200

Public hWaveIn As Long
Public wformat As WAVEFORMAT
Public hmem(NUM_BUFFERS) As Long
Public inHdr(NUM_BUFFERS) As WAVEHDR
Public fRecording As Boolean
Dim lpPrevWndProc As Long

Dim hwnd As Long            ' window handle

Function StartInput() As Boolean
Dim lBuffSize As Long

    If fRecording Then
        StartInput = True
        Exit Function
    End If
    BUFFER_SIZE = (wformat.nSamplesPerSec * wformat.nBlockAlign * wformat.nChannels * 0.1) - ((wformat.nSamplesPerSec * wformat.nBlockAlign * wformat.nChannels * 0.1) Mod (wformat.nBlockAlign))
    For i = 0 To NUM_BUFFERS - 1
        hmem(i) = GlobalAlloc(&H40, BUFFER_SIZE)               'BUFFER_SIZE
        inHdr(i).lpData = GlobalLock(hmem(i))
        inHdr(i).dwBufferLength = BUFFER_SIZE
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
    Next

    rc = waveInOpen(hWaveIn, DEVICEID, wformat, hwnd, 0, CALLBACK_WINDOW)
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg
        StartInput = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    For i = 0 To NUM_BUFFERS - 1
        addData inHdr(i)
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
End Function
Sub addData(iHdr As WAVEHDR)
Dim iRet As Long
Dim sBuff  As String
    
    rc = waveInAddBuffer(hWaveIn, iHdr, Len(iHdr))
    
    sBuff = Space(BUFFER_SIZE)
    CopyMemory ByVal sBuff, ByVal iHdr.lpData, BUFFER_SIZE
    mmioWrite hmmioIn, sBuff, BUFFER_SIZE

End Sub
' Stop receiving audio input on the soundcard
Sub StopInput()
Dim iRet As Long
Dim icount As Long
    
    fRecording = False
    iRet = waveInReset(hWaveIn)
    iRet = waveInStop(hWaveIn)
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    
    iRet = waveInClose(hWaveIn)
    
    'Ascent out of Data Chunk
    If (mmioAscend(hmmioIn, mmckinfoSubchunkIn, 0) <> 0) Then
        MsgBox "Cannot ascend out of DATA CHUNK"
        mmioClose hmmioIn, 0
    End If
    
    If (mmioAscend(hmmioIn, mmckinfoSubchunkIn, 0) <> 0) Then
        MsgBox "Could Ascend Out Of wformat Chunk"
        mmioClose hmmioIn, 0
        Exit Sub
    End If

    'Asecnd out of the RIFF chunk
    If (mmioAscend(hmmioIn, mmckinfoParentIn, 0) <> 0) Then
        MsgBox "Cannot ascend out of RIFF CHUNK"
        mmioClose hmmioIn, 0
    End If
    
    mmioClose hmmioIn, 0
    
End Sub

Public Sub Initialize(hwndIn As Long)
    hwnd = hwndIn
    lpPrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByRef wavhdr As WAVEHDR) As Long

 If uMsg = WIM_DATA Then
    frmSoundRec.Gandu
 End If
 WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, wavhdr)
End Function

