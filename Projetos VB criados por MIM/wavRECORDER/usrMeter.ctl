VERSION 5.00
Begin VB.UserControl usrMeter 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   FillColor       =   &H000000FF&
   ForeColor       =   &H00A4E3AC&
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   151
End
Attribute VB_Name = "usrMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    Reserved As Long
End Type

Private Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Private Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Private Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Private Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Private Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Private Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit

Private Const WAVE_FORMAT_PCM = 1

Private Const WHDR_DONE = &H1&              '/* done bit */
Private Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Private Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Private Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Private Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Private Const WIM_OPEN = &H3BE
Private Const WIM_CLOSE = &H3BF
Private Const WIM_DATA = &H3C0

Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private DevHandle As Long
Private InData(0 To 511) As Byte

'Retorna lista de Devices de som de máquina
Public Function GetDevices() As String
    Dim Caps As WaveInCaps
    Dim DevicePosition As Long
          
    'lista todos os devices encontrados
    For DevicePosition = 0 To waveInGetNumDevs - 1
        'pega propriedades de device de posicao DevicePosition
        Call waveInGetDevCaps(DevicePosition, VarPtr(Caps), Len(Caps))
        'Formato 11.025 kHz, Stereo, 8-bit de PLACA
        If Caps.Formats And WAVE_FORMAT_1S08 Then
            GetDevices = GetDevices & StrConv(Caps.ProductName, vbUnicode) & Chr$(1) & DevicePosition & vbCrLf
        End If
    Next
    
End Function

'Para a analise
Public Sub StopMetering()
    waveInReset DevHandle
    waveInClose DevHandle
    DevHandle = 0
End Sub

Public Function StartMetering(ByVal DevicePosicion As Long) As Boolean
Static WaveFormat As WaveFormatEx
    
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 2 'Two channels -- left and right
        .SamplesPerSec = 11025 '11khz
        .BitsPerSample = 8
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    Debug.Print "waveInOpen:"; waveInOpen(DevHandle, DevicePosicion, VarPtr(WaveFormat), 0, 0, 0)
    
    If DevHandle = 0 Then
        StartMetering = False
        Exit Function
    End If
    
    Debug.Print " "; DevHandle
    
    waveInStart DevHandle

    Visualize

    StartMetering = True

End Function

'Manda analisar dados de entrada para visualizacao
Public Sub Visualize()
Static Wave As WaveHdr
Static EmAcao As Boolean
    
    Wave.lpData = VarPtr(InData(0))
    Wave.dwBufferLength = 512 'This is now 512 so there's still 256 samples per channel
    Wave.dwFlags = 0
    
    'EmAcao = True
    'Do
    
        waveInPrepareHeader DevHandle, VarPtr(Wave), Len(Wave)
        waveInAddBuffer DevHandle, VarPtr(Wave), Len(Wave)
        Do
            'Aguarda até driver marcar como onda pronta
'            DoEvents
        Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
        
        waveInUnprepareHeader DevHandle, VarPtr(Wave), Len(Wave)
        
        'If DevHandle = 0 Then
            'se devide estiver fechado sai
            'Exit Do
        'End If
        
        DrawData
        
    '    DoEvents
        
    'Loop While DevHandle <> 0 'enquanto houver device aberto
    'EmAcao = False

End Sub

'Desenha Osciloscopio
Private Sub DrawData()
Static X As Long
    
    Cls
    
    CurrentX = -1
    CurrentY = ScaleHeight \ 2
    
    'Plot the data...
    For X = 0 To 255
        Line Step(0, 0)-(X, InData(X * 2))
        'Line Step(0, 0)-(X, InData(X * 2 + 1)) 'For a good soundcard...
        
        'Use these to plot dots instead of lines...
        'Scope(0).PSet (X, InData(X * 2))
        'Scope(1).PSet (X, InData(X * 2 + 1)) 'For a good soundcard...

        'My soundcard is pretty cheap... the right is
        'noticably less loud than the left... so I add five to it.
        'Scope(1).Line Step(0, 0)-(X, InData(X * 2 + 1) + 5)
    Next
            
    CurrentY = Width

End Sub

'lista canais
Public Function DeviceChannels() As Variant
    DeviceChannels = InData
End Function

'se form fechado sem finalizar - finaliza
Private Sub UserControl_Hide()
    StopMetering
End Sub

'acerta escala
Private Sub UserControl_Initialize()
    ScaleMode = vbUser
End Sub

'Assim que mover, limpa
Private Sub UserControl_Resize()

    Cls
    ScaleHeight = 256
    ScaleWidth = 123

End Sub

'fecha device
Private Sub UserControl_Terminate()
    StopMetering
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    FillColor = UserControl.ForeColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.ForeColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H404040)
    UserControl.ForeColor = PropBag.ReadProperty("FillColor", &HA4E3AC)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H404040)
    Call PropBag.WriteProperty("FillColor", UserControl.ForeColor, &HA4E3AC)
End Sub

