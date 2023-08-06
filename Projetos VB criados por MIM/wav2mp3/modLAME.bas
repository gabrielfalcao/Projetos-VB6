Attribute VB_Name = "modLAME"
' Let's see how well I understand C++ header files...
' Hope I got em right...

Option Explicit

'/* encoding formats */
Public Const BE_CONFIG_MP3 = 0
Public Const BE_CONFIG_LAME = 256


'/* error codes */
Public Const BE_ERR_SUCCESSFUL = &H0
Public Const BE_ERR_INVALID_FORMAT = &H1
Public Const BE_ERR_INVALID_FORMAT_PARAMETERS = &H2
Public Const BE_ERR_NO_MORE_HANDLES = &H3
Public Const BE_ERR_INVALID_HANDLE = &H4
Public Const BE_ERR_BUFFER_TOO_SMALL = &H5

'/* other constants */
Public Const BE_MAX_HOMEPAGE = 128

'/* format specific variables */
Public Const BE_MP3_MODE_STEREO = 0
Public Const BE_MP3_MODE_JSTEREO = 1
Public Const BE_MP3_MODE_DUALCHANNEL = 2
Public Const BE_MP3_MODE_MONO = 3

Public Const MPEG1 = 1
Public Const MPEG2 = 0


Public Const CURRENT_STRUCT_VERSION = 1
Public Const CURRENT_STRUCT_SIZE = 331 '// is currently 331 bytes

'Public Enum LAME_QUALTIY_PRESET
Public Const LQP_NOPRESET = -1

  '// STANDARD QUALITY PRESETS
Public Const LQP_NORMAL_QUALITY = 0
Public Const LQP_LOW_QUALITY = 1
Public Const LQP_HIGH_QUALITY = 2
Public Const LQP_VOICE_QUALITY = 3

  '// NEW PRESET VALUES
Public Const LQP_PHONE = 1000
Public Const LQP_SW = 2000
Public Const LQP_AM = 3000
Public Const LQP_FM = 4000
Public Const LQP_VOICE = 5000
Public Const LQP_RADIO = 6000
Public Const LQP_TAPE = 7000
Public Const LQP_HIFI = 8000
Public Const LQP_CD = 9000
Public Const LQP_STUDIO = 10000
'End Enum

Public Type PBE_MP3     ' 23 bytes
  dwSampleRate As Long  '// 48000, 44100 and 32000 allowed
  byMode As Byte        '// BE_MP3_MODE_STEREO, BE_MP3_MODE_DUALCHANNEL, BE_MP3_MODE_MONO
  wBitrate As Integer   '// 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256 and 320 allowed
  bPrivate As Long
  bCRC  As Long
  bCopyright As Long
  bOriginal As Long
End Type

Public Type PBE_AAC ' 8 bytes
  dwSampleRate As Long
  byMode As Byte
  wBitrate As Integer
  byEncodingMethod As Byte
End Type

Public Type PBE_LHV1 ' 327 bytes
  '// STRUCTURE INFORMATION
  dwStructVersion As Long
  dwStructSize As Long

  '// BASIC ENCODER SETTINGS
  dwSampleRate As Long    '// SAMPLERATE OF INPUT FILE
  dwReSampleRate As Long  '// DOWNSAMPLERATE, 0=ENCODER DECIDES
  nMode As Long           '// BE_MP3_MODE_STEREO, BE_MP3_MODE_DUALCHANNEL, BE_MP3_MODE_MONO
  dwBitrate As Long       '// CBR bitrate, VBR min bitrate
  dwMaxBitrate As Long    '// CBR ignored, VBR Max bitrate
  nPreset As Long         '// Quality preset, use one of the settings of the LAME_QUALITY_PRESET enum
  dwMpegVersion As Long   '// FUTURE USE, MPEG-1 OR MPEG-2
  dwPsyModel As Long      '// FUTURE USE, SET TO 0
  dwEmphasis As Long      '// FUTURE USE, SET TO 0

  '// BIT STREAM SETTINGS
  bPrivate As Long    '// Set Private Bit (TRUE/FALSE)
  bCRC As Long        '// Insert CRC (TRUE/FALSE)
  bCopyright As Long  '// Set Copyright Bit (TRUE/FALSE)
  bOriginal As Long   '// Set Original Bit (TRUE/FALSE)
  
  '// VBR STUFF
  bWriteVBRHeader As Long '// WRITE XING VBR HEADER (TRUE/FALSE)
  bEnableVBR As Long      '// USE VBR ENCODING (TRUE/FALSE)
  nVBRQuality As Long     '// VBR QUALITY 0..9
  dwVbrAbr_bps As Long    '// Use ABR in stead of nVBRQuality
  bNoRes As Long          '// Disable Bit resorvoir

  btReserved(1 To 255 - 8) As Byte '// FUTURE USE, SET TO 0 (btReserved(255-2*sizeof(DWORD)))
End Type

Public Type PBE_FORMAT  ' This one was interesting... In C header file
  'MP3 As PBE_MP3       ' it looked like all these were in same type
  LHV1 As PBE_LHV1      ' declaration but when I tried... They weren't...
  'AAC As PBE_AAC       ' But then again... Editing LAME was first time
End Type                ' I used C++ ... Hope I got it right...

Public Type PBE_CONFIG
  dwConfig As Long      '// BE_CONFIG_XXXXX
  format As PBE_FORMAT
End Type



Public Type PBE_VERSION
  '// BladeEnc DLL Version number
    byDLLMajorVersion As Byte
    byDLLMinorVersion As Byte

  '// BladeEnc Engine Version Number

  byMajorVersion As Byte
  byMinorVersion As Byte

  '// DLL Release date

  byDay As Byte
  byMonth As Byte
  wYear As Integer

  '// BladeEnc Homepage URL

  zHomepage As String * BE_MAX_HOMEPAGE

  byAlphaLevel As Byte
  byBetaLevel As Byte
  byMMXEnabled As Byte

  btReserved(1 To 125) As Byte
End Type

Public Declare Sub beVersion Lib "lame_enc.dll" (pbeVersion As PBE_VERSION)
Public Declare Function beInitStream Lib "lame_enc.dll" (ByVal pbeConfig As Long, ByVal dwSamples As Long, ByVal dwBufferSize As Long, ByVal phbeStream As Long) As Long
Public Declare Function beEncodeChunk Lib "lame_enc.dll" (ByVal hbeStream As Long, ByVal nSamples As Long, ByVal pSamples As Long, ByVal pOutput As Long, ByVal pdwOutput As Long) As Long
Public Declare Function beDeinitStream Lib "lame_enc.dll" (ByVal hbeStream As Long, ByVal pOutput As Long, ByVal pdwOutput As Long) As Long
Public Declare Function beCloseStream Lib "lame_enc.dll" (ByVal hbeStream As Long) As Long
Public Declare Function beWriteVBRHeader Lib "lame_enc.dll" (ByVal lpszFileName As String) As Long

Public Function GetErrorString(ErrorNo As Long) As String
  Select Case ErrorNo
  Case 0
    GetErrorString = "No Error"
  Case 1
    GetErrorString = "Invalid format"
  Case 2
    GetErrorString = "Invalid format parameters"
  Case 3
    GetErrorString = "No more handles"
  Case 4
    GetErrorString = "Invalid handle"
  Case 5
    GetErrorString = "Buffer too small"
  End Select
End Function
