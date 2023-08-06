VERSION 5.00
Begin VB.Form frmConvert 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   732
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3228
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   732
   ScaleWidth      =   3228
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Text            =   "testcase.wav"
      Top             =   2064
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   648
      Top             =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please wait... WAV to MP3...."
      Height          =   228
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2964
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   708
      Left            =   24
      Top             =   24
      Width           =   3204
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00.00%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   384
      Width           =   2952
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Text1.text = Form1.txtSalva.text
  SetUnhandledExceptionFilter AddressOf MyExceptionFilter

End Sub

Private Sub Form_Unload(Cancel As Integer)
  SetUnhandledExceptionFilter 0
End Sub



Private Sub Status_Click()

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False


  ' Now.. This one is going to be little bit complicated....
  
  On Error GoTo ErrHandler
  ChDrive App.Path
  ChDir App.Path
  
  ' Check that file exists...
  
  Cancelled = False
  
  ' Fill beConfig structure....
  Dim beConfig As PBE_CONFIG
  beConfig.dwConfig = BE_CONFIG_LAME
  
  With beConfig.format.LHV1
    '// this are the default settings for testcase.wav
    .dwStructVersion = 1
    .dwStructSize = Len(beConfig)
    .dwSampleRate = 44100         '// INPUT FREQUENCY
    .dwReSampleRate = 0           '// DON"T RESAMPLE
    .nMode = BE_MP3_MODE_JSTEREO  '// OUTPUT IN STREO
    '.dwBitrate = 128             '// MINIMUM BIT RATE
    '.nPreset = LQP_HIGH_QUALITY  '// QUALITY PRESET SETTING
    .dwMpegVersion = MPEG1        '// MPEG VERSION (I or II)
    .dwPsyModel = 0               '// USE DEFAULT PSYCHOACOUSTIC MODEL
    .dwEmphasis = 0               '// NO EMPHASIS TURNED ON
    .bOriginal = True             '// SET ORIGINAL FLAG
    .bNoRes = True                '// No Bit resorvoir
    
    'Select Case 8
      'Case 0: .nPreset = LQP_LOW_QUALITY
      'Case 1: .nPreset = LQP_NORMAL_QUALITY
      'Case 2: .nPreset = LQP_HIGH_QUALITY
      'Case 3: .nPreset = LQP_VOICE_QUALITY
      'Case 4: .nPreset = LQP_PHONE
      'Case 5: .nPreset = LQP_RADIO
      'Case 6: .nPreset = LQP_TAPE
      'Case 7: .nPreset = LQP_HIFI
      .nPreset = LQP_CD
      'Case 9: .nPreset = LQP_STUDIO
      'Case Else
        'MsgBox "You didn't select quality..."
        'Exit Sub
    'End Select
      .dwBitrate = Val(160)
  End With

  Dim error As Long
  Dim dwSamples As Long, dwMP3Buffer As Long, hbeStream As Long
  
  ' Init MP3 Stream
  error = beInitStream(VarPtr(beConfig), VarPtr(dwSamples), VarPtr(dwMP3Buffer), VarPtr(hbeStream))
    
  '// Check result
  If error <> BE_ERR_SUCCESSFUL Then
    Err.Raise error, "Lame", GetErrorString(error)
  End If
  
  
  ' Open Files...
  Dim toRead As Long, toWrite As Long
  Dim Done As Long
  Dim length As Long
  
  length = FileLen(Text1)
  
  Dim ReadFile As clsFileIo
  Set ReadFile = New clsFileIo
  ReadFile.OpenFile Text1
  
  Dim WriteFile As clsFileIo
  Set WriteFile = New clsFileIo
  'WriteFile.OpenFile Text1.text & ".mp3"
  WriteFile.OpenFile ChangeExt(Text1, "mp3")
  
  ' Allocate memory for buffers... :)
  Dim WavPtr1 As Long
  Dim WavPtr2 As Long
  Dim MP3Ptr1 As Long
  Dim MP3Ptr2 As Long
  WavPtr1 = GlobalAlloc(&H40, dwSamples * 2)
  WavPtr2 = GlobalLock(WavPtr1)
  MP3Ptr1 = GlobalAlloc(&H40, dwMP3Buffer)
  MP3Ptr2 = GlobalLock(MP3Ptr1)
  
  'Skip WAV header
  Dim Temp(1 To 44) As Byte
  Call ReadFile.ReadBytes(VarPtr(Temp(1)), 44)
  
  ' And here we go....
  Do While Done < length
    '//set up how much to readinto the buffer
    If Done + dwSamples * 2 < length Then
      toRead = dwSamples * 2
    Else
      toRead = length - Done
    End If
    
    ' Read into buffer
    Call ReadFile.ReadBytes(WavPtr2, toRead)
    
    Done = Done + toRead
    toRead = toRead / 2
    
    ' Encode buffer
    error = beEncodeChunk(hbeStream, toRead, WavPtr2, MP3Ptr2, VarPtr(toWrite))
    
    ' Check result...
    If error <> BE_ERR_SUCCESSFUL Then
      Call beCloseStream(hbeStream)
      Err.Raise error, "Lame", GetErrorString(error)
    End If
    
    ' Write Buffer...
    If toWrite > 0 Then
      Call WriteFile.WriteBytes(MP3Ptr2, toWrite)
    End If
    
    ' Report status to user....
    Status.Caption = format(Done / length * 100, "00.00") & "%"
    If Cancelled = True Then Exit Do
    DoEvents
  Loop
  
  ' Deinitialize stream and write last bytes to MP3
  error = beDeinitStream(hbeStream, MP3Ptr2, VarPtr(toWrite))

  If toWrite > 0 Then
    WriteFile.WriteBytes MP3Ptr2, toWrite
  End If
  
  
  ' Clear buffers....
  GlobalFree MP3Ptr2
  GlobalFree WavPtr2
  
  ' Close files
  Call WriteFile.CloseFile
  Call ReadFile.CloseFile
  
  Set WriteFile = Nothing
  Set ReadFile = Nothing
  
  ' Close stream
  Call beCloseStream(hbeStream)
  
  
  ' WriteVBRHeader (if we use variable bitrate...)
  'Call beWriteVBRHeader(ChangeExt(Text1, "mp3"))
  Form1.Enabled = True
  Form1.Frame1.Enabled = True
  Form1.Frame2.Enabled = True
  Form1.Command1.Enabled = True
  Form1.Command1.Caption = "CONVERT"
  Unload Me

  Exit Sub
  
ErrHandler:
  ' Damn.. Something went wrong and this one should tell what...
  ' At time of debugin.. This one was place where code ended all the time...
  ' But now it should never come into this point....
  
  MsgBox Err.Description, vbCritical, "Critical error..."
  If WavPtr2 Then GlobalFree WavPtr2
  If MP3Ptr2 Then GlobalFree MP3Ptr2
  
  WriteFile.CloseFile
  ReadFile.CloseFile
  Set WriteFile = Nothing
  Set ReadFile = Nothing
  
  Err.Clear
  Exit Sub
End Sub


