VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSoundRec 
   BackColor       =   &H00AFAFF3&
   BorderStyle     =   0  'None
   Caption         =   "Gravador de SOM"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11955
   Icon            =   "frmSoundRec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSoundRec.frx":0582
   ScaleHeight     =   5325
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar scrLevel 
      Height          =   255
      LargeChange     =   300
      Left            =   1320
      SmallChange     =   300
      TabIndex        =   15
      Top             =   3990
      Width           =   2595
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00D89970&
      Height          =   1485
      Left            =   4920
      ScaleHeight     =   1425
      ScaleWidth      =   3825
      TabIndex        =   7
      Top             =   3165
      Width           =   3885
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00D89970&
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   1095
         TabIndex        =   12
         Top             =   660
         Width           =   1575
         Begin VB.CheckBox chkStereo 
            BackColor       =   &H00D89970&
            Caption         =   "Stereo / Mono"
            Height          =   195
            Left            =   75
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D89970&
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   105
         TabIndex        =   8
         Top             =   45
         Width           =   3615
         Begin VB.OptionButton opt11 
            BackColor       =   &H00D89970&
            Caption         =   "11025 Hz"
            Height          =   195
            Left            =   2400
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt225 
            BackColor       =   &H00D89970&
            Caption         =   "22050 Hz"
            Height          =   195
            Left            =   1320
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt44 
            BackColor       =   &H00D89970&
            Caption         =   "44100 Hz"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   60
      Left            =   4770
      Top             =   2745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   9915
      TabIndex        =   4
      Top             =   4005
      Width           =   1650
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   2400
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save Wave File"
      Filter          =   "*.WAV"
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Salvar como..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   9915
      TabIndex        =   1
      Top             =   2370
      Width           =   1650
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Iniciar Gravação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   9915
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   750
      Width           =   1650
   End
   Begin VB.Label Label1 
      BackColor       =   &H0019892E&
      BackStyle       =   0  'Transparent
      Caption         =   "Min            Nível            Max"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1350
      TabIndex        =   17
      Top             =   3735
      Width           =   2520
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0019892E&
      BackStyle       =   0  'Transparent
      Caption         =   "Volume de Gravação:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1215
      TabIndex        =   16
      Top             =   3375
      Width           =   2805
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configurações da gravação:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5460
      TabIndex        =   14
      Top             =   2835
      Width           =   2790
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0019892E&
      BackStyle       =   0  'Transparent
      Caption         =   "Fim: 00:00:00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4410
      TabIndex        =   6
      Top             =   1245
      Width           =   1500
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0019892E&
      BackStyle       =   0  'Transparent
      Caption         =   "Início: 00:00:00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4290
      TabIndex        =   5
      Top             =   915
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo de gravação:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4140
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   2115
      Index           =   1
      Left            =   2970
      MousePointer    =   5  'Size
      TabIndex        =   3
      Top             =   0
      Width           =   4440
   End
End
Attribute VB_Name = "frmSoundRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type TS
    sfld As String * 255
End Type
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim m_Rgn As CBMPRegion
Private mCaptionlessWindowMover As CCaptionlessWindowMover
Dim nome As String
Dim p_ret As String
Dim volCtrl As MIXERCONTROL ' waveout volume control
Dim micCtrl As MIXERCONTROL ' microphone volume control
Dim rc As Long              ' return code
Dim ok As Boolean           ' boolean return code
Dim sFile As String
Dim blFileSaved As Boolean

Private Sub Check1_Click()
Dim sFmt As String * 255
Dim iRet As Integer

   If (Check1.Value = 1) Then
   
   If (Not blFileSaved) And (Trim(sFile) <> "") Then
        
        iRet = MsgBox("A gravação anterior não foi salva, deseja continuar?", vbYesNoCancel)
        If iRet = vbCancel Then
            'Do nothing
            Exit Sub
        ElseIf iRet = vbNo Then
            cmdSave.Value = True
        Else
            Kill sFile
        End If
        
   End If
   
    blFileSaved = False
    cmdSave.Enabled = False
    Check1.Caption = "Parar Gravação"
    'Create the Wave File
  
    Label5.Caption = "Início: " & Time
    sFile = App.Path & "\" & Format(Now, "MMDDYYYYHH24MISSAMPM") & ".WAV"
    hmmioIn = mmioOpen(sFile, mmioinf, (MMIO_CREATE Or MMIO_WRITE))  'Or MMIO_ALLOCBUF
    If hmmioIn = 0 Then
      MsgBox "Falha ao criar arquivo WAV!"
      Exit Sub
    End If
    
    'Set The WAV wformat
    wformat.wFormatTag = 1
    If chkStereo.Value = 1 Then
        wformat.nChannels = 2
    Else
        wformat.nChannels = 1
    End If
    
    wformat.wBitsPerSample = 16
    If opt44.Value = True Then
        wformat.nSamplesPerSec = 44100
    ElseIf opt11.Value = True Then
        wformat.nSamplesPerSec = 11025
    ElseIf opt225.Value = True Then
        wformat.nSamplesPerSec = 22050
    End If
    wformat.nBlockAlign = wformat.nChannels * wformat.wBitsPerSample / 8
    wformat.nAvgBytesPerSec = wformat.nSamplesPerSec * wformat.nBlockAlign
    wformat.cbSize = Len(wformat)
    
    'Create the RIFF Chunk
    mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
    If (mmioCreateChunk(hmmioIn, mmckinfoParentIn, MMIO_CREATERIFF) <> 0) Then
        MsgBox "Falha ao criar RIFF CHUNK"
        mmioClose hmmioIn, 0
        Exit Sub
    End If
    
    'Create the Fmt Chunk
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("fmt", 0)
    mmckinfoSubchunkIn.ckSize = Len(wformat)
    If (mmioCreateChunk(hmmioIn, mmckinfoSubchunkIn, 0) <> 0) Then
        MsgBox "Falha ao criar FMT CHUNK"
        mmioClose hmmioIn, 0
        Exit Sub
    End If
    
    CopyStringFromStruct sFmt, wformat, Len(wformat)
    
    If (mmioWrite(hmmioIn, sFmt, Len(wformat)) <> Len(wformat)) Then
        MsgBox "Não foi possível escrever o formato WAVE"
        mmioClose hmmioIn, 0
        Exit Sub
    End If
    
    'Create Data Chunk
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("data", 0)
    If (mmioCreateChunk(hmmioIn, mmckinfoSubchunkIn, 0) <> 0) Then
        MsgBox "Falha ao criar DATA CHUNK"
        mmioClose hmmioIn, 0
        Exit Sub
    End If
    Frame1.Enabled = False
    StartInput  ' Start receiving audio input

   Else
      Frame1.Enabled = True
      StopInput   ' Stop receiving audio input
      Label6.Caption = "Fim: " & Time
            Check1.Caption = "Iniciar Gravação"
      cmdSave.Enabled = True
   End If
   
End Sub

Private Sub cmdSave_Click()
Dim o_file As New Scripting.FileSystemObject
    On Error Resume Next
    cmdlg.ShowSave
    If cmdlg.FileName <> "" Then
        o_file.CopyFile sFile, cmdlg.FileName & ".Wav", False
        Kill sFile
        blFileSaved = True
    End If
Set o_file = Nothing
End Sub

Private Sub Command1_Click()
If MsgBox("Deseja realmente sair?", vbQuestion + vbYesNo) = vbYes Then
End
End If
End Sub



Private Sub Command2_Click()
  Call Shell(nome, vbNormalFocus)
  Command2.Visible = False
End Sub

Private Sub Form_Load()
Set m_Rgn = New CBMPRegion
Set mCaptionlessWindowMover = New CCaptionlessWindowMover
  Set mCaptionlessWindowMover.Form = Me
  m_Rgn.CreateFromPic Me.Picture, vbBlack
  SetWindowRgn hwnd, m_Rgn.Handle, True

If Len(App.Path) = 3 Then
nome = App.Path & "Osciloscópio.exe"
Else
nome = App.Path & "\Osciloscópio.exe"
End If


  ''''ORIGINAL CODE
  
    Initialize Me.hwnd
    chkStereo.Value = 1
    opt44.Value = True
    cmdSave.Enabled = False
    scrLevel.Value = 32767
    
    ' Open the mixer with deviceID 0.
    rc = mixerOpen(hMixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Couldn't open the mixer."
        Exit Sub
    End If
    
     'Get the wavein volume control
    ok = GetVolumeControl(hMixer, _
                     MIXERLINE_COMPONENTTYPE_DST_WAVEIN, _
                     MIXERCONTROL_CONTROLTYPE_VOLUME, _
                     volCtrl)
    If (ok = True) Then
        'Adjust the scroll bar
    End If

End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  mCaptionlessWindowMover.HandleMouseDown x, y
End Sub

Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Handle the form's MouseMove event
  mCaptionlessWindowMover.HandleMouseMove x, y
End Sub

Private Sub Label4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  mCaptionlessWindowMover.HandleMouseUp
End Sub

Private Sub scrLevel_Change()
    SetVolumeControl hMixer, volCtrl, (scrLevel.Value * 1.8)
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim iRet As Integer

   If (fRecording = True) Then
       StopInput
   End If
   
   If (Not blFileSaved) And (Trim(sFile) <> "") Then
        
        iRet = MsgBox("Quit Without Saving ?", vbYesNoCancel, "WAV Recorder")
        If iRet = vbCancel Then
            Cancel = 1
        ElseIf iRet = vbNo Then
            cmdSave.Value = True
        Else
            Kill sFile
        End If
        
   End If
 
End Sub

Public Sub Gandu()
Dim iRet As Long
Dim sBuff  As String

   ' Process sound buffer if recording
   If (fRecording) And Check1.Value = 1 Then
      For i = 0 To (NUM_BUFFERS - 1)
         If inHdr(i).dwFlags And WHDR_DONE Then
            rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
            If rc <> 0 Then
                MsgBox "Failed"
                Exit Sub
            End If

            sBuff = Space(BUFFER_SIZE)

            CopyMemory ByVal sBuff, ByVal inHdr(i).lpData, BUFFER_SIZE

            mmioWrite hmmioIn, sBuff, BUFFER_SIZE

         End If
      Next
   End If
   
End Sub


