VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmDownload 
   BackColor       =   &H00F4C686&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "1º Arcanjo Downloader"
   ClientHeight    =   2655
   ClientLeft      =   3795
   ClientTop       =   3645
   ClientWidth     =   8835
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   8835
   Begin VB.Timer tmrTitle 
      Interval        =   2500
      Left            =   7155
      Top             =   15
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
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
      Left            =   7860
      TabIndex        =   19
      Top             =   2280
      Width           =   300
   End
   Begin VB.TextBox txtDest 
      Appearance      =   0  'Flat
      BackColor       =   &H00DADAFC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   675
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2280
      Width           =   7155
   End
   Begin VB.ListBox lstDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6FEDA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00337133&
      Height          =   1200
      ItemData        =   "frmDownload.frx":08CA
      Left            =   75
      List            =   "frmDownload.frx":08CC
      TabIndex        =   16
      Top             =   375
      Width           =   2760
   End
   Begin VB.Timer tmrTimeLeft 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7575
      Top             =   15
   End
   Begin VB.Timer tmrUpdateProgress 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7995
      Top             =   15
   End
   Begin MSWinsockLib.Winsock sckDownload 
      Left            =   8415
      Top             =   15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraProgresso 
      BackColor       =   &H00F4C686&
      Caption         =   "Progresso do download"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   3000
      TabIndex        =   2
      Top             =   150
      Width           =   5775
      Begin VB.PictureBox picProgresso 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFAEA&
         FillColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CF8D45&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   5505
         TabIndex        =   3
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tamanho:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recebidos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   13
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidade:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   12
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo Restante:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo Percorrido:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label lblTRestante 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblTPercorrido 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label lblVelocidade 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblRecebido 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblTam 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   855
         TabIndex        =   5
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F7FED6&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   5535
      End
   End
   Begin VB.CommandButton cmdDownload 
      BackColor       =   &H00FFFAEA&
      Caption         =   "&Download"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1695
      TabIndex        =   1
      Top             =   1635
      Width           =   1095
   End
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      BackColor       =   &H00DADAFC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7290
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destino do arquivo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   705
      MousePointer    =   2  'Cross
      TabIndex        =   20
      Top             =   2055
      Width           =   1410
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programas disponíveis na página:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   17
      Top             =   150
      Width           =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.gabrielfalcao.i8.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4875
      TabIndex        =   15
      Top             =   1845
      Width           =   2340
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim currpos As String
Dim CurrentTime As String
Dim TotalFrames As String
Dim TotalTime As String
Dim FramesPerSecond As String
Dim Paused As Boolean
Dim MP3Path As String
Dim PLVisible As Boolean
Dim temp As String
Dim ok As Boolean
Dim showplaya As Boolean
Dim j As Integer
Dim strCommand As String
Dim SlideFlag As Boolean
Dim IX, IY, TX, TY, FX, FY
Dim tmptit As String
Dim tmpart As String

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type
Dim mp3file As String, v As Single
Dim tmpTitle As String, tmpTrack As String, tmpAlbum As String, tmpArtist As String, tmpComment As String, tmpYear As String


Private Declare Function ShellExecute Lib "Shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private m_sDATA                         As String
Private Percent                         As Integer
Private BeginTransfer                   As Single

Private Header                          As Variant
Private Status                          As String
Private TransferRate                    As Single

Private bDownloadPaused                 As Boolean
Private bDownloadComplete               As Boolean

Private strSvrURL                       As String
Private strSvrPort                      As String
Private URL                             As String
Private strSalvarEm                     As String
Private Filename                        As String
Private FileLength                      As Single
Private Sec                             As Integer
Private Min                             As Integer
Private Hr                              As Integer

Private BytesAlreadySent                As Single
Private BytesRemaining                  As Single

Public Function InStrRevVB5(ByVal StringCheck As String, ByVal StringMatch As String, Optional ByVal Start As Long = -1) As Long
    Dim lPos        As Long
    Dim lSavePos    As Long
    If Start = -1 Then Start = Len(StringCheck)
    lPos = InStr(1, StringCheck, StringMatch, vbBinaryCompare)
    While lPos > 0 And lPos < Start
        lSavePos = lPos
        lPos = InStr(lPos + 1, StringCheck, StringMatch, vbBinaryCompare)
    Wend
    InStrRevVB5 = lSavePos
End Function

Private Function GETDATAHEAD(DATA As Variant, ToRetrieve As String)
    Dim EndBYTES                        As Integer
    Dim A                               As String
    Dim LENGTHEND                       As Integer
    Dim PART                            As Integer
    Dim Part2                           As Integer
    Dim RetrieveLength                  As Integer
    On Error Resume Next
    If DATA = "" Then Exit Function
    If InStr(DATA, ToRetrieve) > 0 Then
        LENGTHEND = Len(DATA)
        PART = InStr(DATA, ToRetrieve)
        RetrieveLength = Len(ToRetrieve)
        A = Right(DATA, LENGTHEND - PART - RetrieveLength)
        LENGTHEND = Len(A)
        If InStr(A, vbCrLf) > 0 Then
            Part2 = InStr(A, vbCrLf)
            A = Left(A, Part2 - 1)
        End If
        GETDATAHEAD = A
    End If
End Function

Public Function StartUpdate(ByVal strURL As String)
    Dim Pos                             As Integer
    Dim LENGTH                          As Integer
    Dim NextPos                         As Integer
    Dim LENGTH2                         As Integer
    Dim POS2                            As Integer
    Dim POS3                            As Integer
    BytesAlreadySent = 1
    If strURL = "" Then
        Exit Function
    End If
    URL = strURL
    Pos = InStr(strURL, "://") 'Record position of ://
    LENGTH2 = Len("://") 'Record the length of it
    LENGTH = Len(strURL) 'Length of the entire url
    If InStr(strURL, "://") Then  ' check if they entered the http:// or ftp://
        strURL = Right(strURL, LENGTH - LENGTH2 - Pos + 1) ' remove http:// or ftp://
    End If
    If InStr(strURL, "/") Then 'looks for the first / mark going from left to right
        POS2 = InStr(strURL, "/") 'gets the position of the / mark
        '-----------------GET THE FILENAME-------------
        Dim strFile                     As String
        strFile = strURL 'load the variables into each other
        Do Until InStr(strFile, "/") = 0 'Do the loop until all is left is the filename
            LENGTH2 = Len(strFile) 'get the length of the filename every time its passed over by the loop
            POS3 = InStr(strFile, "/") 'find the / mark
            strFile = Right(strURL, LENGTH2 - POS3) 'slash it down removing everything before the / mark including the / mark...
        Loop
        
            If InStr(strFile, ":") Then
                Filename = Left(strFile, InStr(strFile, ":") - 1)
            Else
                Filename = strFile
            End If
            
        strSvrURL = Left(strURL, POS2 - 1) 'removes everything after the / mark leaving just the server name as the end result
    End If
    '-----------END TRIM THE URL FOR THE SERVER NAME-----------
End Function

Private Sub CloseSocket()
    Do Until sckDownload.State = 0
        sckDownload.Close
        sckDownload.LocalPort = 0
        Close #1
    Loop
End Sub

Private Sub cmdDownload_Click()
If IsConnected = True Then
    tmrTimeLeft.Enabled = True
    tmrUpdateProgress.Enabled = True
   picProgresso.Visible = True
    lblStatus.Visible = False
    strSvrURL = txtURL
    strSvrPort = 80
    
    StartUpdate txtURL
    
    strSalvarEm = txtDest.Text
    
    sckDownload.Connect strSvrURL, strSvrPort
    Command1.Enabled = False
     '  lstDown.Enabled = False
    cmdDownload.Enabled = False
 Else
 Me.Caption = "Computador está desconectado da internet!"
 MsgBox "O Computador está desconectado da internet, verifique a sua conexão e tente novamente", vbInformation, "1º Arcanjo Downloader"
 End If
End Sub



Private Sub Command1_Click()
  Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
  
  szTitle = vbCr & vbCr & "Selecione a Pasta Desejada:"
  
  With tBrowseInfo
    .hwndOwner = Me.hwnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    If Not Right$(sBuffer, 1) = "\" Then
    strSalvarEm = sBuffer & "\" & Mid$(txtURL, InStrRevVB5(txtURL, "/") + 1)
    Else
    strSalvarEm = sBuffer & Mid$(txtURL, InStrRevVB5(txtURL, "/") + 1)
    txtDest.Text = sBuffer & Mid$(txtURL, InStrRevVB5(txtURL, "/") + 1)
    End If
  End If
End Sub

Private Sub Form_Load()
sckDownload.Close
txtDest.Text = App.Path & "\" & Mid$(txtURL, InStrRevVB5(txtURL, "/") + 1)
'MAHP
lstDown.AddItem "TeraZip" 'http://gabrielfalcaogm.vila.bol.com.br/InstalarTeraZip.exe
lstDown.AddItem "N£OWEB" 'http://www.megaaccesshp.hpg.com.br/neoweb.exe
lstDown.AddItem "TeraCleaner"
lstDown.AddItem "Trava Tudo 1.00" 'http://www.megaaccesshp.hpg.com.br/Instalar%20Trava%20Tudo.exe
lstDown.AddItem "Winzip 8.0" 'http://www.megaaccesshp.hpg.com.br/winzip80.exe
lstDown.AddItem "VB Runtime 6.0" 'http://www.megaaccesshp.hpg.com.br/Vbrun60.exe
lstDown.AddItem "PowerOff XP" 'http://www.megaaccesshp.hpg.ig.com.br/poweroff.zip
lstDown.AddItem "Ordem de Serviço 1.00" 'http://www.megaaccesshp.hpg.ig.com.br/Instalar%20OS.exe
lstDown.ListIndex = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HC00000
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If sckDownload.State = 7 Then
If MsgBox("Existe um download em decorrência, deseja sair assim mesmo?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
    tmrTimeLeft.Enabled = False
    tmrUpdateProgress.Enabled = False
    CloseSocket
    End If
    End If
End Sub

Private Sub fraProgresso_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HC00000
End Sub

Private Sub Label1_Click()
    Dim lRet As Long
    lRet = ShellExecute(0, "open", Label1.Caption, vbNullString, vbNullString, 1)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HC0&
End Sub

Private Sub List1_Click()
cmdDownload.Enabled = True
End Sub

Private Sub lstDown_Click()
If sckDownload.State <> 7 Then
Select Case lstDown.Text
Case Is = "N£OWEB"
txtURL.Text = "http://www.megaaccesshp.hpg.ig.com.br/neoweb.exe"
Case Is = "Trava Tudo 1.00"
txtURL.Text = "http://www.megaaccesshp.hpg.ig.com.br/Instalar%20Trava%20Tudo.exe"
Case Is = "Winzip 8.0"
txtURL.Text = "http://www.megaaccesshp.hpg.ig.com.br/winzip80.exe"
Case Is = "VB Runtime 6.0"
txtURL.Text = "http://www.megaaccesshp.hpg.ig.com.br/Vbrun60.exe"
Case Is = "PowerOff XP"
txtURL.Text = "http://www.megaaccesshp.hpg.ig.com.br/poweroff.zip"
Case Is = "TeraCleaner"
txtURL.Text = "http://www.megaaccesshp.hpg.ig.com.br/Teracleaner.zip"
Case Is = "TeraZip"
txtURL.Text = "http://www.megaaccesshp.hpg.ig.com.br/InstalarTeraZip.zip"
Case Is = "Ordem de Serviço 1.00"
txtURL.Text = "http://www.megaaccesshp.hpg.ig.com.br/Instalar%20OS.zip"
End Select
txtDest.Text = App.Path & "\" & Mid$(txtURL, InStrRevVB5(txtURL, "/") + 1)
cmdDownload.Enabled = True
End If
End Sub

Private Sub sckDownload_Close()
    picProgresso.Visible = False
    lblStatus.Visible = True
    lblStatus.Caption = "Download Completo"
    Command1.Enabled = True
       lstDown.Enabled = True
    cmdDownload.Enabled = True
     sckDownload.Close
End Sub

Private Sub sckDownload_Connect()
    'On Error Resume Next
    
    Dim strCommand As String
    
    strCommand = "GET " + Right(URL, Len(URL) - Len(strSvrURL) - 7) + " HTTP/1.0" + vbCrLf
    strCommand = strCommand + "Accept: *.*, */*" + vbCrLf
    
    strCommand = strCommand + "User-Agent: Primeiro Arcando Downloader" & vbCrLf
    strCommand = strCommand + "Referer: " & strSvrURL & vbCrLf
    strCommand = strCommand + "Host: " & strSvrURL & vbCrLf
    
    strCommand = strCommand + vbCrLf
    sckDownload.SendData strCommand 'sends a header to the server instructing it what to do!
    BeginTransfer = Timer 'start timer for transfer rate
End Sub

Private Sub sckDownload_DataArrival(ByVal bytesTotal As Long)
    Dim Pos                             As Integer
    Dim LENGTH                          As Integer
    Dim HEAD                            As String
    Debug.Print bytesTotal
    sckDownload.GetData m_sDATA, vbString
    
    If InStr(LCase(m_sDATA), "content-type:") Then 'find out if this chunk has the header..you can change that to anything that the header contains
        If InStr(LCase(m_sDATA), "404 not found") > 0 Then
                MsgBox "O arquivo solicitado não foi encontrado no servidor!" & vbCrLf & vbCrLf & "Razões possíveis:" & vbCrLf & "- O arquivo não existe no servidor" _
                & vbCrLf & "- O endereço é um script e os dados são inválidos" & vbCrLf & "- Endereço solicitado está errado" & vbCrLf & "- O servidor está extremamente ocupado" _
                & vbCrLf & vbCrLf & "Você pode tentar baixar novamente.  Se o erro continuar, o endereço está errado.", , "Arquivo não encontrado"
                Reset
                CloseSocket
                Exit Sub
        End If
   
        Pos = InStr(m_sDATA, vbCrLf & vbCrLf) ' find out where the header and the data is split apart
        LENGTH = Len(m_sDATA) 'get the length of the data chunk
        HEAD = Left(m_sDATA, Pos - 1) 'Get the header from the chunk of data and ignore the data content
        m_sDATA = Right(m_sDATA, LENGTH - Pos - 3) 'Get the data from the first chunk that contains the header also
        Header = Header & HEAD 'Append the header to header text box
        
        BytesRemaining = GETDATAHEAD(Header, "Content-Length:")
        
        'frmHeader.txtHeader = Header
    End If
    '-----------BEGIN WRITE CHUNK TO FILE CODE--------
    Open strSalvarEm For Binary Access Write As #1 'opens file for output
    Put #1, BytesAlreadySent, m_sDATA 'writes data to the end of file
    BytesAlreadySent = Seek(1)
    Close #1 'close file for now until next data chunk is available
    '--------------------------------------------------
    
    'If you dont subtract the difference you will get a really large and odd download speed hehe.
    TransferRate = Format(Int((BytesAlreadySent - FileLength) / (Timer - BeginTransfer)) / 1000, "####.00")
End Sub

Private Sub tmrTimeLeft_Timer()
    'On Error Resume Next
    If BytesRemaining > 0 And BytesAlreadySent > 0 And TransferRate > 0 Then
        If BytesRemaining <= BytesAlreadySent Then
            lblVelocidade = 0
            CloseSocket
            lblTPercorrido = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
            'cmdDownload.Enabled = False
            picProgresso.Visible = False
            lblStatus.Visible = True
            lblStatus.Caption = "Download Completo"
            'Reset
        Else
            Sec = Sec + 1
            If Sec >= 60 Then
                Sec = 0
                Min = Min + 1
            ElseIf Min >= 60 Then
                Min = 0
                Hr = Hr + 1
            End If
            'cmdDownload.Enabled = True
            'cmdRun.Enabled = False
            lblTPercorrido = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
            'The reason I divide the difference of bytesalreadysent and bytesremaining is becuase they are in bytes right now.. I want it to be in KB so it can be Kbps and not bps
            lblTRestante = ConvertTime(Int(((BytesRemaining - BytesAlreadySent) / 1024) / TransferRate))
            lblVelocidade = Format(TransferRate, "##.#0#") & " kb/s"

        End If
    End If
End Sub

Private Sub tmrTitle_Timer()
Me.Caption = "1º Arcanjo Downloader"
End Sub

Private Sub tmrUpdateProgress_Timer()
'    On Error Resume Next
    If BytesAlreadySent > 0 Then 'And BytesRemaining > 0 Then

        lblRecebido = File_ByteConversion(BytesAlreadySent)
        If BytesRemaining = 0 Then
            lblTam = "Desconhecido"
        Else
            lblTam = File_ByteConversion(BytesRemaining)
        End If
            If lblTam <> "Desconhecido" Then
            Percent = Format((BytesAlreadySent / BytesRemaining) * 100, "00") 'calculates the percentage completed
            UpdateProgress picProgresso, Percent, True  'updates progress bar with new percentage rate
        End If
    End If
End Sub

Public Function File_ByteConversion(NumberOfBytes As Single) As String
    On Error Resume Next
    If NumberOfBytes < 1024 Then 'checks to see if its so small that it cant be converted into larger grouping
        File_ByteConversion = NumberOfBytes & " Bytes"
    End If
    If NumberOfBytes > 1024 Then  'Checks to see if file is big enough to convert into KB
        File_ByteConversion = Format(NumberOfBytes / 1024, "0.00") & " KB"
    End If
    If NumberOfBytes > 1048576 Then  'Checks to see if its big enough to convert into MB
        File_ByteConversion = Format(NumberOfBytes / 1048576, "###,###,##0.00") & " MB"
    End If
End Function

Public Function ConvertTime(ByVal TheTime As Single) As String
    Dim NewTime                         As String
    Dim Sec                             As Single
    Dim Min                             As Single
    Dim H                               As Single
    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function

Public Function UpdateProgress(pb As Control, ByVal Percent As Integer, Optional ByVal ShowPercent = False)
    'Replacement for progress bar..looks nicer also
    Dim sNum                            As String    'use percent
    'Dim Num$
    If Not pb.AutoRedraw Then 'picture in memory ?
        pb.AutoRedraw = -1 'no, make one
    End If
    pb.Cls 'clear picture in memory
    pb.ScaleWidth = 100 'new sclaemodus
    pb.DrawMode = 10 'not XOR Pen Modus
    If ShowPercent = True Then
    Num$ = Format$(Percent, "###0") + "%"
    pb.CurrentX = 50 - pb.TextWidth(Num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(Num$)) / 2
    pb.Print Num$ 'print percent
    End If
    pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
    pb.Refresh 'show differents
End Function

Private Sub txtDest_Change()
txtDest.SelStart = Len(txtDest.Text)
End Sub
