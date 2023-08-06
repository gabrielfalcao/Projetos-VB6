VERSION 5.00
Begin VB.Form frmGenerator 
   BackColor       =   &H00D89970&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de discos DM"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmGenerator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFormatar 
      Caption         =   "Formatar Disco"
      Height          =   975
      Left            =   3352
      Picture         =   "frmGenerator.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1140
      Width           =   1260
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5475
      Top             =   0
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar Disco"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1282
      Picture         =   "frmGenerator.frx":5C24
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1140
      Width           =   1260
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFD1AD&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   60
      ScaleHeight     =   240
      ScaleWidth      =   5745
      TabIndex        =   5
      Top             =   2190
      Width           =   5775
   End
   Begin VB.PictureBox picProgresso 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFD1AD&
      FillColor       =   &H00C00000&
      ForeColor       =   &H00CF8D45&
      Height          =   330
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   2520
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   165
      Picture         =   "frmGenerator.frx":6E96
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   1965
      Width           =   600
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D89970&
      Caption         =   "www.gabrielfalcao.i8.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3315
      MouseIcon       =   "frmGenerator.frx":C678
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   705
      Width           =   2310
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D89970&
      Caption         =   "gabrielfalcao@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1080
      MouseIcon       =   "frmGenerator.frx":C982
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   705
      Width           =   2310
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "por Gabriel Falcão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1020
      TabIndex        =   2
      Top             =   495
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gerador de Discos DM"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   945
      TabIndex        =   1
      Top             =   240
      Width           =   3600
   End
End
Attribute VB_Name = "frmGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim contador
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
Private status                          As String
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
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwnd As Long, ByVal Drive As Long, ByVal fmtID As Long, ByVal options As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Function pbar(pb As Control, ByVal Percent As Integer, Optional ByVal ShowPercent = False)
    'Replacement for progress bar..looks nicer also
    Dim sNum                            As String    'use percent
    Dim Num$
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
Private Function EspacoLivre(Drive As String) As Integer
On Error Resume Next
  Dim fso, d, s, t
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set d = fso.GetDrive(fso.GetDriveName(Drive))
    
EspacoLivre = FormatNumber(d.FreeSpace / 1024, 0)
End Function
Private Sub des(id As String, destino As String)
On Error Resume Next
Dim p_ret As String
Dim disc As String
disc = "A:\"
p_ret = StrConv(LoadResData(id, "FILES"), vbUnicode)
Open disc & id For Binary As #1
Put #1, , p_ret
Close #1
contador = contador + 1
End Sub
Private Sub cmdFormatar_Click()
On Error Resume Next
If MsgBox("Deseja realmente formatar o disco 'A:\' ?", vbYesNo, Me.Caption) = vbYes Then
Dim DriveLetter$, DriveNumber&, DriveType&
    Dim RetVal&, RetFromMsg%
    DriveLetter = UCase("A:\")
    DriveNumber = (Asc(DriveLetter) - 65)
    DriveType = GetDriveType(DriveLetter)
    If DriveType = 2 Then
        RetVal = SHFormatDrive(Me.hwnd, DriveNumber, 0&, 0&)
End If
End If
End Sub

Private Sub cmdGerar_Click()
On Error Resume Next
contador = 0
pbar picProgresso, 0, True

MsgBox "Coloque um disquete vazio no drive 'A:\' e pressione OK...", , Me.Caption
cmdGerar.Enabled = False
cmdFormatar.Enabled = False
If EspacoLivre("A:\") >= 1200 Then
Timer1.Enabled = True
Else
cmdGerar.Enabled = True
cmdFormatar.Enabled = True
If MsgBox("Espaço no disco A:\ é insufuciente!" & vbCrLf & vbCrLf & "Deseja formatá-lo?", vbCritical + vbYesNo, Me.Caption) = vbYes Then
Dim DriveLetter$, DriveNumber&, DriveType&
    Dim RetVal&, RetFromMsg%
    DriveLetter = UCase("A:\")
    DriveNumber = (Asc(DriveLetter) - 65)
    DriveType = GetDriveType(DriveLetter)
    If DriveType = 2 Then
        RetVal = SHFormatDrive(Me.hwnd, DriveNumber, 0&, 0&)
        End If
        picStatus.Cls
picStatus.Print "Clique em 'Gerar Disco'..."

End If
End If
End Sub


Private Sub Form_Load()
On Error Resume Next
pbar picProgresso, 0, True
picStatus.Print "Clique em 'Gerar Disco'..."
End Sub

Private Sub Label4_Click()
On Error Resume Next
    Dim lRet As Long
    lRet = ShellExecute(0, "open", "mailto:gabrielfalcao@hotmail.com", vbNullString, vbNullString, 1)
End Sub

Private Sub Label5_Click()
On Error Resume Next
    Dim lRet As Long
    lRet = ShellExecute(0, "open", "http://www.gabrielfalcao.i8.com", vbNullString, vbNullString, 1)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If contador = 0 Then des "CDREG.EX_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "CDREG.EX_"
pbar picProgresso, 5, True
If contador = 1 Then des "CDUPDATE.EX_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "CDUPDATE.EX_"
pbar picProgresso, 10, True
If contador = 2 Then des "CDUPDATE.HL_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "CDUPDATE.HL_"
pbar picProgresso, 15, True
If contador = 3 Then des "DDLOADER.BIN", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "DDLOADER.BIN"
pbar picProgresso, 20, True
If contador = 4 Then des "DISCWZRD.EX_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "DISCWZRD.EX_"
pbar picProgresso, 25, True
If contador = 5 Then des "DISCWZRD.HL_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "DISCWZRD.HL_"
pbar picProgresso, 30, True
If contador = 6 Then des "DM1.EXE", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "DM1.EXE"
pbar picProgresso, 35, True
If contador = 7 Then des "DM1.HLP", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "DM1.HLP"
pbar picProgresso, 40, True
If contador = 8 Then des "DM1.REC", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "DM1.REC"
pbar picProgresso, 45, True
If contador = 9 Then des "DM.EXE", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "DM.EXE"
pbar picProgresso, 50, True
If contador = 10 Then des "FILECOPY.EX_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "FILECOPY.EX_"
pbar picProgresso, 55, True
If contador = 11 Then des "FILECOPY.ICO", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "FILECOPY.ICO"
pbar picProgresso, 60, True
If contador = 12 Then des "FILECOPY.TX_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "FILECOPY.TX_"
pbar picProgresso, 65, True
If contador = 13 Then des "JUMPERS.DB_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "JUMPERS.DB_"
pbar picProgresso, 70, True
If contador = 14 Then des "LIC.TX_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "LIC.TX_"
pbar picProgresso, 75, True
If contador = 15 Then des "NOATADRV.TXT", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "NOATADRV.TXT"
pbar picProgresso, 80, True
If contador = 16 Then des "ODIKRNL.BIN", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "ODIKRNL.BIN"
pbar picProgresso, 81, True
If contador = 17 Then des "ONTRACKD.SYS", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "ONTRACKD.SYS"
pbar picProgresso, 82, True
If contador = 18 Then des "ONTRACKS.38_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "ONTRACKS.38_"
pbar picProgresso, 83, True
If contador = 19 Then des "README.TXT", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "README.TXT"
pbar picProgresso, 84, True
If contador = 20 Then des "SEG32BIT.386", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "SEG32BIT.386"
pbar picProgresso, 85, True
If contador = 21 Then des "SEG32BIT.TXT", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "SEG32BIT.TXT"
pbar picProgresso, 86, True
If contador = 22 Then des "SETUP.EXE", "A:\"
picStatus.Cls
pbar picProgresso, 87, True
picStatus.Print "Copiando: " & "SETUP.EXE"
If contador = 23 Then des "TXTSECTS.SR_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "TXTSECTS.SR_"
pbar picProgresso, 92, True
If contador = 24 Then des "UNINSTAL.EX_", "A:\"
picStatus.Cls
picStatus.Print "Copiando: " & "UNINSTAL.EX_"
pbar picProgresso, 93, True
If contador = 25 Then des "XBIOS.OVL", "A:\"
picStatus.Cls
MkDir "A:\X"
picStatus.Print "Copiando: " & "XBIOS.OVL"
pbar picProgresso, 96, True
If contador = 26 Then
picStatus.Cls
picStatus.Print "Disco criado com sucesso!"
pbar picProgresso, 100, True
cmdGerar.Enabled = True
cmdFormatar.Enabled = True
Timer1.Enabled = False
End If
End Sub
