VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boot Creator Gold"
   ClientHeight    =   2820
   ClientLeft      =   2445
   ClientTop       =   4500
   ClientWidth     =   4680
   Icon            =   "criador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4680
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   3675
      Top             =   1500
   End
   Begin VB.Timer Timercopy 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   435
      Top             =   2130
   End
   Begin BootCreator.ProgressBar pb 
      Height          =   330
      Left            =   15
      TabIndex        =   10
      Top             =   2475
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   582
      BackColor       =   0
      ForeColor       =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sobre"
      Height          =   315
      Left            =   3930
      TabIndex        =   9
      Top             =   1725
      Width           =   570
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Formatar Disco"
      Height          =   555
      Left            =   998
      TabIndex        =   5
      Top             =   1005
      Width           =   2685
   End
   Begin VB.OptionButton b 
      Caption         =   "B:\"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3105
      TabIndex        =   4
      Top             =   495
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Escolha o nome do seu drive de discos 1.44 MB 3½ :"
      Height          =   855
      Left            =   180
      TabIndex        =   2
      Top             =   75
      Width           =   4320
      Begin VB.OptionButton a 
         Caption         =   "A:\"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   375
         Value           =   -1  'True
         Width           =   870
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Criar"
      Height          =   555
      Left            =   998
      TabIndex        =   1
      Top             =   1560
      Width           =   2685
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2340
      TabIndex        =   6
      Top             =   1470
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1965
      TabIndex        =   7
      Top             =   1380
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   825
      TabIndex        =   8
      Text            =   "C:\DCBOOT\"
      Top             =   2865
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Hidden          =   -1  'True
      Left            =   1935
      ReadOnly        =   0   'False
      System          =   -1  'True
      TabIndex        =   11
      Top             =   1215
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRONTO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   2130
      Width           =   4680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwnd As Long, ByVal Drive As Long, ByVal fmtID As Long, ByVal options As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Dim da As String
Dim c As String
Dim barra As String
Dim acpb As String
Dim drv As String
Private Function Pb1()
If pb.Value < pb.Max Then
pb.Value = pb.Value + 1
End If
End Function
Private Sub a_Click()
da = a.Caption
b.Value = False
End Sub

Private Sub b_Click()
da = b.Caption
a.Value = False
End Sub

Private Sub Command2_Click()

Dim DriveLetter$, DriveNumber&, DriveType&
    Dim RetVal&, RetFromMsg%
    DriveLetter = UCase(da)
    DriveNumber = (Asc(DriveLetter) - 65) ' Change letter to Number: A=0
    DriveType = GetDriveType(DriveLetter)
    If DriveType = 2 Then  'Floppies, etc
        RetVal = SHFormatDrive(Me.hwnd, DriveNumber, 0&, 0&)
    Else
        RetFromMsg = MsgBox("This drive is NOT a removeable" & vbCrLf & _
            "drive! Format this drive?", 276, "SHFormatDrive Example")
        Select Case RetFromMsg
            Case 6   'Yes
                ' UnComment to do it...
                'RetVal = SHFormatDrive(Me.hwnd, DriveNumber, 0&, 0&)
            Case 7   'No
                ' Do nothing
        End Select
    End If
    
End Sub

Private Sub Command3_Click()
frmSplash.Show
End Sub

Private Sub Form_Load()

If a.Value = True Then
Text2.Text = a.Caption
Else
Text2.Text = b.Caption
End If
da = Text2.Text
Dim pasta As String
If Len(App.Path) > 3 Then
File1.Path = App.Path & "\W32"
pasta = App.Path & "\W32\"
Else
File1.Path = App.Path & "W32\"
pasta = App.Path & "W32\"
End If
pb.Max = File1.ListCount
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen
End Sub

Private Sub Command1_Click()
Label1.Caption = "Preparando"
barra = pb.Value
Timercopy.Enabled = True
Command1.Enabled = False
End Sub

Private Sub Label2_DblClick()
Dim intResp As String

'exibe a inputBox
intResp = InputBox("Digite o diretório(com barra no final)...", "Escolha o diretório de destino:", "C:\Boot\", 2000, 1000)
'tratando o valor retornado pelo usuario

da = intResp
End Sub

Private Sub Timer1_Timer()

Label1.Caption = "PRONTO"
Timer1.Enabled = False
Timercopy.Enabled = False

End Sub

Private Sub Timercopy_Timer()
On Error Resume Next
If Not File1.ListIndex = File1.ListCount Then
If Len(File1.Path) > 3 Then
Label1.Caption = "Copiando arquivo: " & File1.FileName
Pb1
FileCopy File1.Path & "\" & File1.FileName, da & File1.FileName
File1.ListIndex = File1.ListIndex + 1
Else
Label1.Caption = "Copiando arquivo: " & File1.FileName
Pb1
FileCopy File1.Path & File1.FileName, da & File1.FileName
File1.ListIndex = File1.ListIndex + 1
End If

If pb.Value = pb.Max Then
pb.Value = pb.Max
Label1.Caption = "Finalizando instalação..."
pb.Value = pb.Min
MsgBox "Disco de BOOT criado com sucesso!", vbInformation, Me.Caption
Beep
Beep
Beep
Beep
Label1.Caption = "<<--== PRONTO ==-->>"
Command1.Enabled = True
Timercopy.Enabled = False
End If
End If

End Sub
