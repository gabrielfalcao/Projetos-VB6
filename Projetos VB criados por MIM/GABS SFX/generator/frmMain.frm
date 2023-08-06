VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TeraZip SFX Generator"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Gerar Executável"
      Height          =   495
      Left            =   765
      TabIndex        =   3
      Top             =   360
      Width           =   1545
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SALVAR"
      Height          =   315
      Left            =   4320
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abrir"
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
      Left            =   4320
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   810
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fD As cFileDialog
Dim fileContent As String
Dim fileContenta As String
Dim p_ret As String
Dim arqui As String
Dim filelen As String
Dim filen As String
Dim aa As Integer
Private Sub Command1_Click()
    On Error Resume Next
    Dim cD As New cFileDialog
    With cD
        .flags = OFN_FILEMUSTEXIST
        .hwnd = Me.hwnd
                .CancelError = False
                .Filter = "Arquivos EXE|*.exe"
     
        .ShowOpen
  
    End With

Open cD.Filename For Binary Access Read As #1
fileContent = Input(LOF(1), 1)
Close #1
MsgBox "Arquivo1: " & Len(fileContent)
fileContent = fileContent & "|"

    With cD
        .flags = OFN_FILEMUSTEXIST
        .hwnd = Me.hwnd
                .CancelError = False
                .Filter = "Arquivos EXE|*.exe"
     
        .ShowOpen
  
    End With

Open cD.Filename For Binary Access Read As #1
p_ret = Input(LOF(1), 1)
Close #1
MsgBox "Arquivo1: " & Len(p_ret)
fileContent = fileContent & p_ret
aa = Len(cD.Filename) - 4
filen = Left$(cD.Filename, aa) & ".exe"
'AddVariable "MsgBoxMessage", fileContent
'WriteExeFile "c:\SFX.exe", filen
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Dim cD As New cFileDialog
    With cD
        .flags = OFN_FILEMUSTEXIST
        .hwnd = Me.hwnd
                .CancelError = False
                .Filter = "Arquivos EXE|*.exe"
     
        .ShowSave
  
    End With

Open cD.Filename For Binary As #1
Put #1, , fileContent
Close #1
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Dim cD As New cFileDialog
    With cD
        .flags = OFN_FILEMUSTEXIST
        .hwnd = Me.hwnd
                .CancelError = False
                .Filter = "Arquivos EXE|*.exe"
     
        .ShowOpen
  
    End With
Dim aas As String
Open cD.Filename For Binary As #1
aas = Input(LOF(1), 1)
Close #1
'73729
Dim arq As String
arq = Right(aas, 73728)
With cD
        .flags = OFN_FILEMUSTEXIST
        .hwnd = Me.hwnd
                .CancelError = False
                .Filter = "Arquivos EXE|*.exe"
     .DialogTitle = "Salvar novo arquivo como"
        .ShowSave
  
    End With

Open cD.Filename For Binary As #2
Put #2, , arq
Close #2
End Sub

Private Sub Command4_Click()
 On Error Resume Next
    Dim cD As New cFileDialog
    With cD
        .flags = OFN_FILEMUSTEXIST
        .hwnd = Me.hwnd
                .CancelError = False
                .Filter = "Arquivos ZIP|*.ZIP"
     
        .ShowOpen
  
    End With

Open cD.Filename For Binary Access Read As #1
fileContent = Input(LOF(1), 1)
Close #1

    Open "C:\SFX.exe" For Binary Access Read As #1
fileContenta = Input(LOF(1), 1)
Close #1

aa = Len(cD.Filename) - 4
filen = Left$(cD.Filename, aa) & ".exe"
filelen = Len(fileContent)
    

arqui = fileContenta & "|" & fileContent

Open filen For Binary As #2
Put #2, , arqui
Close #2

MsgBox filelen
End Sub
