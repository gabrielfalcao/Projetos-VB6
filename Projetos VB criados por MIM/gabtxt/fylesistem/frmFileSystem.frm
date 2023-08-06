VERSION 5.00
Begin VB.Form frmFileSystem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dr Ali Ezzahir            http://www.geocities.com/athens/aegean/6447       ezzahir@yahoo.com"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   1230
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9000
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   6840
      TabIndex        =   22
      Top             =   2640
      Width           =   2055
      Begin VB.CommandButton cmdAddExtension 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Picture         =   "frmFileSystem.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtAddExtension 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Text            =   "*.ini"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Add extension"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   6840
      TabIndex        =   19
      Top             =   0
      Width           =   2055
      Begin VB.TextBox txtDrive 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtExtensions 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "frmFileSystem.frx":00EA
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   6840
      TabIndex        =   15
      Top             =   3720
      Width           =   2055
      Begin VB.TextBox txtPath 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Text            =   "C:\"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton optOrdering 
         Caption         =   "Descending"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optOrdering 
         Caption         =   "Ascending"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Path to sorting files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Open Selected File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdProperties 
      BackColor       =   &H0000FF00&
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   5595
      TabIndex        =   9
      Top             =   1080
      Width           =   5655
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   5775
      End
   End
   Begin VB.TextBox txtBackPath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   885
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   120
      Width           =   5655
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   6360
      ScaleHeight     =   3075
      ScaleWidth      =   1635
      TabIndex        =   2
      Top             =   7000
      Width           =   1695
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   360
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Text            =   "C:\"
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "Sort"
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtExtensions2 
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Text            =   "*.*"
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdgo 
         Caption         =   "OK"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdSort2 
         Caption         =   "Sort2"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3600
      Width           =   5655
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1560
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Folders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   2400
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Picture         =   "frmFileSystem.frx":0113
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHObjectProperties Lib "Shell32" Alias "#178" (ByVal hOwner As Long, ByVal uFlags As Long, ByVal sName As String, ByVal sParam As String) As Long

Private Function GetAllDrives() As String
  Dim lRet As Long
  Dim temp As String
  temp = Space(64)
  lRet = GetLogicalDriveStrings(Len(temp), temp)
  GetAllDrives = Trim(temp)
End Function
Private Function StripNulls(sDriveList As String) As String
  Dim i As Integer
  Dim sDrive As String

  i = 1
  Do
    DoEvents
    If Mid$(sDriveList, i, 1) = Chr$(0) Then
      sDrive = Mid$(sDriveList, 1, i - 1)
      sDriveList = Mid$(sDriveList, i + 1, Len(sDriveList))
      StripNulls = sDrive
      Exit Function
    End If
    i = i + 1
  Loop
End Function

Private Sub cmdAddExtension_Click()
txtExtensions.Text = txtExtensions.Text & txtAddExtension.Text & vbCrLf
End Sub

Private Sub cmdgo_Click()
Text2.Text = ""
Text3.Text = ""
    findfilesdir Text1.Text, txtExtensions2.Text
Label1 = LineCount(Text2) - 1 & " file(s)"
Label2 = LineCount(Text3) - 1 & " folder(s)"
End Sub




Public Sub findfilesdir(DirPath As String, FileSpec As String)
Dim filestring As String

DirPath = Trim$(DirPath)

If Right$(DirPath, 1) <> "\" Then
  DirPath = DirPath & "\"
End If
On Error Resume Next
filestring = Dir$(DirPath & FileSpec, vbArchive Or vbHidden Or vbSystem Or vbDirectory)
Do
  DoEvents
  If filestring = "" Then
    Exit Do
  Else
    If (GetAttr(DirPath & filestring) And vbDirectory) = vbDirectory Then
      If Left$(filestring, 1) <> "." And Left$(filestring, 2) <> ".." Then
      Text3 = Text3 & filestring & vbCrLf
      
      End If
    Else
Text2 = Text2 & filestring & vbCrLf
    End If
  End If
  On Error GoTo errhandler
  filestring = Dir$
Loop
Exit Sub
errhandler:
End Sub

Private Sub cboExtensions_Click()
cmdgo_Click
End Sub




Private Sub cmdProperties_Click()
PopupMenu frmProperties.mnuProperties, 0, cmdProperties.Left, cmdProperties.Top + cmdProperties.Height
End Sub

Private Sub Command2_Click()
Text2.Text = ""
Text3 = ""
    findfilesdir Text1.Text, txtExtensions2.Text
End Sub

Private Sub Command3_Click()
Dim MyValue As String 'opens the chosen file
MyValue = Shell("rundll32.exe url.dll,FileProtocolHandler " & Text1.Text & LineText(Text2, CurrentLine(Text2)), 1)
End Sub

Private Sub Form_Load()
 Dim sDrives As String
  Dim curDrive As String
  Dim drvType As String
Picture1.Visible = False
Call sCenterForm(Me)
Label4.Caption = "Path to Sort.txt and" & vbCrLf
Label4 = Label4 & "Sortfiles.txt"
  sDrives = GetAllDrives()

  Do Until sDrives = Chr(0)
    DoEvents
    curDrive = StripNulls(sDrives)
  txtDrive.Text = txtDrive.Text & UCase(curDrive) & vbCrLf
  Loop
Text1.Text = "C:\"
Text2.Text = ""
Text3.Text = ""
    findfilesdir "C:\", "*.*"
Label1 = LineCount(Text2) - 1 & " file(s)"
Label2 = LineCount(Text3) - 1 & " folder(s)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Text1_Change()
txtBackPath.Text = txtBackPath.Text & Text1.Text & vbCrLf
End Sub



Private Sub Text2_Click()
Text5.Text = Text1.Text & LineText(Text2, CurrentLine(Text2))
End Sub

Private Sub Text3_Change()
Label1 = LineCount(Text2) - 1 & " file(s)"
Label2 = LineCount(Text3) - 1 & " folder(s)"
optOrdering(0).Value = False
optOrdering(1).Value = False
End Sub

Private Sub Text3_Click()
On Error GoTo errhandler
If LineText(Text3, CurrentLine(Text3)) = "" Or Text1.Text = "" Then Exit Sub
Text1.Text = Text1.Text & LineText(Text3, CurrentLine(Text3)) & "\"
Label1.Caption = ""
cmdgo_Click
Label1 = LineCount(Text2) - 1 & " file(s)"
Label2 = LineCount(Text3) - 1 & " folder(s)"
Command2_Click
Exit Sub
errhandler:
End Sub

Private Sub txtBackPath_Click()
Text2.Text = ""
Text3 = ""
optOrdering(0).Value = False
optOrdering(1).Value = False
txtExtensions2.Text = "*.*"
Text1.Text = LineText(txtBackPath, CurrentLine(txtBackPath)) ' lstdirs.Text
 If Text1.Text = "" Then Exit Sub
    findfilesdir LineText(txtBackPath, CurrentLine(txtBackPath)), txtExtensions2.Text
cmdgo_Click
End Sub

Private Sub txtDrive_Click()
Text2.Text = ""
Text3 = ""
optOrdering(0).Value = False
optOrdering(1).Value = False
txtExtensions2.Text = "*.*"
Text1.Text = LineText(txtDrive, CurrentLine(txtDrive)) ' lstdirs.Text
Text4.Text = LineText(txtDrive, CurrentLine(txtDrive))
If Text1.Text = "" Then Exit Sub
    findfilesdir LineText(txtDrive, CurrentLine(txtDrive)), txtExtensions2.Text
cmdgo_Click
End Sub
Private Sub cmdSort_Click()
Dim lines() As String
Dim new_line As String
Dim num_lines As Integer
Dim FileNum As Integer
Dim i As Integer
FileNum = FreeFile
Open txtPath.Text & "Sort.txt" For Output As FileNum
Print #FileNum, Text3.Text
Close FileNum
FileNum = FreeFile
On Error GoTo errhandler
    Open txtPath.Text & "Sort.txt" For Input As FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, new_line
        new_line = Trim$(new_line)

        num_lines = num_lines + 1
        ReDim Preserve lines(1 To num_lines)
        lines(num_lines) = new_line
    Loop
    Close FileNum
    SelectionSort lines, 1, num_lines
    Open txtPath.Text & "Sort.txt" For Output As FileNum
    If optOrdering(0).Value Then
         For i = 1 To num_lines
            Print #FileNum, lines(i)
        Next i
    Else
          For i = num_lines To 1 Step -1
            Print #FileNum, lines(i)
        Next i
    End If
    Close FileNum
   Open txtPath.Text & "Sort.txt" For Input As FileNum
   Text3.Text = Input(LOF(FileNum), FileNum)
   Close FileNum
    MousePointer = vbDefault
Kill txtPath.Text & "Sort.txt"
cmdSort2_Click
Exit Sub
errhandler:
End Sub

Private Sub SelectionSort(list() As String, ByVal min As Integer, ByVal max As Integer)
Dim i As Integer
Dim j As Integer
Dim best_j As Integer
Dim best_str As String
Dim temp_str As String

    For i = min To max - 1
        best_j = i
        best_str = list(i)
        For j = i + 1 To max
            If StrComp(list(j), best_str, vbTextCompare) < 0 Then
                best_str = list(j)
                best_j = j
            End If
        Next j
        list(best_j) = list(i)
        list(i) = best_str
    Next i
End Sub


Private Sub optOrdering_Click(Index As Integer)
cmdSort_Click
End Sub
Private Sub cmdSort2_Click()
Dim lines() As String
Dim new_line As String
Dim num_lines As Integer
Dim FileNum As Integer
Dim i As Integer

    MousePointer = vbHourglass
    DoEvents
FileNum = FreeFile
Open txtPath.Text & "Sortfiles.txt" For Output As FileNum
Print #FileNum, Text2.Text
Close FileNum
FileNum = FreeFile
    Open txtPath.Text & "Sortfiles.txt" For Input As FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, new_line
        new_line = Trim$(new_line)

        num_lines = num_lines + 1
        ReDim Preserve lines(1 To num_lines)
        lines(num_lines) = new_line
    Loop
    Close FileNum
    SelectionSort lines, 1, num_lines
    Open txtPath.Text & "Sortfiles.txt" For Output As FileNum
    If optOrdering(0).Value Then
         For i = 1 To num_lines
            Print #FileNum, lines(i)
        Next i
    Else
           For i = num_lines To 1 Step -1
            Print #FileNum, lines(i)
        Next i
    End If
    Close FileNum
   Open txtPath.Text & "Sortfiles.txt" For Input As FileNum
   Text2.Text = Input(LOF(FileNum), FileNum)
   Close FileNum
    MousePointer = vbDefault
Kill txtPath.Text & "Sortfiles.txt"
End Sub

Private Sub txtExtensions_Click()
optOrdering(0).Value = False
optOrdering(1).Value = False
txtExtensions2.Text = ""
txtExtensions2 = txtExtensions2.Text & LineText(txtExtensions, CurrentLine(txtExtensions))
cmdgo_Click
End Sub
