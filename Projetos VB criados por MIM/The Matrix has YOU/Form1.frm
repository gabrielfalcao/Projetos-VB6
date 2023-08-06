VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2730
      Top             =   1410
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1995
      Top             =   2250
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The Matrix has YOU..."
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1470
      Left            =   510
      TabIndex        =   0
      Top             =   615
      Width           =   3225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Dim a As Integer
Const SM_CXSCREEN = 0
Const SM_CYSCREEN = 1
Const HWND_TOP = 0
Const SWP_SHOWWINDOW = &H40

'Wanna quit early?  Press any key and its all gone.
Private Sub Form_KeyPress(KeyAscii As Integer)
'Reincarnate the cursor
ShowCursor (True)
End
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then Unload Me
Dim cx As Long
Dim cy As Long
Dim RetVal As Long
' Determine if screen is already maximized.
If Me.WindowState = vbMaximized Then
   ' Set window to normal size
   Me.WindowState = vbNormal
End If         ' Get full screen width.
cx = GetSystemMetrics(SM_CXSCREEN)         ' Get full screen height.
cy = GetSystemMetrics(SM_CYSCREEN)
' Call API to set new size of window.
RetVal = SetWindowPos(Me.hwnd, HWND_TOP, 0, 0, cx, cy, SWP_SHOWWINDOW)

'Murder the cursor
ShowCursor (False)



End Sub





Private Sub Timer1_Timer()
'Make the time between each letter printed random.
Randomize
Timer1.Interval = Int((1 * Rnd) + 250)
'The text, this is editable to whatever you wish.
sMarquee = "Wake up NEO" & vbCrLf & vbCrLf & vbCrLf & "The Matrix has you"
'Take one more letter each time and put it on the screen
a = a + 1
Label1.Caption = Left$(sMarquee, a)
'If it is a few seconds after the whole text is printed,
If a > Len(sMarquee) + 3 Then
'Then reincarnate the cursor
'ShowCursor (True)
'and kill my app
Timer2.Enabled = True

End If
End Sub

Private Sub Timer2_Timer()
If Label1.Caption = "The Matrix has YOU...€" Then
Label1.Caption = "Wake up NEO" & vbCrLf & vbCrLf & vbCrLf & "The Matrix has you"
Else
Label1.Caption = "Wake up NEO" & vbCrLf & vbCrLf & vbCrLf & "The Matrix has you€"
End If
End Sub
