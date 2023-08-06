VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Right Click Menu In Browser Control"
   ClientHeight    =   3945
   ClientLeft      =   2355
   ClientTop       =   2775
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6720
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   420
      TabIndex        =   0
      Text            =   "http://www.planet-source-code.com"
      Top             =   3480
      Width           =   4395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Navigate!"
      Default         =   -1  'True
      Height          =   375
      Left            =   4980
      TabIndex        =   1
      Top             =   3480
      Width           =   1515
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2955
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   6135
      ExtentX         =   10821
      ExtentY         =   5212
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu mnuBrowser 
      Caption         =   "Browser Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuNavigate 
         Caption         =   "&Navigate"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hw As Long

Private Sub Command1_Click()
    WebBrowser1.Navigate Text1.Text
    WebBrowser1.Visible = True
End Sub

Private Sub Form_Load()
    Dim h As Long, aClass As String, k As Long
    h = GetWindow(hwnd, GW_CHILD)
    aClass = Space$(128)
    Do While h
        k = GetClassName(h, aClass, 128)
        If Left$(aClass, k) = "Shell Embedding" Then hw = h: Exit Do
        h = GetWindow(h, GW_HWNDNEXT)
    Loop

    WebBrowser1.Navigate ""
    origWndProc = SetWindowLong(hw, GWL_WNDPROC, AddressOf AppWndProc)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetWindowLong hw, GWL_WNDPROC, origWndProc
End Sub

Private Sub mnuPrint_Click()
    MsgBox "Print!"
End Sub

Private Sub mnuNavigate_Click()
Command1_Click
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub
