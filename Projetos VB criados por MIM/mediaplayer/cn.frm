VERSION 5.00
Begin VB.Form cn 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   ClientHeight    =   6870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   ControlBox      =   0   'False
   Icon            =   "cn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   465
      Left            =   795
      ScaleHeight     =   405
      ScaleWidth      =   2205
      TabIndex        =   3
      Top             =   4995
      Width           =   2265
      Begin VB.CommandButton Command5 
         Caption         =   "Pause"
         Height          =   405
         Left            =   1470
         TabIndex        =   6
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Stop"
         Height          =   405
         Left            =   735
         TabIndex        =   5
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Play"
         Height          =   405
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0C0FF&
      Height          =   2235
      Left            =   555
      Pattern         =   "*.mpg;*.mpeg;*.avi;*.wmv;*.wma;*.mp3;*.wav;*.mid;*.dat"
      TabIndex        =   2
      Top             =   2250
      Width           =   2745
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0FFC0&
      Height          =   1440
      Left            =   555
      TabIndex        =   1
      Top             =   795
      Width           =   2745
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   555
      TabIndex        =   0
      Top             =   495
      Width           =   2745
   End
   Begin VB.Image Image2 
      Height          =   1185
      Left            =   270
      Picture         =   "cn.frx":08CA
      Top             =   5550
      Width           =   3285
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   1800
      Index           =   3
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   4875
      Width           =   3120
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H80000001&
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   1800
      Index           =   2
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   4875
      Width           =   3120
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H80000001&
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   4335
      Index           =   1
      Left            =   315
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   3120
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H80000001&
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   4335
      Index           =   0
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   315
      Width           =   3120
   End
End
Attribute VB_Name = "cn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file As String
Option Explicit
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwnd As Long, ByVal Drive As Long, ByVal fmtID As Long, ByVal options As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Dim OldFilePathName As String
Dim NewFilePathName As String
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Const RGN_COPY = 5
Private ResultRegion As Long

Private Const RectXRound As Integer = 28
Private Const RectYRound As Integer = 28

Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim Corraction As Integer
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    Dim i As Integer
    
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)
    
    For i = shpBorder.LBound To shpBorder.UBound
        Select Case shpBorder(i).Shape
            Case 0: 'rectangle & square
                ObjectRegion = CreateRectRgn( _
                        shpBorder(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
            Case 1: 'circle
                If shpBorder(i).Width > shpBorder(i).Height Then
                    Corraction = (shpBorder(i).Width - shpBorder(i).Height) / 2
                        
                    ObjectRegion = CreateRectRgn( _
                            (shpBorder(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left - Corraction + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                Else
                    Corraction = (shpBorder(i).Height - shpBorder(i).Width) / 2
                        
                    ObjectRegion = CreateRectRgn( _
                            (shpBorder(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top - Corraction + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                End If
            Case 4:  'round square
                ObjectRegion = CreateRoundRectRgn( _
                        shpBorder(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                        RectXRound, RectYRound)
            Case 5: 'round square
                If shpBorder(i).Width > shpBorder(i).Height Then
                    Corraction = (shpBorder(i).Width - shpBorder(i).Height) / 2
                        
                    ObjectRegion = CreateRoundRectRgn( _
                            (shpBorder(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left - Corraction + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                            RectXRound, RectYRound)
                Else
                    Corraction = (shpBorder(i).Height - shpBorder(i).Width) / 2
                        
                    ObjectRegion = CreateRoundRectRgn( _
                            (shpBorder(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top - Corraction + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                            RectXRound, RectYRound)
                End If
            Case 3: 'circle
                If shpBorder(i).Width > shpBorder(i).Height Then
                    Corraction = (shpBorder(i).Width - shpBorder(i).Height) / 2
                        
                    ObjectRegion = CreateEllipticRgn( _
                            (shpBorder(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left - Corraction + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                Else
                    Corraction = (shpBorder(i).Height - shpBorder(i).Width) / 2
                        
                    ObjectRegion = CreateEllipticRgn( _
                            (shpBorder(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top - Corraction + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                End If
            Case Else:  'oval
                shpBorder(i).Shape = 2
                ObjectRegion = CreateEllipticRgn( _
                        shpBorder(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
        End Select
        nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
        nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
        DeleteObject ObjectRegion
    Next i
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function
Private Sub Check1_Click()

End Sub

Private Sub Command3_Click()
If pl.mp1.URL = Empty Then
Else
pl.mp1.URL = file
pl.mp1.Controls.play
End If
End Sub

Private Sub Command4_Click()
If pl.mp1.URL = Empty Then
Else
pl.mp1.Controls.stop
End If
End Sub

Private Sub Command5_Click()
If pl.mp1.URL = Empty Then
Else
pl.mp1.Controls.pause
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
If Len(File1.Path) = 3 Then
file = File1.Path & File1.FileName
Else
file = File1.Path & "\" & File1.FileName
End If
End Sub

Private Sub File1_DblClick()

pl.mp1.URL = file
pl.mp1.Controls.play

End Sub

Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hwnd, CreateFormRegion(1, 1, 0, 0), True)
    'If the above two lines are modified or moved a second copy of
    'them may be added again if the form is later Modified by VBSFC.
pl.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload pl
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
    'If the above line is modified or moved a second copy of it
    'may be added again if the form is later Modified by VBSFC.
End Sub

