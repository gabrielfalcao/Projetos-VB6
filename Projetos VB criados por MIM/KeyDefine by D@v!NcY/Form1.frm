VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "KeyDefine by D@v!NcY"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   6270
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2490
      TabIndex        =   14
      Top             =   495
      Width           =   4710
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2490
      TabIndex        =   13
      Top             =   810
      Width           =   4710
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C8F1FD&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4395
      Width           =   4710
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C8F1FD&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4095
      Width           =   4710
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C8F1FD&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3795
      Width           =   4710
   End
   Begin VB.TextBox Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C8F1FD&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3495
      Width           =   4710
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C8F1FD&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3195
      Width           =   4710
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C8F1FD&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2895
      Width           =   4710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1815
      TabIndex        =   0
      Top             =   1875
      Width           =   4560
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.astalavista.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   4005
      TabIndex        =   19
      ToolTipText     =   "www.astalavista.com"
      Top             =   5790
      Width           =   2160
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site de Cracks:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2190
      TabIndex        =   18
      Top             =   5790
      Width           =   1590
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   825
      Index           =   8
      Left            =   3570
      Top             =   4815
      Width           =   1080
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KeyDefine by D@v!NcY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Left            =   2655
      TabIndex        =   17
      Top             =   120
      Width           =   2715
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7530
      MouseIcon       =   "Form1.frx":C6C0C
      Picture         =   "Form1.frx":C74D6
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7530
      MouseIcon       =   "Form1.frx":C7DA0
      Picture         =   "Form1.frx":C866A
      Top             =   3015
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome de Usuário:"
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
      Height          =   195
      Left            =   930
      TabIndex        =   16
      Top             =   555
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organização:"
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
      Height          =   195
      Left            =   1110
      TabIndex        =   15
      Top             =   870
      Width           =   1140
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   1080
      Index           =   0
      Left            =   4050
      Top             =   1875
      Width           =   135
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H80000007&
      FillColor       =   &H8000000F&
      Height          =   585
      Index           =   6
      Left            =   915
      Top             =   1800
      Width           =   6450
   End
   Begin VB.Image btsai 
      Height          =   300
      Left            =   6960
      Picture         =   "Form1.frx":C8F34
      Stretch         =   -1  'True
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID do seu Windows:"
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
      Height          =   195
      Left            =   840
      TabIndex        =   12
      Top             =   4470
      Width           =   1725
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versão do seu Windows:"
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
      Height          =   195
      Left            =   645
      TabIndex        =   5
      Top             =   4170
      Width           =   2115
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do seu Windows:"
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
      Height          =   195
      Left            =   690
      TabIndex        =   4
      Top             =   3855
      Width           =   2010
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome de Usuário atual:"
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
      Height          =   195
      Left            =   690
      TabIndex        =   3
      Top             =   2910
      Width           =   2010
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organização atual:"
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
      Height          =   195
      Left            =   885
      TabIndex        =   2
      Top             =   3225
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha do seu Windows:"
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
      Height          =   195
      Left            =   660
      TabIndex        =   1
      Top             =   3540
      Width           =   2070
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   390
      Index           =   7
      Left            =   6600
      Shape           =   1  'Square
      Top             =   1860
      Width           =   375
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   825
      Index           =   3
      Left            =   6285
      Shape           =   3  'Circle
      Top             =   3780
      Width           =   855
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   1215
      Index           =   2
      Left            =   885
      Top             =   75
      Width           =   6450
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   1170
      Index           =   1
      Left            =   3960
      Top             =   1200
      Width           =   300
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   2145
      Index           =   5
      Left            =   165
      Top             =   2775
      Width           =   7950
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   540
      Index           =   4
      Left            =   1890
      Top             =   5640
      Width           =   4545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   X As Long
   Y As Long
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

Private Sub btsai_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btsai.BorderStyle = 1
End Sub

Private Sub btsai_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btsai.BorderStyle = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
    'If the above line is modified or moved a second copy of it
    'may be added again if the form is later Modified by VBSFC.
End Sub


Private Sub btsai_Click()
Unload Me
Form3.Enabled = True
Form3.Visible = True
End Sub

Private Sub Command1_Click()
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner", (Text1.Text)
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization", (Text2.Text)
Label4.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "ProductKey")
Text4.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
Text3.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
Text6.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "Version")
Text5.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "ProductName")
Text7.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "ProductID")
Text4.Refresh
Text3.Refresh
Text5.Refresh
Text6.Refresh
Text7.Refresh
End Sub

Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)
    'If the above two lines are modified or moved a second copy of
    'them may be added again if the form is later Modified by VBSFC.
Text1.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
Text2.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
Label4.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "ProductKey")
Text4.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
Text3.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
Text6.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "Version")
Text5.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "ProductName")
Text7.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "ProductID")
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = False
Image2.Visible = True
CreateKey "HKEY_LOCAL_MACHINE\Software\KeyDefine"
SetStringValue "HKEY_LOCAL_MACHINE\Software\KeyDefine", "KeyDefineStatus", ("Foi Usado")
SetStringValue "HKEY_LOCAL_MACHINE\Software\KeyDefine", "Windows Status", ("Crackeado")

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = True
Image2.Visible = False
End Sub

Private Sub Label12_Click()
Clipboard.SetText "http://" & Label12.Caption
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 SendKeys "{TAB}"
 End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 SendKeys "{TAB}"
  SendKeys "{ENTER}"
 End If
End Sub

