VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "KeyDefine by D@v!NcY"
   ClientHeight    =   7740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8325
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":08CA
   ScaleHeight     =   7740
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1245
      Left            =   1950
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Form3.frx":CDC2C
      Top             =   2250
      Width           =   4500
   End
   Begin VB.CheckBox c3 
      Caption         =   "ndis.vxd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4095
      Width           =   1200
   End
   Begin VB.CheckBox c2 
      Caption         =   "vnetbios.vxd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4095
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recuperar!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2835
      TabIndex        =   4
      Top             =   4635
      Width           =   2880
   End
   Begin VB.CheckBox c1 
      Caption         =   "vnetsup.vxd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4095
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escolha o arquivo a recuperar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2910
      TabIndex        =   2
      Top             =   3780
      Width           =   2640
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3660
      TabIndex        =   1
      Top             =   5805
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WinBug Killer"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   3225
      TabIndex        =   0
      Top             =   1680
      Width           =   1950
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
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   585
      Index           =   6
      Left            =   885
      Top             =   1800
      Width           =   6450
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
      Left            =   2010
      Shape           =   3  'Circle
      Top             =   3330
      Width           =   855
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   1215
      Index           =   2
      Left            =   825
      Top             =   2370
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
      Height          =   7530
      Index           =   5
      Left            =   135
      Shape           =   2  'Oval
      Top             =   135
      Width           =   8070
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      Height          =   540
      Index           =   4
      Left            =   3030
      Top             =   3540
      Width           =   4545
   End
End
Attribute VB_Name = "Menu"
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
Dim vnetbios As String
Dim vnetsup As String
Dim ndis As String

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

Private Sub Command2_Click()
Me.Enabled = False
Form1.Show
Me.Visible = False
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.FontUnderline = False
Label2.FontBold = False
Label2.FontItalic = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
    'If the above line is modified or moved a second copy of it
    'may be added again if the form is later Modified by VBSFC.
End Sub


Private Sub btsai_Click()
Unload Me
End Sub

Private Sub Command1_Click()
vnetsup = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\vxd\vnetsup"
vnetbios = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\vxd\ndis"
ndis = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\vxd\vnetbios"
If c1.Value = 1 Then
DeleteKey vnetsup
End If
If c2.Value = 1 Then
DeleteKey vnetbios
End If
If c3.Value = 1 Then
DeleteKey ndis
End If
End Sub

Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)
    'If the above two lines are modified or moved a second copy of
    'them may be added again if the form is later Modified by VBSFC.
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

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontUnderline = True
Label2.FontBold = True
End Sub
