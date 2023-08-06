VERSION 5.00
Begin VB.Form frmCrack 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H1T - Password Cracker"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCrackPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTextBox 
      BackColor       =   &H00404040&
      Caption         =   "Notificar quando é o melhor momento para soltar o mouse"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2175
      TabIndex        =   13
      Top             =   1440
      Width           =   2910
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   5130
      ScaleHeight     =   1905
      ScaleWidth      =   1305
      TabIndex        =   5
      Top             =   1215
      Width           =   1335
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblYCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblXCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblTrack 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coordenadas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mire com isso:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   930
      End
      Begin VB.Image imgTarget 
         Height          =   480
         Left            =   360
         Picture         =   "frmCrackPass.frx":74F2
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.CheckBox chkOnTop 
      BackColor       =   &H00404040&
      Caption         =   "Manter o programa ""sempre por cima"""
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   2925
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   3450
      TabIndex        =   3
      Top             =   2145
      Width           =   1095
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2730
      Width           =   4650
   End
   Begin VB.Label lblRelease 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Solte o Botão!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   75
      TabIndex        =   12
      Top             =   1215
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image imgNull 
      Height          =   15
      Left            =   3420
      Picture         =   "frmCrackPass.frx":81BC
      Top             =   3855
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgCross 
      Height          =   480
      Left            =   2880
      Picture         =   "frmCrackPass.frx":8602
      Top             =   3855
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Resultado:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2490
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCrackPass.frx":92CC
      ForeColor       =   &H00FFFF00&
      Height          =   885
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   5775
   End
End
Attribute VB_Name = "frmCrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkOnTop_Click()

If chkOnTop.Value = 1 Then
               SetWindowPos frmCrack.hWnd, HWND_TOPMOST, frmCrack.Left / 15, _
                            frmCrack.Top / 15, frmCrack.Width / 15, _
                            frmCrack.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
Else
               SetWindowPos frmCrack.hWnd, HWND_NOTOPMOST, frmCrack.Left / 15, _
                            frmCrack.Top / 15, frmCrack.Width / 15, _
                            frmCrack.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End If

End Sub

Private Sub cmdAbout_Click()
frmAbout.Show 1

End Sub

Private Sub cmdExit_Click()

Unload Me

End Sub

Private Sub Form_Load()
chkTextBox.Value = 1

End Sub


Private Sub imgTarget_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
' Visual effects
    Targeting = True
    imgTarget.Picture = imgNull.Picture
    lblTrack.Visible = True
    lblXCap.Visible = True
    lblYCap.Visible = True
    lblX.Visible = True
    lblY.Visible = True
    
    Me.MousePointer = 99
    Me.MouseIcon = imgCross.Picture
    txtOutput.Text = ""
        
End Sub

Private Sub imgTarget_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim sName As String, sClassName As String * 255, TempHwnd As Long
    
    If Targeting = False Then Exit Sub
    Call GetCursorPos(CursorPosition)
    ' Display mouse's location
    lblX.Caption = CursorPosition.x
    lblY.Caption = CursorPosition.y

    ' Check whether the cursor is pointing to a TextBox or not
    TempHwnd = WindowFromPoint(CursorPosition.x, CursorPosition.y)
    Call GetClassName(TempHwnd, sClassName, 255)
    sName = Trim(Left(sClassName, InStr(sClassName, vbNullChar) - 1))
    
    If chkTextBox.Value = 1 Then
    
    If sName = "Edit" Or InStr(sName, "TextBox") > 0 Then
        lblRelease.Visible = True
    Else
        lblRelease.Visible = False
    End If
   End If
   

End Sub

Private Sub imgTarget_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim TargetLen As Long, TempString As String, hWnd As Long
' Visual effects
    Targeting = False
    lblTrack.Visible = False
    lblXCap.Visible = False
    lblYCap.Visible = False
    lblX.Visible = False
    lblY.Visible = False
    lblRelease.Visible = False
    
    imgTarget.Picture = imgCross.Picture
    Me.MousePointer = 0

'
Call GetCursorPos(CursorPosition)
    hWnd = WindowFromPoint(CursorPosition.x, CursorPosition.y) ' Get target window's handle
    hWnd = GetTopLevelParent(hWnd)          ' Get target window's parent's handle
    TargetLen& = SendMessage(hWnd&, WM_GETTEXTLENGTH, 0&, 0&)
    TempString$ = String(TargetLen&, 0&)
    Call sendmessagebystring(hWnd&, WM_GETTEXT, TargetLen& + 1, TempString$)
    txtOutput.Text = TempString$


End Sub
