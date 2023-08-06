VERSION 5.00
Begin VB.Form frmEnab 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H1T - Super Cracker Habilitador"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   Icon            =   "frmEnab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   525
      Top             =   555
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   3870
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   1
      Top             =   90
      Width           =   915
      Begin VB.Image imgTarget 
         Height          =   480
         Left            =   180
         Picture         =   "frmEnab.frx":08CA
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "Habilitar/Desabilitar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   1890
   End
   Begin VB.Label lblClassMane 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   300
      Width           =   2145
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00691412&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEnab.frx":0BD4
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1785
      Left            =   90
      TabIndex        =   4
      Top             =   960
      Width           =   4770
   End
   Begin VB.Image imgCross 
      Height          =   480
      Left            =   345
      Picture         =   "frmEnab.frx":0D62
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgNull 
      Height          =   15
      Left            =   270
      Picture         =   "frmEnab.frx":106C
      Top             =   2055
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Handle da Janela:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   30
      TabIndex        =   3
      Top             =   45
      Width           =   1290
   End
   Begin VB.Label lblHWnd 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "0000000000"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1350
      TabIndex        =   2
      Top             =   45
      Width           =   900
   End
End
Attribute VB_Name = "frmEnab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/******************************************************************************
'Name: frmMain.frm (frmMain)
'
'Description: Main form for this project contains the use interface controls.
'
'Date Updated: 04/July/2003.
'
'Author: Peter Gransden.
'/******************************************************************************

'/******************************************************************************
Private Sub cmdEnable_Click()
On Error Resume Next
'/******************************************************************************
'Description: Button event to enable or disable a window/control,
'When you press the button it will send the EnableWindow API
'call to the to the windows handle, this
'
'Inputs: None
'
'Returns: None
'/******************************************************************************
    
    'This only works like a switch just sends ether an ON or Off,
    'It doesn't correspond to the current windows/controls state.
    If ControllEnabled = True Then
        cmdEnable.Caption = "Habilitar/Desabilitar"
        EnableWindow lblHWnd.Caption, 0 ' Disable < This is the API call
        ControllEnabled = False ' Sets swich
    Else
        cmdEnable.Caption = "Habilitar/Desabilitar"
        EnableWindow lblHWnd.Caption, 1 ' Enable < This is the API call
        ControllEnabled = True ' Sets Swich
    End If

End Sub


'/******************************************************************************
Private Sub imgTarget_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'/******************************************************************************
'Description: This changes the mouse cursor to +, and sets Targeting to True.
'
'Inputs: None
'
'Returns: None
'/******************************************************************************
    
    'We are Targeting
    Targeting = True
    'Blanks the target icon in the Picture Box
    imgTarget.Picture = imgNull.Picture
    'Sets the mouse cursor to Custom.
    Me.MousePointer = 99
    'Sets the mouse cursor to the +, in the picture box
    Me.MouseIcon = imgCross.Picture

End Sub

'/******************************************************************************
Private Sub imgTarget_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'/******************************************************************************
'Description: Gets all of the information about what's underneath the mouse cursor.
'
'Inputs: None
'
'Returns: None
'/******************************************************************************
Dim Child As Long ' Holds the Child's Hwnd
Dim WindowX As Long ' Holds then X posishin of the mouse cursor
Dim WindowY As Long ' Holds then Y posishin of the mouse cursor
Dim sName As String ' Holds the display name of the control
Dim sClassName As String * 255 ' Holds the class name returned by the GetClassName API
Dim TempHwnd As Long ' Holds the Hind of an object
    
    'If we aren't targeting then do nothing.
    If Targeting = False Then Exit Sub
        'Call to get the mouse position.
        Call GetCursorPos(CursorPosition)
        'Get the windows Handle from the cursors position
        TempHwnd = WindowFromPoint(CursorPosition.X, CursorPosition.Y)
        'Find the point on a window.
        GetWindowPoint TempHwnd, WindowX, WindowY
        'Find the Child object on a window (if any).
        Child = ChildWindowFromPoint(TempHwnd, WindowX, WindowY)
        
        'Ether use the child or windows Hwnd
        If Child = 0 Then
            'Get the class name of a window.
            Call GetClassName(TempHwnd, sClassName, 255)
            lblHWnd.Caption = TempHwnd
            ControllEnabled = True
        Else
            'Get the class name of the Child
            Call GetClassName(Child, sClassName, 255)
            lblHWnd.Caption = Child
            ControllEnabled = False
        End If
        
        'Format and display the information.
        sName = Trim(Left(sClassName, InStr(sClassName, vbNullChar) - 1))
        lblClassMane.Caption = sName

    
End Sub

'/******************************************************************************
Private Sub imgTarget_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'/******************************************************************************
'Description: This changes the mouse cursor back to default
'and sets Targeting to False.
'
'Inputs: None
'
'Returns: None
'/******************************************************************************
    
    'We are not Targeting anymore.
    Targeting = False
    'Sets the picture box back the + cursor
    imgTarget.Picture = imgCross.Picture
    'Sets the mouse cursor back to default.
    Me.MousePointer = 0

End Sub

Private Sub Timer1_Timer()
If cmdEnable.Enabled = False Then cmdEnable.Enabled = True

End Sub
