VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Hackeador 1.6 FREEWARE"
   ClientHeight    =   5625
   ClientLeft      =   1650
   ClientTop       =   2205
   ClientWidth     =   8055
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8055
   Begin VB.CheckBox Check1 
      Caption         =   "&Janelas e Nomes"
      Height          =   195
      Left            =   1920
      TabIndex        =   7
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hackeá-lo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Change text of any window :)"
      Top             =   4800
      Width           =   1230
   End
   Begin MSComctlLib.ListView View2 
      Height          =   4215
      Left            =   4080
      TabIndex        =   4
      ToolTipText     =   "Child windows"
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7435
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView View1 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Parent windows"
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7435
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&numerar!"
      Default         =   -1  'True
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Sair"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Criado por 7£r@70$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5250
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "+Programas Hacker"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4920
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Clique com os botões Esquerdo e/ou Direito nas Handles e Veja o que acontece"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6645
   End
   Begin VB.Menu Options 
      Caption         =   "Opções"
      Begin VB.Menu Show 
         Caption         =   "&Mostrar Janela Usando ShowWindow API"
      End
      Begin VB.Menu Show_BWTT 
         Caption         =   "Mostrar Janela Usando BringWindowToTop API"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu Max 
         Caption         =   "Ma&ximizar"
      End
      Begin VB.Menu Min 
         Caption         =   "Mi&nimizar"
      End
      Begin VB.Menu Restore 
         Caption         =   "&Restaurar"
      End
      Begin VB.Menu Hide 
         Caption         =   "&Esconder"
      End
      Begin VB.Menu Close 
         Caption         =   "&Fechar esta janela"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu SpyMenu 
         Caption         =   "Espionar os &Menus"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "menu2"
      Visible         =   0   'False
      Begin VB.Menu BnClick 
         Caption         =   "&Clicar"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------
'           Author: Muhammad Abubakar
'           http://go.to/abubakar
'           <joehacker@yahoo.com>
'------------------------------------
Option Explicit
Private ClassResize As New CResize

'API to open the browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
    String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub BnClick_Click()
    SendMessage Val(View2.SelectedItem), BM_CLICK, 0, 0
    
End Sub

Private Sub Close_Click()
    'close window code goes here:
    Dim lhwnd As Long
    
    On Error Resume Next
    lhwnd = Val(View1.SelectedItem)
    SendMessage lhwnd, WM_CLOSE, 0, 0

End Sub

Private Sub Command1_Click()
    'Free the memory occupied by the Object
    Set ClassResize = Nothing
    Unload Me

End Sub
Private Sub Command2_Click()
    Command2.Caption = "&Refresh"
    View1.ListItems.Clear
    View2.ListItems.Clear
    View1.GridLines = True
    Dim myLong As Long
    VCount = 1
    myLong = EnumWindows(AddressOf WndEnumProc, View1)

End Sub

Private Sub Command3_Click()
    Form2.Show vbModal
    
End Sub

Private Sub Form_Load()
    With ClassResize
        .hParam = Form1.Height
        .wParam = Form1.Width
        .Map Command1, RS_Top_Left
        .Map Command2, RS_Top_Left
        .Map Command3, RS_Top_Left
        .Map Label2, RS_TopOnly
        .Map Label3, RS_LeftOnly
        .Map View1, RS_HeightOnly
        .Map View2, RS_HeightOnly
        .Map Check1, RS_Top_Left
    End With
    Form1.Width = 11000
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    View1.View = lvwReport
    With View1.ColumnHeaders
        .Add , , "Handle", 1000
        .Add , , "Nome da Classe", 1500
        .Add , , "Texto", 4500
    End With
    VCount = 1
    View2.View = lvwReport
    With View2.ColumnHeaders
        .Add , , "Handle", 1000
        .Add , , "Nome da Classe", 1500
        .Add , , "Texto", 4500
        .Add , , "É Campo de Senha?", 1000
        
    End With
    ICount = 1
    Options.Visible = False
End Sub

Private Sub Form_Resize()
    ClassResize.rSize Form1
    
    'OK now resize if you must!
     View2.Left = Int(Form1.Width / 2)
     View1.Width = View2.Left - 255
     View2.Width = Int(Form1.Width / 2) - 255
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
End Sub

Private Sub Hide_Click()
    ShowWindow Val(View1.SelectedItem), SW_HIDE
End Sub

Private Sub Label2_Click()
    Dim ret As Long
    ret = ShellExecute(Me.hwnd, "Open", "http://tebugho.i8.com", "", App.Path, 1)

End Sub



Private Sub Max_Click()
    ShowWindow Val(View1.SelectedItem), SW_MAXIMIZE
    
End Sub

Private Sub Min_Click()
    ShowWindow Val(View1.SelectedItem), SW_MINIMIZE
End Sub

Private Sub Restore_Click()
    ShowWindow Val(View1.SelectedItem), SW_RESTORE
End Sub

Private Sub Show_BWTT_Click()
    Dim lhwnd As Long
    
    On Error GoTo bugging
    lhwnd = Val(View1.SelectedItem)
    'ShowWindow lhwnd, SW_SHOW
    BringWindowToTop lhwnd
    
    Exit Sub
bugging:
    Rem Do Nothing
    
End Sub

Private Sub Show_Click()
    'show window code goes here:
    Dim lhwnd As Long
    On Error Resume Next

    lhwnd = Val(View1.SelectedItem)
    ShowWindow lhwnd, SW_SHOW
End Sub

Private Sub SpyMenu_Click()
    Dim st As RECT
    
    Spy_Form.Show
    SpyHwnd = Val(View1.SelectedItem)
    Spy_Form.Tree.Nodes.Clear
    'If its a MDI type window and its child windows are maximized
    'then 'GetMenuItemInfo' crashes the 'EnumerationX'.
    'I tried to cascade the windows of other app but that doesnt
    'happen, do you know how I can do this?
    'MsgBox CascadeWindows(SpyHwnd, MDITILE_SKIPDISABLED, st, 0, 0)
    'SendMessage SpyHwnd, WM_MDICASCADE, MDITILE_SKIPDISABLED, 0
    'SendMessage SpyHwnd, WM_MDITILE, MDITILE_HORIZONTAL, 0
    
    SMenu GetMenu(SpyHwnd), Spy_Form.Tree
        
End Sub

Private Sub View1_Click()
    GotoChild
End Sub

Private Sub View1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then GotoChild
                'So that you are able to see child windows easily by
                'scrolling through up-down arrow keys instead of
                'clicking the parent window handle every time.
    
End Sub
Private Sub GotoChild()
    On Error GoTo HandleErrorPlz
    
    Dim Num As Long
    Dim myLong As Long
    Num = Val(View1.SelectedItem)
    View2.ListItems.Clear
    View2.GridLines = True
    ICount = 1
    myLong = EnumChildWindows(Num, AddressOf WndEnumChildProc, View2)

HandleErrorPlz:
    'Exit Sub ' As simple as that :)
End Sub

Private Sub View1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And View1.ListItems.Count > 0 Then
        If GetMenu(Val(View1.SelectedItem)) > 0 Then
            SpyMenu.Enabled = True
        Else
            SpyMenu.Enabled = False
        End If
               
        PopupMenu Options
    End If
    
End Sub
Private Sub View2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And View2.ListItems.Count > 0 Then
        PopupMenu menu2
    End If

End Sub
