VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Imagem"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   LinkTopic       =   "Form2"
   ScaleHeight     =   5730
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   0
      Left            =   -1125
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   0
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   -690
      Pattern         =   "*.jpg;*.jpeg;*.gif;*.png;*.bmp;*.gab"
      TabIndex        =   0
      Top             =   1425
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clique com o botão direito para ver + opções"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   1035
      TabIndex        =   3
      Top             =   1440
      Width           =   6225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clique com o botão direito para ver + opções"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1050
      TabIndex        =   2
      Top             =   1455
      Width           =   6225
   End
   Begin VB.Image Image1 
      Height          =   5685
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8145
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clique com o botão direito para ver + opções"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   1020
      TabIndex        =   4
      Top             =   1440
      Width           =   6225
   End
   Begin VB.Menu imgMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu tamori 
         Caption         =   "Tamanho Original"
         Shortcut        =   {F2}
      End
      Begin VB.Menu ajust 
         Caption         =   "Ajustar à Tela"
         Shortcut        =   {F3}
      End
      Begin VB.Menu next 
         Caption         =   "Próxima Imagem"
         Shortcut        =   {F4}
      End
      Begin VB.Menu back 
         Caption         =   "Imagem Anterior"
         Shortcut        =   {F5}
      End
      Begin VB.Menu del 
         Caption         =   "Deletar Imagem"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu wall 
         Caption         =   "Definir como papel de parede"
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim ChangeWP
Dim source As String


Private Sub ajust_Click()
On Error Resume Next
Image1.Width = Me.Width - 120
Image1.Height = Me.Height - 170

Image1.Stretch = True
Me.Caption = File1.FileName
End Sub

Private Sub back_Click()
On Error Resume Next
If File1.ListIndex > 0 Then
File1.ListIndex = File1.ListIndex - 1
Else
File1.ListIndex = 0
End If
If Len(File1.Path) > 3 Then
Image1.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
Else
Image1.Picture = LoadPicture(File1.Path & File1.FileName)
End If
Me.Caption = File1.FileName
End Sub

Private Sub del_Click()
If MsgBox("Deseja realmente deletar o arquivo " & Chr(34) & File1.FileName & Chr(34) & " permanentemente?", vbYesNo, "E-Mager 1.0") = vbYes Then
If Len(File1.Path) > 3 Then
Kill File1.Path & "\" & File1.FileName
Else
Kill File1.Path & File1.FileName
End If
End If
File1.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
Image1.Left = 0
Image1.Top = 0
File1.ListIndex = 1
If Len(File1.Path) > 3 Then
Image1.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
Else
Image1.Picture = LoadPicture(File1.Path & File1.FileName)
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Form1.Show
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Image1.Stretch = True Then
Image1.Width = Me.Width - 120
Image1.Height = Me.Height - 170

End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
PopupMenu imgMenu
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
End If
End Sub

Private Sub next_Click()
On Error Resume Next
If File1.ListIndex < File1.ListCount - 1 Then
File1.ListIndex = File1.ListIndex + 1
Else
File1.ListIndex = 0
End If
If Len(File1.Path) > 3 Then
Image1.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
Else
Image1.Picture = LoadPicture(File1.Path & File1.FileName)
End If
Me.Caption = File1.FileName
End Sub

Private Sub tamori_Click()
On Error Resume Next
Image1.Stretch = False
Me.Caption = File1.FileName
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 39 Then
If File1.ListIndex < File1.ListCount - 1 Then
File1.ListIndex = File1.ListIndex + 1
Else
File1.ListIndex = 0
End If
If Len(File1.Path) > 3 Then
Image1.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
Else
Image1.Picture = LoadPicture(File1.Path & File1.FileName)
End If
End If
If KeyCode = 37 Then
If File1.ListIndex > 0 Then
File1.ListIndex = File1.ListIndex - 1
Else
File1.ListIndex = File1.ListCount - 1
End If
If Len(File1.Path) > 3 Then
Image1.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
Else
Image1.Picture = LoadPicture(File1.Path & File1.FileName)
End If
End If
If KeyCode = 8 Then
Unload Me
End If
If KeyCode = 38 Then
Image1.Width = Me.Width - 120
Image1.Height = Me.Height - 170

Image1.Stretch = True
End If
If KeyCode = 40 Then
Image1.Stretch = False
End If
Me.Caption = File1.FileName
End Sub

Private Sub wall_Click()
Dim wall As String
If MsgBox("Deseja definir a imagem " & Chr(34) & Me.Caption & Chr(34) & " como papel de parede?", vbYesNo, App.ProductName) = vbYes Then

wall = "C:\wallpaper.bmp"
SavePicture Image1.Picture, wall
ChangeWP = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, wall, 0)
End If
File1.Refresh
End Sub
