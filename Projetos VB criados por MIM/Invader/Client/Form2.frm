VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                          Remote Control"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "razajunaid@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   840
      MouseIcon       =   "Form2.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Created n Developed By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Muhammad Junaid Raza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   3240
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Waooo(frm As Form)
   
   Dim a As Integer
   Dim b As Integer
   Dim C As Integer
   Dim d As Integer
   Dim e As Integer
   Dim f As Integer
   Dim w As Integer
   Dim X As Integer
   Dim Y As Integer
   Dim z As Integer
   Dim current As Double
   Call frm.Move(0, 0)
   w = frm.Height: X = frm.Width: Y = frm.Top: z = frm.Left
   a = 0: b = 0: C = w: d = X: e = Y: f = z
   Do While a < frm.Height / 15 Or b < frm.Width / 15
      a = a + 25
      b = b + 25
      e = e + 40
      f = f + 148
      If a > frm.Height / 15 Then a = a - 23
      If b > frm.Width / 15 Then b = b - 23
      Call frm.Move(f, e, d, C)
      current = Timer
      Do While Timer - current < 0.01
         DoEvents
      Loop
      Call SetWindowRgn(frm.hwnd, CreateEllipticRgn(0, 0, b, a), True)
   Loop
   current = Timer
   Do While Timer - current < 1
      DoEvents
   Loop
  
End Sub

Private Sub Form_Activate()
Waooo Form2
End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Label3_Click()
ShellExecute hwnd, "open", "mailto:razajunaid@hotmail.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Timer1_Timer()
Label1.Top = Label1.Top - 50

If Label1.Top < 0 Then Label1.Top = 3240

End Sub
