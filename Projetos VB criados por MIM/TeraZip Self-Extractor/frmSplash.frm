VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00EFD1AD&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2130
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFD1AD&
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3915
      Begin VB.Timer Timer1 
         Interval        =   4000
         Left            =   2805
         Top             =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "gabrielfalcao@hotmail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00990B0B&
         Height          =   165
         Left            =   2055
         TabIndex        =   6
         Top             =   945
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desenvolvido para eliminar a necessidade de uso de programas pesados e igualmente funcionais"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00990B0B&
         Height          =   360
         Left            =   390
         TabIndex        =   4
         Top             =   1215
         Width           =   3375
      End
      Begin VB.Image imgLogo 
         Height          =   720
         Left            =   210
         Picture         =   "frmSplash.frx":000C
         Top             =   375
         Width           =   720
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00990B0B&
         Height          =   195
         Left            =   1635
         TabIndex        =   1
         Top             =   540
         Width           =   330
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TeraZip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00990B0B&
         Height          =   240
         Left            =   960
         TabIndex        =   3
         Top             =   345
         Width           =   705
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "por Gabriel Falcão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00990B0B&
         Height          =   165
         Left            =   1680
         TabIndex        =   2
         Top             =   765
         Width           =   1125
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carregando programa..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00754709&
      Height          =   165
      Left            =   1117
      TabIndex        =   5
      Top             =   1905
      Width           =   1800
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Path As String
Dim p_ret As String
Dim dll1 As String
Dim dll2 As String
Dim dll3 As String
Dim dll4 As String
Dim dll5 As String
Dim dll6 As String
Dim dll7 As String
Dim ok As Boolean
Private Sub Form_Load()
ok = False
If Len(App.Path) = 3 Then
Path = App.Path
Else
Path = App.Path & "\"
End If
dll1 = Path & "MSVCRT.DLL"
p_ret = StrConv(LoadResData("MSVCRT.DLL", "NEED"), vbUnicode)
Open dll1 For Binary As #1
Put #1, , p_ret
Close #1
dll2 = Path & "SCRRUN.DLL"
p_ret = StrConv(LoadResData("SCRRUN.DLL", "NEED"), vbUnicode)
Open dll2 For Binary As #1
Put #1, , p_ret
Close #1
dll3 = Path & "ZIPIT.DLL"
p_ret = StrConv(LoadResData("ZIPIT.DLL", "NEED"), vbUnicode)
Open dll3 For Binary As #1
Put #1, , p_ret
Close #1
dll4 = Path & "WSCRIPT.EXE"
p_ret = StrConv(LoadResData("WSCRIPT.EXE", "NEED"), vbUnicode)
Open dll4 For Binary As #1
Put #1, , p_ret
Close #1
dll5 = Path & "ZIPDLL.DLL"
p_ret = StrConv(LoadResData("ZIPDLL.DLL", "NEED"), vbUnicode)
Open dll5 For Binary As #1
Put #1, , p_ret
Close #1
dll6 = Path & "UNZDLL.DLL"
p_ret = StrConv(LoadResData("UNZDLL.DLL", "NEED"), vbUnicode)
Open dll6 For Binary As #1
Put #1, , p_ret
Close #1
dll7 = Path & "ZIP32.DLL"
p_ret = StrConv(LoadResData("ZIP32.DLL", "NEED"), vbUnicode)
Open dll7 For Binary As #1
Put #1, , p_ret
Close #1
ok = True
End Sub
Private Sub Timer1_Timer()
If ok = True Then
FrmMenu.Show
Unload Me
End If
End Sub
