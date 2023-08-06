VERSION 5.00
Begin VB.Form Help 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   6840
   ClientLeft      =   1455
   ClientTop       =   960
   ClientWidth     =   7785
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Bineries in: http://home.domaindlx.com/ribafs2/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   5055
   End
   Begin VB.Label lblUser 
      Caption         =   "  USER EDITOR: To Change User in Editor: Open File User.txt and alter user name."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5640
      Width           =   7335
   End
   Begin VB.Label lblPartial 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   7575
   End
   Begin VB.Label lblFull 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim strFull As String, strPartial As String
    
    strFull = "FULL PROJECT" & vbCrLf & vbCrLf & vbCrLf
    strFull = strFull & "1 - Change path of Project if prefer in 'Project Folder Path' TextBox" & vbCrLf
    strFull = strFull & "2 - In Menu click 'Project Type - Full'" & vbCrLf
    strFull = strFull & "3 - Select a MS Access Database" & vbCrLf
    strFull = strFull & "4 - Wait message 'Project created sucessuful!'" & vbCrLf
    strFull = strFull & "5 - If message 'Unable to create project' then verify Database." & vbCrLf
    strFull = strFull & "     Project don't yet accept Database with password" & vbCrLf & vbCrLf
    lblFull.Caption = strFull
    
    strPartial = "   PARTIAL PROJECT" & vbCrLf & vbCrLf
    strPartial = strPartial & "   1 - Change path of Project if prefer in 'Project Folder Path' TextBox" & vbCrLf
    strPartial = strPartial & "   2 - In menu click 'Project Type'" & vbCrLf
    strPartial = strPartial & "   3 - Click in Add, View, Update or Delete to generate page corresponding" & vbCrLf
    strPartial = strPartial & "   4 - Select a MS Access Database" & vbCrLf
    strPartial = strPartial & "   4 - Select a Table in ListBox only clicking" & vbCrLf
    strPartial = strPartial & "   5 - Wait message 'Project created sucessuful!'"
    lblPartial.Caption = strPartial
End Sub

Private Sub OKButton_Click()
    Unload Help
End Sub
