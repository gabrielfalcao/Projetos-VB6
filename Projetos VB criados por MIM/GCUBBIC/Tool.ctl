VERSION 5.00
Begin VB.UserControl Tool 
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   ScaleHeight     =   885
   ScaleWidth      =   885
   ToolboxBitmap   =   "Tool.ctx":0000
   Begin VB.Image Image1 
      Height          =   795
      Left            =   25
      Picture         =   "Tool.ctx":0312
      Stretch         =   -1  'True
      Top             =   60
      Width           =   795
   End
End
Attribute VB_Name = "Tool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
'Const m_def_Enabled = 0
'Property Variables:
'Dim m_Enabled As Boolean
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
'Event Declarations:
Event Click() 'MappingInfo=Image1,Image1,-1,Click
Event DblClick() 'MappingInfo=Image1,Image1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Image1,Image1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Image1,Image1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Image1,Image1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
'Event Click()
'Event DblClick()
'Event KeyDown(KeyCode As Integer, Shift As Integer)
'Event KeyPress(KeyAscii As Integer)
'Event KeyUp(KeyCode As Integer, Shift As Integer)
'Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Public Function ExecutarPrograma(pastaearquivo As String) As Long
    Call Shell(pastaearquivo, 1)
End Function
Public Function DesligarComputador() As Long
    Dim desligar As String
    desligar = "RUNDLL.EXE user.exe,exitwindows"
    Call Shell(desligar, 1)
End Function
Public Function CopiarArquivo(origem As String, destino As String) As Long
FileCopy origem, destino
End Function
Public Function DeletarArquivo(pastaearquivo As String) As Long
Kill pastaearquivo
End Function
Public Function CriarPasta(nome As String)
MkDir nome
End Function
Public Function TocarSomWave(pastaearquivo As String)
SoundFile$ = pastaearquivo
wFlags% = SND_ASYNC Or SND_NODEFAULT
x% = sndPlaySound(SoundFile$, wFlags%)
End Function

Private Sub UserControl_Initialize()
UserControl.Height = 900
UserControl.Width = 900
End Sub
Public Function ChamarUmaConexãoDialUp(nomedaconexão As String) As Long
Dim x
x = Shell("rundll32.exe rnaui.dll,RnaDial " & nomedaconexão, 1)
DoEvents
End Function
Private Sub UserControl_Resize()
UserControl.Height = 900
UserControl.Width = 900
End Sub
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,0
'Public Property Get Enabled() As Boolean
'    Enabled = m_Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    m_Enabled = New_Enabled
'    PropertyChanged "Enabled"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=5
'Public Sub Refresh()
'  UserControl.Refresh
'End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_Enabled = m_def_Enabled
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Image1.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
    Call PropBag.WriteProperty("Enabled", Image1.Enabled, True)
'    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14
'Public Function CriarChave() As Variant

    
'End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8
Public Function DefinirOValorDeUmaString(objeto As String, EndereçoDaString As String, NomeDaString As String) As Long
SetStringValue EndereçoDaString, NomeDaString, (objeto)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8
Public Function VerOValorDeUmaString(objeto As String, EndereçoDaString As String, NomeDaString As String) As Long
objeto = GetStringValue(EndereçoDaString, NomeDaString)
End Function
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8
Public Function CriarChave(NomeEEndereçoDaChave As String) As Long
CreateKey NomeEEndereçoDaChave
End Function
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8
'Public Function MostrarAsHoras() As Long
'Time
'End Function
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8
'Public Function MostrarAData() As Long
'Date
'End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Image1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Image1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Image1.Refresh
End Sub

Private Sub Image1_Click()
    RaiseEvent Click
End Sub

Private Sub Image1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

