VERSION 5.00
Begin VB.UserControl Command 
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   PropertyPages   =   "botao.ctx":0000
   ScaleHeight     =   2370
   ScaleWidth      =   3435
   ToolboxBitmap   =   "botao.ctx":0014
   Begin VB.OptionButton Option1 
      Caption         =   "Botão Colorido"
      Height          =   2310
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   30
      Width           =   3390
   End
End
Attribute VB_Name = "Command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"

'Event Declarations:
Event Click() 'MappingInfo=Option1,Option1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=Option1,Option1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Option1,Option1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Option1,Option1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Option1,Option1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Option1,Option1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Option1,Option1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Option1,Option1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
'Public Sub about()
'Form1.Show
'End Sub
'Default Property Values:
'Const m_def_about = 0
'Property Variables:
'Dim m_about As Variant



Private Sub Option1_Click()
    RaiseEvent Click
If Option1.Value = True Then
Option1.Value = False
End If
End Sub

Private Sub UserControl_Initialize()
Option1.Width = UserControl.Width - 25
Option1.Height = UserControl.Height - 25
End Sub

Private Sub UserControl_Resize()
Option1.Width = UserControl.Width - 25
Option1.Height = UserControl.Height - 25
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Option1,Option1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Option1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Option1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Option1,Option1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Option1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Option1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Option1,Option1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Option1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Option1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Option1,Option1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Option1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Option1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Option1,Option1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Option1.Refresh
End Sub

Private Sub Option1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Option1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Option1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Option1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Option1,Option1,-1,Caption
Public Property Get Texto() As String
Attribute Texto.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Texto = Option1.Caption
End Property

Public Property Let Texto(ByVal New_Texto As String)
    Option1.Caption() = New_Texto
    PropertyChanged "Texto"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer

    Option1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Option1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Option1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Option1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Option1.Caption = PropBag.ReadProperty("Texto", "Botão Colorido")
    Set Picture = PropBag.ReadProperty("Figura", Nothing)
'    m_about = PropBag.ReadProperty("about", m_def_about)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    UserControl.PropertyPages(Index) = PropBag.ReadProperty("about" & Index, "0")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer

    Call PropBag.WriteProperty("BackColor", Option1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Option1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", Option1.Enabled, True)
    Call PropBag.WriteProperty("Font", Option1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Texto", Option1.Caption, "Botão Colorido")
    Call PropBag.WriteProperty("Figura", Picture, Nothing)
'    Call PropBag.WriteProperty("about", m_about, m_def_about)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("about" & Index, UserControl.PropertyPages(Index), "0")
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Option1,Option1,-1,Picture
Public Property Get Figura() As Picture
Attribute Figura.VB_Description = "Returns/sets a graphic to be displayed in a CommandButton, OptionButton or CheckBox control, if Style is set to 1."
    Set Figura = Option1.Picture
End Property

Public Property Set Figura(ByVal New_Figura As Picture)
    Set Option1.Picture = New_Figura
    PropertyChanged "Figura"
End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=20
''Public Function about() As Hyperlink
''Form1.Show
''End Function
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=UserControl,UserControl,-1,PropertyChanged
''Public Sub about(Optional ByVal PropertyName As Variant)
''    UserControl.PropertyChanged PropertyName
''    Form1.Show
''
''End Sub
''
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get about() As Variant
'    about = m_about
'End Property
'
'Public Property Let about(ByVal New_about As Variant)
'    m_about = New_about
'    PropertyChanged "about"
'End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_about = m_def_about
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PropertyPages
Public Property Get about(ByVal Index As Integer) As String
Attribute about.VB_Description = "Returns or sets a string from an array that is the name of a property page that is associated with the control represented by a UserControl class."
Attribute about.VB_ProcData.VB_Invoke_Property = "Sobre"
    Call pp
End Property

Public Property Let about(ByVal Index As Integer, ByVal New_about As String)
    UserControl.PropertyPages(Index) = New_about
    PropertyChanged "about"
End Property

