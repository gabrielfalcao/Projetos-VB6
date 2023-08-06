VERSION 5.00
Begin VB.UserControl Label3D 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   ScaleHeight     =   300
   ScaleWidth      =   720
   ToolboxBitmap   =   "Label3D.ctx":0000
   Begin VB.Label lblFace 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3D"
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   600
   End
   Begin VB.Label lblShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3D"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   600
   End
End
Attribute VB_Name = "Label3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFace,lblFace,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblFace.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblFace.Caption() = New_Caption
    lblShadow.Caption() = New_Caption
    PropertyChanged "Caption"
    Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFace,lblFace,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblFace.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblFace.Font = New_Font
    Set lblShadow.Font = New_Font
    PropertyChanged "Font"
    Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFace,lblFace,-1,ForeColor
Public Property Get FaceColor() As OLE_COLOR
Attribute FaceColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    FaceColor = lblFace.ForeColor
End Property

Public Property Let FaceColor(ByVal New_FaceColor As OLE_COLOR)
    lblFace.ForeColor() = New_FaceColor
    PropertyChanged "FaceColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblShadow,lblShadow,-1,ForeColor
Public Property Get ShadowColor() As OLE_COLOR
Attribute ShadowColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ShadowColor = lblShadow.ForeColor
End Property

Public Property Let ShadowColor(ByVal New_ShadowColor As OLE_COLOR)
    lblShadow.ForeColor() = New_ShadowColor
    PropertyChanged "ShadowColor"
End Property

Private Sub UserControl_Initialize()
    lblFace.Left = 1
    lblFace.Top = 1
    lblShadow.Left = lblFace.Left + (lblFace.FontSize * (lblFace.FontSize / 2))
    lblShadow.Top = lblFace.Top + (lblFace.FontSize * (lblFace.FontSize / 2))
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblFace.Caption = PropBag.ReadProperty("Caption", "Label3D")
    lblShadow.Caption = PropBag.ReadProperty("Caption", "Label3D")
    Set lblFace.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set lblShadow.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblFace.ForeColor = PropBag.ReadProperty("FaceColor", &H80000012)
    lblShadow.ForeColor = PropBag.ReadProperty("ShadowColor", &HFFFFFF)
    lblShadow.Visible = PropBag.ReadProperty("Shadow", True)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = lblFace.Height + 100
    UserControl.Width = lblFace.Width + 200
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", lblFace.Caption, "Label3D")
    Call PropBag.WriteProperty("Font", lblFace.Font, Ambient.Font)
    Call PropBag.WriteProperty("FaceColor", lblFace.ForeColor, &H80000012)
    Call PropBag.WriteProperty("ShadowColor", lblShadow.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("Shadow", lblShadow.Visible, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub AboutBox()
Attribute AboutBox.VB_Description = "Auther & verssion Information\r\n"
Attribute AboutBox.VB_UserMemId = -552
    Load frm3DLAbout
    frm3DLAbout.Show vbModal
End Sub


Public Property Get Shadow() As Boolean
    Shadow = lblShadow.Visible
End Property

Public Property Let Shadow(ByVal New_Value As Boolean)
    lblShadow.Visible() = New_Value
    PropertyChanged "Shadow"
End Property
