VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCaptionlessWindowMover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class       : CCaptionlessWindowMover
' Description : Allows the user to move a form even if it doesn't
'               have a caption
' Source      : Total VB SourceBook 6

' Declarations for Windows API calls
Private Declare Function ReleaseCapture _
  Lib "user32" () _
  As Long
  
Private Declare Function SetCapture _
  Lib "user32" _
  (ByVal hwnd As Long) _
  As Long

' Local variables to hold Public Property values
Private mfAllowFormMove As Boolean
Private msngMouseStartX As Single
Private msngMouseStartY As Single
Private m_frmForm As Form

Private Sub Class_Terminate()
  ' Release storage for local object variables
  ' Source: Total VB SourceBook 6

  On Error GoTo PROC_ERR
  
  Set m_frmForm = Nothing

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Class_Terminate"
  Resume PROC_EXIT

End Sub

Public Property Get Form() As Form
  ' Returns: a pointer to the form previously assigned
  ' Source: Total VB SourceBook 6
  On Error GoTo PROC_ERR
  
  Set Form = m_frmForm

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Form"
  Resume PROC_EXIT

End Property


Public Property Set Form(frmValue As Form)
  ' frmValue: The form to assign to this class
  ' Source: Total VB SourceBook 6

  On Error GoTo PROC_ERR
  
  Set m_frmForm = frmValue

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Form"
  Resume PROC_EXIT

End Property

Public Sub HandleMouseDown( _
  ByVal sngX As Single, _
  ByVal sngY As Single)
  ' Comments  : Handle the MouseDown event of the encapsulated form
  ' Parameters: sngX - value of the X position of the mouse when
  '             the form receives the MouseDown event
  '             sngY - value of the Y position of the mouse when
  '             the form receives the MouseDown event
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR
  
  ' Capture all mouse events while the mouse button is pressed
  SetCapture m_frmForm.hwnd
  
  ' Set flag indicating that the form is being moved
  mfAllowFormMove = True
  
  ' Save current position for offsetting in the MouseMove event
  msngMouseStartX = sngX
  msngMouseStartY = sngY
  
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "HandleMouseDown"
  Resume PROC_EXIT
  
End Sub

Public Sub HandleMouseMove(ByVal sngX As Single, ByVal sngY As Single)
  ' Comments  : Handle the MouseMove event of the encapsulated form
  ' Parameters: sngX - value of the X position of the mouse when
  '             the form receives the MouseMove event
  '             sngY - value of the Y position of the mouse when
  '             the form receives the MouseMove event
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  ' If the form is currently being moved because the mouse was
  ' clicked on it, move the form in response to the MouseMove
  ' event.
  If mfAllowFormMove = True Then
    m_frmForm.Move m_frmForm.Left - msngMouseStartX + sngX, _
      m_frmForm.Top - msngMouseStartY + sngY
  End If

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "HandleMouseMove"
  Resume PROC_EXIT
End Sub

Sub HandleMouseUp()
  ' Comments  : Handle the MouseUp event of the encapsulated form
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  ' Stop moving the form
  ReleaseCapture

  ' Reset flag to stop handling MouseMove events
  mfAllowFormMove = False

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "HandleMouseUp"
  Resume PROC_EXIT

End Sub




