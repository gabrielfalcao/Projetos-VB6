VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Win32 API declaration
'GetLastError return the code number of the error when the systray access fail
Private Declare Function GetLastError Lib "kernel32" () As Long
'Shell_NotifyIcon is THE function to add, modify or delete an existing icon.
'If it return 'True', that mean that the call was successful.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'Arguments of 'Shell_NotifyIcon':
                  'dwMessage: double word message. This represent the action
                  'to execute on the Systray. The message can be one of
                  'the following constants:
Private Const NIM_ADD = &H0 'Add a new icon to the Systray
Private Const NIM_MODIFY = &H1 'Modify an existing icon
Private Const NIM_DELETE = &H2 'Delete an existing icon
                  
                  'pnid: pointer to a NOTIFYICONDATA type (see below)
                  'This represent the info on the icon.

'Constants for the member 'uFlag' of the NOTIFYICONDATA type
'uFlag can be a combinaison of these three constants
Private Const NIF_MESSAGE = &H1 'Tell that the message has been updated
Private Const NIF_ICON = &H2 'Tell that the icon picture has been changed
Private Const NIF_TIP = &H4 'Tell that a new ToopTip for the icon is set

Private Const WM_MOUSEMOVE = &H200 'Used as the ID of the callback message

Private Const MAX_TIP_LENGTH As Long = 64 'This is the max length
                  'of a ToolTip. This value has been tested for Win95.
                  'For Win98 and NT, try changing this value. Tell me
                  'if it is a different value.

Private Type NOTIFYICONDATA
    cbSize As Long 'The size of this type
    hWnd As Long 'The hWnd that will receive the CallBack message
    uId As Long 'The ID of the application. Zero represent this application
    uFlags As Long 'The flags. Look at the constants beginning by NIF_ for the flags and their definition
    uCallbackMessage As Long 'This is the callback message
    hIcon As Long 'An handle to the icon that will be displayed
    szTip As String * MAX_TIP_LENGTH 'The string of the ToopTip. Must be terminating by a null zero (chr(0))
End Type

'The variable that will be used to interfere with the Systray
Private nidTrayIcon As NOTIFYICONDATA

Private bIconDisplayed As Boolean 'The status of the icon. True=Displayed
Private bUpdateOnChange As Boolean 'If True, Shell_NotifyIcon is call
                                  'whenever a change to a properties
                                  'of nidTrayIcon is made
                                  'True when class is initialized.

Public Event NIError(ByVal ErrorNumber As Long)

'Now for the Class Members
Public Function Initialize(ByVal hWnd As Long, ByVal hIcon As Long, ByVal sTip As String, Optional ByVal uCallbackMessage As Long = WM_MOUSEMOVE) As Long
  'Initialize the systray icon (The icon won't show)
  'THIS FUNCTION MUST BE THE FIRST MEMBER TO BE CALL AFTER
  'CREATING A INSTANCE OF THIS CLASS
  '
  'Input:   hWnd: Handle of the window that receives notification
  '               messages associated with an icon in the taskbar
  '               status area.
  '         hIcon: Handle of the icon to add, modify, or delete.
  '                This is not directly the bitmap. It is his handle
  '                in memory. So instead of using Picture1.Picture,
  '                you can use LoadPicture() or LoadResPicture() if
  '                you don't want to use a Picturebox.
  '
  '         szTip: Tooltip text to display for the icon.
  '                Max length: MAX_TIP_LENGTH - 1
  '                The class will add the null zero needed.
  '                The class will truncate the string if too long.
  '
  '         uCallbackMessage: Application-defined message
  '               identifier. The system uses the specified
  '               identifier for notification messages that it
  '               sends to the window identified by hWnd
  '               whenever a mouse event occurs in the bounding
  '               rectangle of the icon.
  '               Default: &H200 (512, WM_MOUSEMOVE)
  With nidTrayIcon
   .cbSize = Len(nidTrayIcon)
   .hIcon = hIcon
   .hWnd = hWnd
   .szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar
   .uCallbackMessage = uCallbackMessage
   .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
   .uId = vbNull
  End With
  bIconDisplayed = False
  bUpdateOnChange = True
End Function

Public Function ShowIcon() As Boolean
  'Show the icon specified by hIcon in the systray.
  If Not bIconDisplayed Then 'If the icon is not already displayed...
    ShowIcon = Shell_NotifyIcon(NIM_ADD, nidTrayIcon)
    If ShowIcon = False Then 'If there was an error
      RaiseEvent NIError(GetLastError) 'Return the error number
    Else
      bIconDisplayed = True 'The icon is displayed
    End If
  End If
End Function

Public Function HideIcon() As Boolean
  'Remove the icon from the systray.
  If bIconDisplayed Then 'If the icon is displayed...
    HideIcon = Shell_NotifyIcon(NIM_DELETE, nidTrayIcon)
    If HideIcon = False Then 'If there was an error
      RaiseEvent NIError(GetLastError) 'Return the error number
    Else
      bIconDisplayed = False 'The icon is not displayed
    End If
  End If
End Function

Public Property Let IconHandle(ByVal hIcon As Long)
  'Change the icon displayed in the systray for the icon handled
  'by hIcon. bIconStatus must be true.
  nidTrayIcon.hIcon = hIcon
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_ICON
    Update 'Make the icon change appear
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property

Public Property Let TipText(ByVal sTip As String)
  'Change the tooltip text of the icon by sTip
  'Remove the last character of it length is MAX_TIP_LENGTH because
  'we need to include a Null Zero
  nidTrayIcon.szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_TIP
    Update
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property

Public Property Let CallbackMessage(ByVal uCallbackMessage As Long)
  'Change the callback message that is send to the form.
  'The new message is uCallbackMessage.
  nidTrayIcon.uCallbackMessage = uCallbackMessage
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_MESSAGE
    Update
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property

Public Function Update() As Boolean
  'Update the icon on the systray. Call this function when you
  'want to refresh the icon or when you have made a change to
  'a property when bUpdateOnChange is False.
  'Also called internaly when bUpdateOnChange is true.
  If bIconDisplayed Then 'The icon must be showed to make change
    Update = Shell_NotifyIcon(NIM_MODIFY, nidTrayIcon)
    If Update = False Then 'If there was an error
      RaiseEvent NIError(GetLastError) 'Return the error number
    End If
  End If
End Function

Public Property Get IconHandle() As Long
  'Return the handle of the current icon
  IconHandle = nidTrayIcon.hIcon
End Property

Public Property Get TipText() As String
  'Return the tooltip text w/o the ending null zero
  TipText = Left(nidTrayIcon.szTip, Len(nidTrayIcon.szTip) - 1)
End Property

Public Property Get CallbackMessage() As Long
  'Return the current callback message
  CallbackMessage = nidTrayIcon.uCallbackMessage
End Property

Public Property Let UpdateOnChange(bUpdate As Boolean)
  'Set the variable bUpdateOnChange to a new value. If bUpdate is
  'True, then when a change to a properties of nidTrayIcon is made,
  'Shell_NotifyIcon is call. Otherwise, the function Update must
  'be called to make the modification appear.
  bUpdateOnChange = bUpdate
End Property

Private Property Get UpdateOnChange() As Boolean
  'Return True if bUpdateOnChange is True
  UpdateOnChange = bUpdateOnChange
End Property

Private Sub Class_Terminate()
  'The destructor of the class. It remove the icon from the Systray
  HideIcon
End Sub

Public Property Get Visible() As Boolean
  'Return if the icon in the systray is visible
  If bIconDisplayed Then
    Visible = True
  End If
End Property

Public Property Let Visible(ByVal bVisible As Boolean)
  'This property does the same thing than calling ShowIcon() and
  'HideIcon()
  If bVisible Then
    ShowIcon
  Else
    HideIcon
  End If
End Property

