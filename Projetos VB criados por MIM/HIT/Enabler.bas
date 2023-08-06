Attribute VB_Name = "Enabler"
'/******************************************************************************
'Name: basMain.bas (basMain)
'
'Description: Contains all of the API Declarations, User defined type
'for this application.
'
'Date Updated: 04/July/2003.
'
'Author: Peter Gransden.
'/******************************************************************************

'API calls
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GetCursorPos
'Reads the current position of the mouse cursor.
'The x and y coordinates of the cursor (relative to the screen)
'are put into the variable passed as lpPoint.
'The function returns 0 if an error occured or 1 if it is successful.
'Public Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINT) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'WindowFromPoint
'Determines the handle of the window located at a specific
'point on the screen. Note that the active window could be a text box,
'list box, button, or some other object sitting inside a program window.
'In this case, the handle returned will be to this control and not the
'program window. If successful, the function returns the handle to the
'window at that point. If there is no window at that point,
'or if an error occurred, the function instead returns 0.
Public Declare Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GetClassName
'Retrieves the name of the window class to which a window belongs.
'The name of the class is placed into the string passed as lpClassName.
'If an error occurred, the function returns 0, If successful, the function
'returns the number of characters copied into the string passed as lpClassName.
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'EnableWindow
'enables or disables a window. If a window is disabled, it cannot
'receive the focus and will ignore any attempted input. Some types of windows,
'such as buttons and other controls, will appear grayed when disabled,
'although any window can be enabled or disabled.
'The function returns 0 if the window had previously been enabled,
'or a non-zero value if the window had been disabled.
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GetWindowRect
'Reads the size and position of a window. This information is put into the
'variable passed as lpRect. The rectangle receives the coordinates of the
'upper-left and lower-right corners of the window. If the window is past one
'of the edges of the screen, the values will reflect that
'(for example, if the left edge of a window is off the screen, the
'rectangle's .Left property will be negative). The function returns 0
'if an error occured, or 1 if successful.
'Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'IsWindow
'Determines if a given handle refers to a window or not. Handles can
'not only refer to windows but also to many objects such as fonts,
'brushes, bitmaps, and even registry keys. The function returns 0
'if the handle does not refer to a window, or a non-zero value if
'it does refer to a window.
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GetParent
'Returns the handle of the parent window of another window. For example,
'the parent of a button would normally be the form window it is in.
'If successful, the function returns a handle to the parent window.
'If it fails (for example, trying to find the parent of a non-window),
'it returns 0.
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'WindowFromPoint
'Determines the handle of the window located at a specific point on the screen.
'Note that the active window could be a text box, list box, button, or some
'other object sitting inside a program window. In this case, the handle
'returned will be to this control and not the program window. If successful,
'the function returns the handle to the window at that point. If there is no
'window at that point, or if an error occured, the function instead returns 0.
Public Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Public verables
Public ControllEnabled As Boolean


'/******************************************************************************
Public Function GetWindowPoint(hwnd As Long, x As Long, y As Long) As Boolean
'/******************************************************************************
'Description: Display the width and height of window Width and height can
'be calculated from the coordinates returned in the rectangle.
'
'Inputs: hWind of target window.
'
'Returns: None at the moment.
'/******************************************************************************
Dim r As RECT  ' Receives window rectangle
Dim retval As Long  ' Return value
Dim TopLevelPerent As Long ' Holds the top level parents Hwnd

'Get current cursor position.
Call GetCursorPos(CursorPosition)
    
    'Get top level Parent
    TopLevelPerent = GetTopLevelParent(hwnd)
    retval = GetWindowRect(TopLevelPerent, r)  ' set r equal to Form1's rectangle
    
    y = CursorPosition.y - r.top - 20 ' < You can adjust this if there is an offset affect'
    x = CursorPosition.x - r.left
    
End Function
'/******************************************************************************



