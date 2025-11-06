Attribute VB_Name = "AZ_HideXButton_Mod"
'Private Const GWL_STYLE = (-16)
'Private Const WS_SYSMENU = &H80000 '0x00080000L
'Private Const WS_THICKFRAME = &H40000

Public Enum mcWindowStyle
    GWL_STYLE = &HFFF0
    WS_BORDER = &H800000   'L The window has a thin-line border
    WS_CAPTION = &HC00000  'L The window has a title bar (includes the WS_BORDER style).
    WS_CHILD = &H40000000 'L The window is a child window. A window with this style cannot have a menu bar. This style cannot be used with the WS_POPUP style.
    WS_CHILDWINDOW = &H40000000 'L Same as the WS_CHILD style.
    WS_CLIPCHILDREN = &H2000000 'L Excludes the area occupied by child windows when drawing occurs within the parent window. This style is used when creating the parent window.
    WS_CLIPSIBLINGS = &H4000000 'L Clips child windows relative to each other; that is, when a particular child window receives a WM_PAINT message, the WS_CLIPSIBLINGS style clips all other overlapping child windows out of the region of the child window to be updated. If WS_CLIPSIBLINGS is not specified and child windows overlap, it is possible, when drawing within the client area of a child window, to draw within the client area of a neighboring child window.
    WS_DISABLED = &H8000000 'L The window is initially disabled. A disabled window cannot receive input from the user. To change this after a window has been created, use the EnableWindow function.
    WS_DLGFRAME = &H400000  'L The window has a border of a style typically used with dialog boxes. A window with this style cannot have a title bar.
    WS_GROUP = &H20000    'L The window is the first control of a group of controls. The group consists of this first control and all controls defined after it, up to the next control with the WS_GROUP style. The first control in each group usually has the WS_TABSTOP style so that the user can move from group to group. The user can subsequently change the keyboard focus from one control in the group to the next control in the group by using the direction keys.
    'You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function.
    WS_HSCROLL = &H100000  'L The window has a horizontal scroll bar.
    WS_ICONIC = &H20000000 'L The window is initially minimized. Same as the WS_MINIMIZE style.
    WS_MAXIMIZE = &H1000000 'L The window is initially maximized.
    WS_MAXIMIZEBOX = &H10000   'L The window has a maximize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
    WS_MINIMIZE = &H20000000 'L The window is initially minimized. Same as the WS_ICONIC style.
    WS_MINIMIZEBOX = &H20000   'L The window has a minimize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
    WS_OVERLAPPED = &H0        'L The window is an overlapped window. An overlapped window has a title bar and a border. Same as the WS_TILED style.
    WS_POPUP = &H80000000 'L The window is a pop-up window. This style cannot be used with the WS_CHILD style.
    WS_SIZEBOX = &H40000   'L The window has a sizing border. Same as the WS_THICKFRAME style.
    WS_SYSMENU = &H80000   'L The window has a window menu on its title bar. The WS_CAPTION style must also be specified.
    WS_TABSTOP = &H10000   'L The window is a control that can receive the keyboard focus when the user presses the TAB key. Pressing the TAB key changes the keyboard focus to the next control with the WS_TABSTOP style.
    'You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function. For user-created windows and modeless dialogs to work with tab stops, alter the message loop to call the IsDialogMessage function.
    WS_THICKFRAME = &H40000    'L The window has a sizing border. Same as the WS_SIZEBOX style.
    WS_TILED = &H0        'L The window is an overlapped window. An overlapped window has a title bar and a border. Same as the WS_OVERLAPPED style.
    WS_VISIBLE = &H10000000 'L The window is initially visible. This style can be turned on and off by using the ShowWindow or SetWindowPos function.
    WS_VSCROLL = &H200000  'L The window has a vertical scroll bar.
    WS_OVERLAPPEDWINDOW = &HCF0000 '(WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX) 'The window is an overlapped window. Same as the WS_TILEDWINDOW style.
    WS_TILEDWINDOW = &HCF0000 '(WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX) 'The window is an overlapped window. Same as the WS_OVERLAPPEDWINDOW style.
    WS_POPUPWINDOW = &H80880000 '(WS_POPUP Or WS_BORDER Or WS_SYSMENU) 'The window is a pop-up window. The WS_CAPTION and WS_POPUPWINDOW styles must be combined to make the window menu visible.
End Enum

'Windows API calls to handle windows
'#If VBA7 Then
'    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'    Private Declare PtrSafe Function SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long) As Long
'#Else
'    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'    Private Declare Function SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long) As Long
'#End If
#If VBA7 Then
    #If Win64 Then
        Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As LongPtr
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As LongPtr
    Public TimerID As LongPtr
'    Dim lngWindow As LongPtr, lFrmHdl As LongPtr
#Else
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Public TimerID As Long
'    Dim lngWindow As Long, lFrmHdl As Long
#End If

Public Sub RemoveCloseButton(Frm As Object)
    Dim lngStyle As LongPtr, lngHWnd As LongPtr
    lngHWnd = FindWindow(vbNullString, Frm.Caption)
    lngStyle = GetWindowLong(lngHWnd, GWL_STYLE)
    If lngStyle And WS_SYSMENU > 0 Then SetWindowLong lngHWnd, GWL_STYLE, (lngStyle And Not WS_SYSMENU)
End Sub
Public Sub ResizableForm(Frm As Object)
    Dim lngStyle As LongPtr, lngHWnd As LongPtr
    lngHWnd = FindWindow(vbNullString, Frm.Caption)
    lngStyle = GetWindowLong(lngHWnd, GWL_STYLE) Or WS_THICKFRAME
    SetWindowLong lngHWnd, GWL_STYLE, lngStyle
End Sub
Sub HideTitleBar(Frm As Object)
'    #If VBA7 Then
    Dim lngStyle As LongPtr, lngHWnd As LongPtr
    lngHWnd = FindWindow(vbNullString, Frm.Caption)
    lngStyle = GetWindowLong(lngHWnd, GWL_STYLE)
    lngStyle = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lngHWnd, GWL_STYLE, lngStyle)
    Call DrawMenuBar(lngHWnd)
'    #Else
'        Dim lngWindow As Long, lFrmHdl As Long
'        lFrmHdl = FindWindow(vbNullString, Frm.Caption)
'        lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
'        lngWindow = lngWindow And (Not WS_CAPTION)
'        Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
'        Call DrawMenuBar(lFrmHdl)
'    #End If
End Sub
Sub CenterUserform(uf As Object)
    'On Error Resume Next
  With uf
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
'    .Top = (Application.Height / 2) - (.Height / 2)
'    .Left = (Application.Width / 2) - (.Width / 2)
  End With
End Sub
