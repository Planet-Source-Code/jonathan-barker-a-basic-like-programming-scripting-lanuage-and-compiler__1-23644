Attribute VB_Name = "Module1"
Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Declare Function DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'  Define information of the window (pointed to by hWnd)
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
Type POINTAPI
    x As Long
    y As Long
End Type
Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

' Class styles
Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_DBLCLKS = &H8
Public Const CS_OWNDC = &H20
Public Const CS_CLASSDC = &H40
Public Const CS_PARENTDC = &H80
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOCLOSE = &H200
Public Const CS_SAVEBITS = &H800
Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_PUBLICCLASS = &H4000
' Window styles
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_CHILDWINDOW = (WS_CHILD)
' ExWindowStyles
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&
' Color constants
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
' Window messages
Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
' ShowWindow commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10
' Standard ID's of cursors
Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_WAIT = 32514&
Public Const IDC_CROSS = 32515&
Public Const IDC_UPARROW = 32516&
Public Const IDC_SIZE = 32640&
Public Const IDC_ICON = 32641&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_NO = 32648&
Public Const IDC_APPSTARTING = 32650&
Public Const GWL_WNDPROC = -4
    

Dim hwnd2 As Long, hwnd3 As Long, old_proc As Long, new_proc As Long
Public Function Main()
    Dim lngTemp As Long
    If MyRegisterClass Then
        If MyCreateWindow Then
'            new_proc = GetMyWndProc(AddressOf ButtonProc)
'            old_proc = SetWindowLong(hwnd2, GWL_WNDPROC, new_proc)
'            MyMessageLoop
        End If
        MyUnregisterClass
    End If
End Function
Private Function MyRegisterClass() As Boolean
    ' WNDCLASS-structure
    Dim wndcls As WNDCLASS
    wndcls.style = CS_HREDRAW + CS_VREDRAW
    wndcls.lpfnwndproc = GetMyWndProc(AddressOf MyWndProc)
    wndcls.cbClsextra = 0
    wndcls.cbWndExtra2 = 0
    wndcls.hInstance = App.hInstance
    wndcls.hIcon = 0
    wndcls.hCursor = LoadCursor(0, IDC_ARROW)
    wndcls.hbrBackground = COLOR_WINDOW
    wndcls.lpszMenuName = 0
    wndcls.lpszClassName = "myWindowClass"
    ' Register class
    MyRegisterClass = (RegisterClass(wndcls) <> 0)
End Function
Private Sub MyUnregisterClass()
    UnregisterClass "myWindowClass", App.hInstance
End Sub
Private Function MyCreateWindow() As Boolean
    ' Create the window
    hWnd = CreateWindowEx(0, "myWindowClass", "My Window", WS_OVERLAPPEDWINDOW, 0, 0, 400, 300, 0, 0, App.hInstance, ByVal 0&)
    ' The Button and Textbox are child windows
    hwnd2 = CreateWindowEx(0, "Button", "My button", WS_CHILD, 50, 55, 100, 25, hWnd, 0, App.hInstance, ByVal 0&)
    hwnd3 = CreateWindowEx(0, "edit", "My textbox", WS_CHILD, 50, 25, 100, 25, hWnd, 0, App.hInstance, ByVal 0&)
    If hWnd <> 0 Then ShowWindow hWnd, SW_SHOWNORMAL
    ' Show them
    ShowWindow hwnd2, SW_SHOWNORMAL
    ShowWindow hwnd3, SW_SHOWNORMAL
    ' Go back
    MyCreateWindow = (hWnd <> 0)
End Function
Private Function MyWndProc(ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case message
        Case WM_DESTROY
            ' Destroy window
            PostQuitMessage (0)
    End Select
    ' calls the default window procedure
    MyWndProc = DefWindowProc(hWnd, message, wParam, lParam)
End Function
Function GetMyWndProc(ByVal lWndProc As Long) As Long
    GetMyWndProc = lWndProc
End Function
Private Sub MyMessageLoop()
    Dim aMsg As Msg
    Do While GetMessage(aMsg, 0, 0, 0)
        DispatchMessage aMsg
    Loop
End Sub
Private Function ButtonProc(ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim x As Integer
    If (message = 533) Then
        x = MsgBox("You clicked on the button", vbOKOnly)
    End If
    ' calls the window procedure
    ButtonProc = CallWindowProc(old_proc, hWnd, message, wParam, lParam)
End Function


