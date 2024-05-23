Attribute VB_Name = "modUIWindowsApi"
Option Explicit

Private Const CLASS_NAME As String = "modWindowsApi"
'----------------------------------------------------------------------------------------------------------

#If VBA7 And Win64 Then
    Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal punk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal punktTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function EnableMenuItem Lib "user32" (ByVal hMenu As LongPtr, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    'Public Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As LongPtr
    Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As LongPtr) As Long
    Public Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As LongPtr
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function GetForegroundWindow Lib "user32.dll" () As LongPtr
    Public Declare PtrSafe Function GetMenuItemCount Lib "user32" (ByVal hMenu As LongPtr) As Long
    Public Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr
    Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Boolean
    'Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As LongPtr) As Long
    Public Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, ByRef lpdwProcessId As LongPtr) As LongPtr
    Public Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As LongPtr) As Long
    Public Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As LongPtr) As Long
    Public Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As Long, LPType As LongPtr, LPData As Any, lpcbData As LongPtr) As Long
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Public Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare PtrSafe Function setParent Lib "user32" Alias "SetParent" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
#ElseIf VBA7 Then
    Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal punk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal punktTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function EnableMenuItem Lib "user32" (ByVal hMenu As LongPtr, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    'Public Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As LongPtr
    Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As LongPtr) As Long
    Public Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As LongPtr
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function GetForegroundWindow Lib "user32.dll" () As LongPtr
    Public Declare PtrSafe Function GetMenuItemCount Lib "user32" (ByVal hMenu As LongPtr) As Long
    Public Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr
    Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Boolean
    Public Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As Long
    'Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As LongPtr) As Long
    Public Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, ByRef lpdwProcessId As LongPtr) As LongPtr
    Public Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As LongPtr) As Long
    Public Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As LongPtr) As Long
    Public Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As Long, LPType As LongPtr, LPData As Any, lpcbData As LongPtr) As Long
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Public Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare PtrSafe Function setParent Lib "user32" Alias "SetParent" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
    Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal punk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal punktTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As Long) As Long
    Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
    Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    'Public Declare Function GetActiveWindow Lib "user32.dll" () As Long
    Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function GetDesktopWindow Lib "user32" () As Long
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
    Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
    Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Boolean
    'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
    Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
    Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
    Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, LPType As Long, LPData As Any, lpcbData As Long) As Long
    Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare Function setParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
#End If




'[Enums] --------------------------------------------------------------------------------------------------
Public Enum FORM_PARENT_WINDOW_TYPE
    FORM_PARENT_NONE = 0
    FORM_PARENT_APPLICATION = 1
    FORM_PARENT_WINDOW = 2
End Enum

'Constants for window size
Public Const SW_HIDE = 0                            'Hide the window.
Public Const SW_MAXIMIZE = 3                        'Maximize the window.
Public Const SW_MINIMIZE = 6                        'Minimize the window.
Public Const SW_RESTORE = 9                         'Restore the window (not maximized nor minimized).
Public Const SW_SHOW = 5                            'Show the window.
Public Const SW_SHOWMAXIMIZED = 3                   'Show the window maximized.
Public Const SW_SHOWMINIMIZED = 2                   'Show the window minimized.
Public Const SW_SHOWMINNOACTIVE = 7                 'Show the window minimized but do not activate it.
Public Const SW_SHOWNA = 8                          'Show the window in its current state but do not activate it.
Public Const SW_SHOWNOACTIVATE = 4                  'Show the window in its most recent size and position but do not activate it.
Public Const SW_SHOWNORMAL = 1                      'Show the window and activate it (as usual).

'Constants for title bar
Public Enum WindowLongFlags
     DWLP_DLGPROC = &H4
     DWLP_MSGRESULT = &H0
     DWLP_USER = &H8
     GWL_EXSTYLE = -20                              'The offset of a window's extended style
     GWL_ID = -12
     GWL_STYLE = -16                                'The offset of a window's style
     GWL_USERDATA = -21
     GWL_WNDPROC = -4
     GWLP_HINSTANCE = -6
     GWLP_HWNDPARENT = -8
End Enum


Public Const WS_CAPTION As Long = &HC00000          'Style to add a titlebar
Public Const WS_EX_DLGMODALFRAME As Long = &H1      'Controls if the window has an icon
Public Const WS_EX_LAYERED = &H8000
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_POPUP As Long = &H80000000
Public Const WS_SIZEBOX = &H40000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000


'Constants for transparency
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2

'Other constants
Public Const C_USERFORM_CLASSNAME As String = "ThunderDFrame"
Public Const HWND_TOP = 0
Public Const SWP_DRAWFRAME = &H20
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
