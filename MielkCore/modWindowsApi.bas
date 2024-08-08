Attribute VB_Name = "modWindowsApi"
'Public Declare PtrSafe Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
'Public Declare PtrSafe Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
'Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Option Explicit

Private Const CLASS_NAME As String = "WindowsApi"
'----------------------------------------------------------------------------------------------------------

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If VBA7 And Win64 Then
    Private Type POINTAPI_AsLongLong
        value As LongLong
    End Type
    
    Public Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As LongPtr
    Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal length As LongLong)
    Public Declare PtrSafe Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Public Declare PtrSafe Function EnumDisplayMonitors Lib "user32.dll" (ByVal hDC As LongPtr, ByRef lprcClip As Any, ByVal lpfnEnum As LongPtr, ByVal dwData As LongPtr) As Boolean
    Public Declare PtrSafe Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As LongPtr, ByVal param As LongPtr) As Long
    Public Declare PtrSafe Function FindWindows Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function apiGetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Boolean
    Public Declare PtrSafe Function GetDpiForMonitor Lib "shcore" (ByVal hMonitor As LongPtr, ByVal dpiType As MONITOR_DPI_TYPE, ByRef dpiX As Long, ByRef dpiY As Long) As Long
    Public Declare PtrSafe Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As LongPtr, ByRef lpMI As MONITORINFOEX) As Boolean
    Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal index As Integer) As Integer
    Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    'Public Declare PtrSafe Function MonitorFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As LongPtr) As LongPtr
    'Public Declare PtrSafe Function MonitorFromPoint Lib "user32" (point As POINTAPI, ByVal dwFlags As LongPtr) As LongPtr
    Public Declare PtrSafe Function MonitorFromPointInternal Lib "user32.dll" Alias "MonitorFromPoint" (ByVal pt As LongLong, ByVal dwFlags As Long) As LongPtr
    Public Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal dwFlags As MONITOR_DEFAULTS) As LongPtr
    Public Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
    '[Decimal separator]
    Public Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Public Declare PtrSafe Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
    Public Declare PtrSafe Function GetUserDefaultLCID% Lib "kernel32" ()
#ElseIf VBA7 Then
    Public Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As LongPtr
    Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal length As Long)
    Public Declare PtrSafe Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Public Declare PtrSafe Function EnumDisplayMonitors Lib "user32.dll" (ByVal hDC As LongPtr, ByRef lprcClip As Any, ByVal lpfnEnum As LongPtr, ByVal dwData As LongPtr) As Boolean
    Public Declare PtrSafe Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As LongPtr, ByVal param As LongPtr) As Long
    Public Declare PtrSafe Function FindWindows Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function apiGetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Boolean
    Public Declare PtrSafe Function GetDpiForMonitor Lib "shcore" (ByVal hMonitor As LongPtr, ByVal dpiType As MONITOR_DPI_TYPE, ByRef dpiX As Long, ByRef dpiY As Long) As Long
    Public Declare PtrSafe Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As LongPtr, ByRef lpMI As MONITORINFOEX) As Boolean
    Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal index As Integer) As Integer
    Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    'Public Declare PtrSafe Function MonitorFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As LongPtr) As LongPtr
    'Public Declare PtrSafe Function MonitorFromPoint Lib "user32" (point As POINTAPI, ByVal dwFlags As LongPtr) As LongPtr
    Public Declare PtrSafe Function MonitorFromPointInternal Lib "user32.dll" Alias "MonitorFromPoint" (ByVal X As Long, ByVal Y As Long, ByVal dwFlags As Long) As LongPtr
    Public Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal dwFlags As MONITOR_DEFAULTS) As LongPtr
    Public Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
    '[Decimal separator]
    Public Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Public Declare PtrSafe Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
    Public Declare PtrSafe Function GetUserDefaultLCID% Lib "kernel32" ()
#Else
    Public Declare Function GetActiveWindow Lib "user32.dll" () As Long
    Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
    Public Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Public Declare Function EnumDisplayMonitors Lib "user32.dll" (ByVal hdc As Long, ByRef lprcClip As Any, ByVal lpfnEnum As Long, ByVal dwData As Long) As Boolean
    Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal param As Long) As Long
    Public Declare Function FindWindows Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function apiGetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Boolean
    Public Declare Function GetDpiForMonitor Lib "shcore" (ByVal hMonitor As Long, ByVal dpiType As MONITOR_DPI_TYPE, ByRef dpiX As Long, ByRef dpiY As Long) As Long
    Public Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFOEX) As Boolean
    Public Declare Function GetSystemMetrics Lib "user32" (ByVal index As Integer) As Integer
    Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    'Public Declare Function MonitorFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
    Public Declare Function MonitorFromPoint Lib "user32" (point As POINTAPI, ByVal dwFlags As Long) As Long
    Public Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As Long, ByVal dwFlags As MONITOR_DEFAULTS) As Long
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
    '[Decimal separator]
    Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
    Public Declare Function GetUserDefaultLCID% Lib "kernel32" ()
#End If


'Regional settings
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_SLIST = &HC
Public Const LOCALE_SDATE = &H1D            ' date separator
Public Const LOCALE_STIME = &H1E            ' time separator
Public Const LOCALE_SSHORTDATE = &H1F       ' short date format string
Public Const LOCALE_SLONGDATE = &H20        ' long date format string


'Screen
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_XVIRTUALSCREEN = 76
Public Const SM_YVIRTUALSCREEN = 77
Public Const SM_CMONITORS = 80              ' number of display monitors

'DPI
Public Const S_OK = 0

Public Enum MONITOR_DPI_TYPE
    MDT_EFFECTIVE_DPI = 0
    MDT_ANGULAR_DPI = 1
    MDT_RAW_DPI = 2
    MDT_DEFAULT = MDT_EFFECTIVE_DPI
End Enum

Public Enum MONITOR_DEFAULTS
    MONITOR_DEFAULTTONULL = &H0&
    MONITOR_DEFAULTTOPRIMARY = &H1&
    MONITOR_DEFAULTTONEAREST = &H2&
End Enum

'Monitor info
Public Const CCHDEVICENAME = 32
Public Const MONITORINFOF_PRIMARY = &H1

Public Type MONITORINFOEX
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
    szDevice As String * CCHDEVICENAME
End Type

Public Enum REG_DATA_TYPE
    REG_DATA_TYPE_DEFAULT = 0   ' Default based on data type of value.
    REG_INVALID = -1            ' Invalid
    REG_SZ = 1                  ' String
    REG_DWORD = 4               ' Long
End Enum



'[Callbacks]            Must be declared in regular module.

#If VBA7 Then
    Public Function monitorEnumProc(ByVal hMonitor As LongPtr, ByVal hdcMonitor As LongPtr, ByRef rMonitor As RECT, ByVal dwData As LongPtr) As Boolean
#Else
    Public Function monitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, ByRef rMonitor As RECT, ByVal dwData As Long) As Boolean
#End If
    Dim str As String
    Dim path As String
    Dim intFile As Integer
    '------------------------------------------------------------------------------------------------------
    
    path = getScreenHelperTextFilePath
    str = hMonitor & "," & rMonitor.Top & "," & rMonitor.Right & "," & rMonitor.Bottom & "," & rMonitor.Left & VBA.vbCrLf
    
    intFile = VBA.FreeFile
    Open path For Append As #intFile
    Print #intFile, str;
    Close intFile
    
    monitorEnumProc = True
End Function

Public Function getScreenHelperTextFilePath() As String
    getScreenHelperTextFilePath = ThisWorkbook.path & "\screens.txt"
End Function



#If Win64 Then
    Public Function getMonitorFromPoint(pt As POINTAPI, ByVal dwFlags As Long) As LongPtr
        Dim t As POINTAPI_AsLongLong
        LSet t = pt
      
        getMonitorFromPoint = MonitorFromPointInternal(t.value, dwFlags)
    End Function
#Else
    Public Function getMonitorFromPoint(pt As POINTAPI, ByVal dwFlags As Long) As LongPtr
        getMonitorFromPoint = MonitorFromPointInternal(pt.X, pt.Y, dwFlags)
    End Function
#End If

