Attribute VB_Name = "modWindowsApiDeclarations"
Option Explicit

'[Constants] ----------------------------------------------------------------------------------------------
Public Const C_ALPHA_FULL_OPAQUE As Byte = 255
Public Const C_ALPHA_FULL_TRANSPARENT As Byte = 0
Public Const C_EXCEL_APP_CLASSNAME = "XLMain"
Public Const C_EXCEL_DESK_CLASSNAME = "XLDesk"
Public Const C_EXCEL_WINDOW_CLASSNAME = "EXCEL7"
Public Const HWND_NOTOPMOST = -2
Public Const ICON_BIG = 1&
Public Const ICON_SMALL = 0&
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const MF_BYPOSITION = &H400
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_REMOVE = &H1000
Public Const WM_SETICON = &H80
Public Const HKEY_CLASSES_ROOT  As Long = &H80000000
Public Const HKEY_CURRENT_CONFIG  As Long = &H80000005
Public Const HKEY_DYN_DATA  As Long = &H80000006
Public Const HKEY_PERFORMANCE_DATA  As Long = &H80000004
Public Const HKEY_USERS  As Long = &H80000003
Public Const KEY_ALL_ACCESS  As Long = &H3F
Public Const ERROR_SUCCESS  As Long = 0&
Public Const HKCU  As Long = HKEY_CURRENT_USER
Public Const HKLM  As Long = HKEY_LOCAL_MACHINE
'----------------------------------------------------------------------------------------------------------

