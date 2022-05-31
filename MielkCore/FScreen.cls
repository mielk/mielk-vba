VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FScreen"
'----------------------------------------------------------------------------------------------------------

Public Function getSize() As area
    With getSize
        .width = GetSystemMetrics(SM_CXSCREEN) * PIXEL_SIZE
        .height = GetSystemMetrics(SM_CYSCREEN) * PIXEL_SIZE
        .left = GetSystemMetrics(SM_XVIRTUALSCREEN) * PIXEL_SIZE
        .top = GetSystemMetrics(SM_YVIRTUALSCREEN) * PIXEL_SIZE
    End With
End Function

Public Function getScreenWidth() As Single
    getScreenWidth = GetSystemMetrics(SM_CXSCREEN)
End Function

Public Function getScreenHeight() As Single
    getScreenHeight = GetSystemMetrics(SM_CYSCREEN)
End Function




Public Function getActiveMonitorWorkingArea() As area
#If Win64 Then
    Dim monitorHandle As LongPtr
#Else
    Dim monitorHandle As Long
#End If
    Dim cursorPosition As POINTAPI
    Dim info As MONITORINFOEX
    '------------------------------------------------------------------------------------------------------
    
    Call GetCursorPos(cursorPosition)
    monitorHandle = MonitorFromPoint(cursorPosition.x, cursorPosition.y, MONITOR_DEFAULTTONEAREST)
    info.cbSize = Len(info)
    If GetMonitorInfo(monitorHandle, info) = False Then
        Call VBA.err.raise(1, description:="Error when trying to read monitor info")
    End If
    
    With info.rcWork
        getActiveMonitorWorkingArea.left = .left * PIXEL_SIZE
        getActiveMonitorWorkingArea.top = .top * PIXEL_SIZE
        getActiveMonitorWorkingArea.width = (.right - .left) * PIXEL_SIZE
        getActiveMonitorWorkingArea.height = (.bottom - .top) * PIXEL_SIZE
    End With
    
End Function

Public Function getDpiScaling() As Coordinate
    Const DPI_BASE_SETTING As Long = 96
    '------------------------------------------------------------------------------------------------------
    Dim dpiX As Long
    Dim dpiY As Long
    '------------------------------------------------------------------------------------------------------

    If GetDpiForMonitor(MonitorFromWindow(Excel.Application.hWnd, MONITOR_DEFAULTTONEAREST), MDT_DEFAULT, dpiX, dpiY) = S_OK Then
        With getDpiScaling
            .x = dpiX / DPI_BASE_SETTING
            .y = dpiY / DPI_BASE_SETTING
        End With
    Else
        Call VBA.err.raise(1, description:="Error when trying to read DPI scaling")
    End If

End Function

Public Function getCursorPosition() As Coordinate
    Dim coordinates As POINTAPI
    '------------------------------------------------------------------------------------------------------
    
    Call GetCursorPos(coordinates)
    With getCursorPosition
        .x = coordinates.x
        .y = coordinates.y
    End With
    
End Function
