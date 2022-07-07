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
Private pMonitors As VBA.Collection
'----------------------------------------------------------------------------------------------------------




Public Sub reloadMonitorsInfo()
    Call f.files.deleteFile(getScreenHelperTextFilePath)
    Call enumerateMonitors
    Call readMonitorsInfoFromScreenHelperTextFile
End Sub

Private Sub enumerateMonitors()
    EnumDisplayMonitors 0, 0, AddressOf monitorEnumProc, 0
End Sub

Private Sub readMonitorsInfoFromScreenHelperTextFile()
    Dim filepath As String
    Dim info As Variant
    Dim i As Long
    Dim Monitor As Monitor
    '------------------------------------------------------------------------------------------------------
    
    Set pMonitors = New VBA.Collection
    
    filepath = getScreenHelperTextFilePath
    If f.files.fileExists(filepath) Then
        info = f.TextFiles.getCsvContentAsArray(filepath, ",")
        If f.arrays.isDefined(info) Then
            For i = LBound(info, 1) To UBound(info, 1)
                Set Monitor = New Monitor
                With Monitor
                    Call .setHandle(VBA.CLngPtr(info(i, 1)))
                    Call .setBounds(VBA.CLng(info(i, 2)), VBA.CLng(info(i, 3)), VBA.CLng(info(i, 4)), VBA.CLng(info(i, 5)))
                End With
                Call pMonitors.add(Monitor)
            Next i
        End If
    End If
    
End Sub

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

#If VBA7 And Win64 Then
    Dim monitorHandle As LongPtr
#Else
    Dim monitorHandle As Long
#End If
    Dim cursorPosition As POINTAPI
    Dim info As MONITORINFOEX
    '------------------------------------------------------------------------------------------------------
    
    Call GetCursorPos(cursorPosition)
    
#If VBA7 And Win64 Then
'    Dim Monitor As Monitor
'
'    Call reloadMonitorsInfo
'
'    If f.Collections.isNonEmptyCollection(pMonitors) Then
'        For Each Monitor In pMonitors
'            If Monitor.isPointInMonitor(cursorPosition) Then
'                monitorHandle = Monitor.getHandle
'                'getActiveMonitorWorkingArea = monitor.getWorkingArea
'            End If
'        Next Monitor
'    End If
    monitorHandle = MonitorFromPoint(cursorPosition, MONITOR_DEFAULTTONEAREST)
#Else
    monitorHandle = MonitorFromPoint(cursorPosition, MONITOR_DEFAULTTONEAREST)
#End If

    info.cbSize = Len(info)
    If getMonitorInfo(monitorHandle, info) = False Then
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

Public Sub test()
    EnumDisplayMonitors 0, 0, AddressOf monitorEnumProc, 0
End Sub
