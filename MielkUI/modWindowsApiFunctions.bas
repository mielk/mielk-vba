Attribute VB_Name = "modWindowsApiFunctions"
Option Explicit
Option Compare Text

Private Const CLASS_NAME As String = "modFormControl"



Public Function ShowMaximizeButton(uf As MSForms.UserForm, HideButton As Boolean) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim winInfo As LongPtr
    Dim R As LongPtr
#Else
    Dim UFHWnd As Long
    Dim winInfo As Long
    Dim R As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        ShowMaximizeButton = False
        Exit Function
    End If
    
    winInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If HideButton = False Then
        winInfo = winInfo Or WS_MAXIMIZEBOX
    Else
        winInfo = winInfo And (Not WS_MAXIMIZEBOX)
    End If
    R = SetWindowLong(UFHWnd, GWL_STYLE, winInfo)
    
    ShowMaximizeButton = (R <> 0)

End Function

Public Function ShowMinimizeButton(uf As MSForms.UserForm, HideButton As Boolean) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim winInfo As LongPtr
    Dim R As LongPtr
#Else
    Dim UFHWnd As Long
    Dim winInfo As Long
    Dim R As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        ShowMinimizeButton = False
        Exit Function
    End If
    
    winInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If HideButton = False Then
        winInfo = winInfo Or WS_MINIMIZEBOX
    Else
        winInfo = winInfo And (Not WS_MINIMIZEBOX)
    End If
    R = SetWindowLong(UFHWnd, GWL_STYLE, winInfo)
    
    ShowMinimizeButton = (R <> 0)

End Function

Public Function HasMinimizeButton(uf As MSForms.UserForm) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim winInfo As LongPtr
    Dim R As LongPtr
#Else
    Dim UFHWnd As Long
    Dim winInfo As Long
    Dim R As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        HasMinimizeButton = False
        Exit Function
    End If
    
    winInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    
    If winInfo And WS_MINIMIZEBOX Then
        HasMinimizeButton = True
    Else
        HasMinimizeButton = False
    End If
    
End Function

Public Function HasMaximizeButton(uf As MSForms.UserForm) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim winInfo As LongPtr
    Dim R As LongPtr
#Else
    Dim UFHWnd As Long
    Dim winInfo As Long
    Dim R As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        HasMaximizeButton = False
        Exit Function
    End If
    
    winInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    
    If winInfo And WS_MAXIMIZEBOX Then
        HasMaximizeButton = True
    Else
        HasMaximizeButton = False
    End If
    
End Function


Public Function SetFormParent(uf As MSForms.UserForm, parent As FORM_PARENT_WINDOW_TYPE) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim WindHWnd As LongPtr
    Dim R As LongPtr
#Else
    Dim UFHWnd As Long
    Dim WindHWnd As Long
    Dim R As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        SetFormParent = False
        Exit Function
    End If
    
    Select Case parent
        Case FORM_PARENT_APPLICATION
            R = setParent(UFHWnd, Application.hWnd)
        Case FORM_PARENT_NONE
            R = setParent(UFHWnd, 0&)
        Case FORM_PARENT_WINDOW
            If Application.ActiveWindow Is Nothing Then
                SetFormParent = False
                Exit Function
            End If
            WindHWnd = WindowHWnd(Application.ActiveWindow)
            If WindHWnd = 0 Then
                SetFormParent = False
                Exit Function
            End If
            R = setParent(UFHWnd, WindHWnd)
        Case Else
            SetFormParent = False
            Exit Function
    End Select
    
    SetFormParent = (R <> 0)

End Function


Public Function IsCloseButtonVisible(uf As MSForms.UserForm) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim winInfo As LongPtr
    Dim R As LongPtr
#Else
    Dim UFHWnd As Long
    Dim winInfo As Long
    Dim R As Long
#End If
    '------------------------------------------------------------------------------------------------------

    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        IsCloseButtonVisible = False
        Exit Function
    End If
    
    winInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    IsCloseButtonVisible = (winInfo And WS_SYSMENU)
    
End Function


Public Function ShowCloseButton(uf As MSForms.UserForm, HideButton As Boolean) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim winInfo As LongPtr
    Dim R As LongPtr
#Else
    Dim UFHWnd As Long
    Dim winInfo As Long
    Dim R As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        Exit Function
    End If
    
    winInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If HideButton = False Then
        ' set the SysMenu bit
        winInfo = winInfo Or WS_SYSMENU
    Else
        ' clear the SysMenu bit
        winInfo = winInfo And (Not WS_SYSMENU)
    End If
    
    R = SetWindowLong(UFHWnd, GWL_STYLE, winInfo)
    ShowCloseButton = (R <> 0)
    
End Function


Public Function IsCloseButtonEnabled(uf As MSForms.UserForm) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim hMenu As LongPtr
    Dim ItemCount As Long
    Dim PrevState As Long
#Else
    Dim UFHWnd As Long
    Dim hMenu As Long
    Dim ItemCount As Long
    Dim PrevState As Long
#End If
    '------------------------------------------------------------------------------------------------------

    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        IsCloseButtonEnabled = False
        Exit Function
    End If
    ' Get the menu handle
    hMenu = GetSystemMenu(UFHWnd, 0&)
    If hMenu = 0 Then
        IsCloseButtonEnabled = False
        Exit Function
    End If
    
    ItemCount = GetMenuItemCount(hMenu)
    ' Disable the button. This returns MF_DISABLED or MF_ENABLED indicating
    ' the previous state of the item.
    PrevState = EnableMenuItem(hMenu, ItemCount - 1, MF_DISABLED Or MF_BYPOSITION)
    
    If PrevState = MF_DISABLED Then
        IsCloseButtonEnabled = False
    Else
        IsCloseButtonEnabled = True
    End If
    ' restore the previous state
    EnableCloseButton uf, (PrevState = MF_DISABLED)

    DrawMenuBar UFHWnd

End Function


Public Function EnableCloseButton(uf As MSForms.UserForm, Disable As Boolean) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim hMenu As LongPtr
    Dim ItemCount As Long
    Dim res As Long
#Else
    Dim UFHWnd As Long
    Dim hMenu As Long
    Dim ItemCount As Long
    Dim res As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    ' Get the HWnd of the UserForm.
    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        EnableCloseButton = False
        Exit Function
    End If
    ' Get the menu handle
    hMenu = GetSystemMenu(UFHWnd, 0&)
    If hMenu = 0 Then
        EnableCloseButton = False
        Exit Function
    End If
    
    ItemCount = GetMenuItemCount(hMenu)
    If Disable = True Then
        res = EnableMenuItem(hMenu, ItemCount - 1, MF_DISABLED Or MF_BYPOSITION)
    Else
        res = EnableMenuItem(hMenu, ItemCount - 1, MF_ENABLED Or MF_BYPOSITION)
    End If
    If res = -1 Then
        EnableCloseButton = False
        Exit Function
    End If
    DrawMenuBar UFHWnd
    
    EnableCloseButton = True
    
End Function

Public Function showTitleBar(uf As MSForms.UserForm, titleBarVisible As Boolean) As Boolean
#If VBA7 Then
    Dim ufHandle As LongPtr
    Dim winInfo As LongPtr
    Dim result As LongPtr
#Else
    Dim ufHandle As Long
    Dim winInfo As Long
    Dim result As Long
#End If
    '------------------------------------------------------------------------------------------------------

    ufHandle = HWndOfUserForm(uf)
    If ufHandle = 0 Then
        showTitleBar = False
        Exit Function
    End If
    
    winInfo = GetWindowLong(ufHandle, GWL_STYLE)
    
    If titleBarVisible Then
        winInfo = winInfo Or WS_CAPTION
    Else
        winInfo = winInfo And (Not WS_CAPTION)
    End If
    
    result = SetWindowLong(ufHandle, GWL_STYLE, winInfo)
    showTitleBar = (result <> 0)
    
End Function

Public Function hideTitleBar(uf As MSForms.UserForm) As Boolean
#If VBA7 Then
    Dim lngWindow As LongPtr
    Dim handle As LongPtr
#Else
    Dim lngWindow As Long
    Dim handle As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    handle = HWndOfUserForm(uf)
    'lFrmHdl = FindWindowA(vbNullString, frm.caption)
    lngWindow = GetWindowLong(handle, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(handle, GWL_STYLE, lngWindow)
    Call DrawMenuBar(handle)

End Function


Public Function IsTitleBarVisible(uf As MSForms.UserForm) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim winInfo As LongPtr
    Dim R As LongPtr
#Else
    Dim UFHWnd As Long
    Dim winInfo As Long
    Dim R As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        IsTitleBarVisible = False
        Exit Function
    End If
    
    winInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    IsTitleBarVisible = (winInfo And WS_CAPTION)
    
End Function

Public Function MakeFormResizable(uf As MSForms.UserForm, Sizable As Boolean) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim winInfo As LongPtr
    Dim R As LongPtr
#Else
    Dim UFHWnd As Long
    Dim winInfo As Long
    Dim R As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        MakeFormResizable = False
        Exit Function
    End If
    
    winInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If Sizable = True Then
        winInfo = winInfo Or WS_SIZEBOX
    Else
        winInfo = winInfo And (Not WS_SIZEBOX)
    End If
    
    R = SetWindowLong(UFHWnd, GWL_STYLE, winInfo)
    MakeFormResizable = (R <> 0)
        
End Function

Public Function IsFormResizable(uf As MSForms.UserForm) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim winInfo As LongPtr
    Dim R As LongPtr
#Else
    Dim UFHWnd As Long
    Dim winInfo As Long
    Dim R As Long
#End If
    '------------------------------------------------------------------------------------------------------

    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        IsFormResizable = False
        Exit Function
    End If
    
    winInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    IsFormResizable = (winInfo And WS_SIZEBOX)
    
End Function


Public Function SetFormOpacity(uf As MSForms.UserForm, Opacity As Byte) As Boolean
#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim WinL As LongPtr
    Dim res As Variant
#Else
    Dim UFHWnd As Long
    Dim WinL As Long
    Dim res As Variant
#End If
    '------------------------------------------------------------------------------------------------------
    
    SetFormOpacity = False
    
    UFHWnd = HWndOfUserForm(uf)
    If UFHWnd = 0 Then
        Exit Function
    End If
    
    WinL = GetWindowLong(UFHWnd, GWL_EXSTYLE)
    If WinL = 0 Then
        Exit Function
    End If
    
    res = SetWindowLong(UFHWnd, GWL_EXSTYLE, WinL Or WS_EX_LAYERED)
    If res = 0 Then
        Exit Function
    End If
    
    res = SetLayeredWindowAttributes(UFHWnd, 0, Opacity, LWA_ALPHA)
    If res = 0 Then
        Exit Function
    End If
    
    SetFormOpacity = True
    
End Function



#If VBA7 Then
Public Function HWndOfUserForm(uf As MSForms.UserForm) As LongPtr
    Dim appHWnd As LongPtr
    Dim deskHWnd As LongPtr
    Dim WinHWnd As LongPtr
    Dim UFHWnd As LongPtr
    Dim cap As String
    Dim WindowCap As String
#Else
Public Function HWndOfUserForm(uf As MSForms.UserForm) As Long
    Dim appHWnd As Long
    Dim deskHWnd As Long
    Dim WinHWnd As Long
    Dim UFHWnd As Long
    Dim cap As String
    Dim WindowCap As String
#End If
    '------------------------------------------------------------------------------------------------------
    
    cap = uf.caption
    
    ' First, look in top level windows
    UFHWnd = FindWindow(C_USERFORM_CLASSNAME, cap)
    If UFHWnd <> 0 Then
        HWndOfUserForm = UFHWnd
        Exit Function
    End If
    
    ' Not a top level window. Search for child of application.
    appHWnd = Application.hWnd
    UFHWnd = FindWindowEx(appHWnd, 0&, C_USERFORM_CLASSNAME, cap)
    If UFHWnd <> 0 Then
        HWndOfUserForm = UFHWnd
        Exit Function
    End If
    
    ' Not a child of the application.
    ' Search for child of ActiveWindow (Excel's ActiveWindow, not
    ' Window's ActiveWindow).
    If Application.ActiveWindow Is Nothing Then
        HWndOfUserForm = 0
        Exit Function
    End If
    WinHWnd = WindowHWnd(Application.ActiveWindow)
    UFHWnd = FindWindowEx(WinHWnd, 0&, C_USERFORM_CLASSNAME, cap)
    HWndOfUserForm = UFHWnd
    
End Function

Public Sub maximizeUserForm(uf As UserForm)
#If VBA7 Then
    Dim hWnd As LongPtr
#Else
    Dim hWnd As Long
#End If
    '------------------------------------------------------------------------------------------------------
    hWnd = HWndOfUserForm(uf)
    Call ShowWindow(hWnd, SW_MAXIMIZE)
End Sub

#If VBA7 Then
Public Sub showFormOnAppTaskBar(form As UserForm, Optional ByVal hWnd As LongPtr)
    Dim wStyle As LongPtr
    Dim result As LongPtr
    Dim resultPos As Long
#Else
Public Sub showFormOnAppTaskBar(form As UserForm, Optional ByVal hWnd As Long)
    Dim wStyle As Long
    Dim result As Long
    Dim resultPos As Long
#End If
    '------------------------------------------------------------------------------------------------------
    If hWnd = 0 Then hWnd = FindWindow(vbNullString, form.caption)
    wStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    wStyle = wStyle Or WS_EX_APPWINDOW
    resultPos = SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_HIDEWINDOW)
    result = SetWindowLong(hWnd, GWL_EXSTYLE, wStyle)
    resultPos = SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
End Sub

#If VBA7 Then
Public Sub AddIcon(form As UserForm, Optional ByVal hWnd As LongPtr)
    Dim lngRet As Variant
    Dim hIcon As Long
#Else
Public Sub AddIcon(form As UserForm, Optional ByVal hWnd As Long)
    Dim lngRet As Variant
    Dim hIcon As Long
#End If
    '------------------------------------------------------------------------------------------------------
    'hIcon = Sheet1.Image1.picture.Handle
    If hWnd = 0 Then hWnd = FindWindow(vbNullString, form.caption)
    lngRet = SendMessage(hWnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
    lngRet = SendMessage(hWnd, WM_SETICON, ICON_BIG, ByVal hIcon)
    lngRet = DrawMenuBar(hWnd)
End Sub


Public Sub RemoveFrame()
#If VBA7 Then
    Dim bitmask As LongPtr
    Dim hWnd As LongPtr
    Dim WindowStyle As LongPtr
#Else
    Dim bitmask As Long
    Dim hWnd As Long
    Dim WindowStyle As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    hWnd = GetForegroundWindow
    WindowStyle = GetWindowLong(hWnd, GWL_STYLE)
    bitmask = WindowStyle And (Not WS_DLGFRAME)
    Call SetWindowLong(hWnd, GWL_STYLE, bitmask)
    
End Sub

Public Sub flatBorder(uf As UserForm)
#If VBA7 Then
    Dim handle As LongPtr
    Dim lngWindow As LongPtr
    Dim lngHandle As LongPtr
    Dim bitmask As LongPtr
#Else
    Dim handle As Long
    Dim lngWindow As Long
    Dim lngHandle As Long
    Dim bitmask As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    handle = HWndOfUserForm(uf)
    lngHandle = GetWindowLong(handle, GWL_EXSTYLE)
    bitmask = lngHandle And (Not WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE Or WS_EX_DLGMODALFRAME Or WS_EX_STATICEDGE)
    Call SetWindowLong(handle, GWL_EXSTYLE, bitmask)
    Call UpdateWindow(handle)
    
End Sub


#If VBA7 Then
Public Sub removeTitleBar(ByVal hWnd As LongPtr)
#Else
Public Sub removeTitleBar(ByVal hWnd As Long)
#End If
    Call SetWindowLong(hWnd, GWL_STYLE, WS_POPUP)
End Sub


Public Sub makeUserFormTransparent(frm As Object, Optional color As Variant)
#If VBA7 Then
    Dim handle As LongPtr
#Else
    Dim handle As Long
#End If
    Dim bytOpacity As Byte
    '------------------------------------------------------------------------------------------------------
    
    handle = FindWindow(vbNullString, frm.caption)
    If IsMissing(color) Then color = vbWhite 'default to vbwhite
    bytOpacity = 100 ' variable keeping opacity setting
    
    SetWindowLong handle, GWL_EXSTYLE, GetWindowLong(handle, GWL_EXSTYLE) Or WS_EX_LAYERED
    'The following line makes only a certain color transparent so the
    ' background of the form and any object whose BackColor you've set to match
    ' vbColor (default vbWhite) will be transparent.
    frm.backColor = color
    SetLayeredWindowAttributes handle, color, bytOpacity, LWA_COLORKEY
    
End Sub

 
Public Sub HideTitleBarAndBorder(frm As Object)
#If VBA7 Then
    Dim lngWindow As LongPtr
    Dim handle As LongPtr
#Else
    Dim lngWindow As Long
    Dim handle As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    handle = FindWindow(vbNullString, frm.caption)
    lngWindow = GetWindowLong(handle, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong handle, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(handle, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong handle, GWL_EXSTYLE, lngWindow
    DrawMenuBar handle
    
End Sub





Private Function DoesWindowsHideFileExtensions() As Boolean
    Const KEY_NAME = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Const VALUE_NAME = "HideFileExt"
    '------------------------------------------------------------------------------------------------------
    Dim res As Long
#If VBA7 Then
    Dim RegKey As LongPtr
#Else
    Dim RegKey As Long
#End If
    Dim v As Long
    '------------------------------------------------------------------------------------------------------

    
    res = RegOpenKeyEx(HKey:=HKCU, _
                        lpSubKey:=KEY_NAME, _
                        ulOptions:=0&, _
                        samDesired:=KEY_ALL_ACCESS, _
                        phkResult:=RegKey)
    
    If res <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get the value of the "HideFileExt" named value.
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    res = RegQueryValueEx(HKey:=RegKey, _
                        lpValueName:=VALUE_NAME, _
                        lpReserved:=0&, _
                        LPType:=REG_DWORD, _
                        LPData:=v, _
                        lpcbData:=Len(v))
    
    If res <> ERROR_SUCCESS Then
        RegCloseKey RegKey
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Close the key and return the result.
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    RegCloseKey RegKey
    DoesWindowsHideFileExtensions = (v <> 0)

End Function


Function WindowCaption(w As Excel.window) As String
    Dim HideExt As Boolean
    Dim cap As String
    Dim pos As Long
    '------------------------------------------------------------------------------------------------------
    
    HideExt = DoesWindowsHideFileExtensions()
    cap = w.caption
    If HideExt = True Then
        pos = InStrRev(cap, ".")
        If pos > 0 Then
            cap = left(cap, pos - 1)
        End If
    End If
    
    WindowCaption = cap

End Function



#If VBA7 Then
    Function WindowHWnd(w As Excel.window) As LongPtr
        Dim appHWnd As LongPtr
        Dim deskHWnd As LongPtr
        Dim wHWnd As LongPtr
        Dim cap As String
#Else
    Function WindowHWnd(w As Excel.window) As Long
        Dim appHWnd As Long
        Dim deskHWnd As Long
        Dim wHWnd As Long
        Dim cap As String
#End If

    appHWnd = Application.hWnd
    deskHWnd = FindWindowEx(appHWnd, 0&, C_EXCEL_DESK_CLASSNAME, vbNullString)
    If deskHWnd > 0 Then
        cap = WindowCaption(w)
        wHWnd = FindWindowEx(deskHWnd, 0&, C_EXCEL_WINDOW_CLASSNAME, cap)
    End If
    WindowHWnd = wHWnd

End Function


#If VBA7 Then
Function WindowText(hWnd As LongPtr) As String
#Else
Function WindowText(hWnd As Long) As String
#End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WindowText
' This just wraps up GetWindowText.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim s As String
    Dim n As Long
    '------------------------------------------------------------------------------------------------------
    n = 255
    s = String$(n, vbNullChar)
    n = GetWindowText(hWnd, s, n)
    If n > 0 Then
        WindowText = left(s, n)
    Else
        WindowText = vbNullString
    End If
End Function



#If VBA7 Then
Function WindowClassName(hWnd As LongPtr) As String
#Else
Function WindowClassName(hWnd As Long) As String
#End If
    Dim s As String
    Dim n As Long
    '------------------------------------------------------------------------------------------------------
    
    n = 255
    s = String$(n, vbNullChar)
    n = GetClassName(hWnd, s, n)
    If n > 0 Then
        WindowClassName = VBA.left(s, n)
    Else
        WindowClassName = vbNullString
    End If

End Function


