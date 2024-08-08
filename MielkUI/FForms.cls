VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FForms"
'----------------------------------------------------------------------------------------------------------



'Method to convert all pixels with the given [color] on the given form [frm] to transparent pixels.
'Source:    https://wellssr.com/vba/2017/excel/transaprent-userform-background-with-vba/
Public Sub makeUserFormTransparent(frm As Object, Optional color As Variant = TRANSPARENCY_LAYER_COLOR)


#If Win64 Then
    Dim frmHandle As LongPtr
#Else
    Dim frmHandle As Long
#End If
    Dim bytOpacity As Byte:                         bytOpacity = 100
    '------------------------------------------------------------------------------------------------------
    
    frmHandle = getWindowHandle(frm)
    
    SetWindowLong frmHandle, GWL_EXSTYLE, GetWindowLong(frmHandle, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes frmHandle, color, bytOpacity, LWA_COLORKEY
    
End Sub




Public Sub HideTitleBarAndBorder(frm As Object)
#If Win64 Then
    Dim frmHandle As LongPtr
    Dim lngWindow As LongPtr
#Else
    Dim frmHandle As Long
    Dim lngWindow As Long
#End If
    '------------------------------------------------------------------------------------------------------
    
    frmHandle = getWindowHandle(frm)
    lngWindow = GetWindowLong(frmHandle, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong frmHandle, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(frmHandle, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong frmHandle, GWL_EXSTYLE, lngWindow
    DrawMenuBar frmHandle
 
End Sub



#If Win64 Then
    Public Function getWindowHandle(window As Object) As LongPtr
#Else
    Public Function getWindowHandle(window As Object) As Long
#End If
    
    If window Is Nothing Then
        getWindowHandle = 0
    ElseIf TypeOf window Is MSForms.UserForm Then
        getWindowHandle = FindWindow(vbNullString, window.caption)
    End If
    
End Function


Public Function getTotalValidationStatus(validations As Scripting.Dictionary) As Boolean
    Dim key As Variant
    Dim validation As ValidationObject
    '------------------------------------------------------------------------------------------------------
    
    getTotalValidationStatus = True
    For Each key In validations.keys
        Set validation = validations.item(key)
        If Not validation.getStatus Then
            getTotalValidationStatus = False
            Exit For
        End If
    Next key
    
End Function


Public Sub clearMsFrame(frm As MSForms.frame)
    Dim ctrl As MSForms.control
    Dim controlsCounter As Long
    '------------------------------------------------------------------------------------------------------
    
    If UI.controls.isValid(frm) Then
        controlsCounter = frm.controls.count
        Do While controlsCounter
            Set ctrl = frm.controls.item(0)
            Call frm.controls.remove(ctrl.name)
            If frm.controls.count = controlsCounter Then
                Call VBA.Err.Raise(1, description:="Error when trying to remove control " & ctrl.name)
            Else
                controlsCounter = frm.controls.count
            End If
        Loop
    End If
    
End Sub


Public Sub removeControls(frame As MSForms.frame, ParamArray controls() As Variant)
    Dim varControl As Variant
    Dim ctrl As MSForms.control
    '------------------------------------------------------------------------------------------------------
    
    For Each varControl In controls
        On Error Resume Next
        Set ctrl = varControl
        If UI.controls.isValid(ctrl) Then
            Call frame.controls.remove(ctrl.name)
        End If
    Next varControl
    
End Sub






Public Sub offsetControls(x As Single, Y As Single, ParamArray controls() As Variant)
    Dim ctrl As Variant
    Dim customControl As IControl
    Dim msFormsControl As MSForms.control
    '------------------------------------------------------------------------------------------------------
    
    For Each ctrl In controls
        If Not VBA.IsObject(ctrl) Then
            ' skip
        ElseIf ctrl Is Nothing Then
            ' skip
        ElseIf TypeOf ctrl Is IControl Then             '[IControl]
            Set customControl = ctrl
            Call customControl.offset(x, Y)
        ElseIf TypeOf ctrl Is MSForms.control Then      '[MSForms.Control]
            Set msFormsControl = ctrl
            If UI.controls.isValid(msFormsControl) Then
                With msFormsControl
                    .left = F.Numerics.roundTo(.left + x, PIXEL_SIZE)
                    .top = F.Numerics.roundTo(.top + Y, PIXEL_SIZE)
                End With
            End If
        End If
    Next ctrl
    
End Sub


Public Function calculateRightBottomPoint(container As Object, Optional onlyVisible As Boolean = True) As Coordinate
    Dim ctrl As MSForms.control
    Dim x As Single
    Dim Y As Single
    '------------------------------------------------------------------------------------------------------
    
    For Each ctrl In container.controls
        If ctrl.parent Is container Then
            If onlyVisible = False Or ctrl.visible = True Then
                x = UI.controls.getRight(ctrl)
                Y = UI.controls.getBottom(ctrl)
                
                If x > calculateRightBottomPoint.x Then
                    calculateRightBottomPoint.x = x
                End If
                
                If Y > calculateRightBottomPoint.Y Then
                    calculateRightBottomPoint.Y = Y
                End If
                
            End If
        End If
    Next ctrl

End Function

Public Function getControl(frm As MSForms.frame, controlName As String) As MSForms.control
    On Error Resume Next
    Set getControl = frm.controls(controlName)
End Function


Public Sub SetFormOpacity(uf As MSForms.UserForm, Opacity As Byte)
#If VBA7 Then
    Dim frmHandle As LongPtr
    Dim winLng As LongPtr
#Else
    Dim frmHandle As Long
    Dim winLng As Long
#End If
    Dim result As Variant
    '------------------------------------------------------------------------------------------------------
    
    frmHandle = handleOfUserForm(uf)
    If frmHandle = 0 Then Exit Sub
    
    winLng = GetWindowLong(frmHandle, GWL_EXSTYLE)
    If winLng = 0 Then Exit Sub
    
    result = SetWindowLong(frmHandle, GWL_EXSTYLE, winLng Or WS_EX_LAYERED)
    If result = 0 Then Exit Sub
    
    result = SetLayeredWindowAttributes(frmHandle, 0, Opacity, LWA_ALPHA)
    If result = 0 Then Exit Sub
    
End Sub




#If Win64 Then
Public Function handleOfUserForm(uf As MSForms.UserForm) As LongPtr
    Dim applicationHandle As LongPtr
    Dim deskHandle As LongPtr
    Dim windowHandle As LongPtr
    Dim formHandle As LongPtr
#Else
Public Function handleOfUserForm(uf As MSForms.UserForm) As Long
    Dim applicationHandle As Long
    Dim deskHandle As Long
    Dim windowHandle As Long
    Dim formHandle As Long
#End If
    Dim caption As String
    Dim WindowCaption As String
    '------------------------------------------------------------------------------------------------------
    
    caption = uf.caption
    
    formHandle = FindWindow(C_USERFORM_CLASSNAME, caption)
    If formHandle <> 0 Then
        handleOfUserForm = formHandle
        Exit Function
    End If
    
    applicationHandle = Excel.Application.hWnd
    formHandle = FindWindowEx(applicationHandle, 0&, C_USERFORM_CLASSNAME, caption)
    If formHandle <> 0 Then
        handleOfUserForm = formHandle
        Exit Function
    End If

End Function





Public Sub printFrameControls(frame As MSForms.UserForm)
    Dim ctrl As MSForms.control
    Dim description As String
    '------------------------------------------------------------------------------------------------------
    If Not frame Is Nothing Then
        For Each ctrl In frame.controls
            description = ctrl.name & " | " & VBA.TypeName(ctrl) & " | Size: " & ctrl.width & "x" & ctrl.height & _
                        " | Position: (" & ctrl.left & ", " & ctrl.top & ") | Visible: " & ctrl.visible
            On Error Resume Next
            description = description & " | Text: " & ctrl.caption
            description = description & " | Value: " & ctrl.value
            Debug.Print description
        Next ctrl
    End If
End Sub






'[EXTREME CONTROLS]
Public Function getMostTopControl(container As Object, Optional onlyVisible As Boolean) As MSForms.control
    Dim ctrl As MSForms.control
    Dim top As Single
    Dim minValue As Single
    '-------------------------------------------------------------------------------
    
    For Each ctrl In container.controls
        If Not onlyVisible Or ctrl.visible = True Then
            top = ctrl.top
            If top < minValue Then
                Set getMostTopControl = ctrl
                minValue = top
            End If
        End If
    Next ctrl
End Function


Public Function getMostBottomControl(container As Object, Optional onlyVisible As Boolean) As MSForms.control
    Dim ctrl As MSForms.control
    Dim bottom As Single
    Dim maxValue As Single
    '-------------------------------------------------------------------------------
    
    For Each ctrl In container.controls
        If Not onlyVisible Or ctrl.visible = True Then
            bottom = UI.controls.getBottom(ctrl)
            If bottom > maxValue Then
                Set getMostBottomControl = ctrl
                maxValue = bottom
            End If
        End If
    Next ctrl
End Function

Public Function getMostLeftControl(container As Object, Optional onlyVisible As Boolean) As MSForms.control
    Dim ctrl As MSForms.control
    Dim left As Single
    Dim minValue As Single
    '-------------------------------------------------------------------------------
    
    For Each ctrl In container.controls
        If Not onlyVisible Or ctrl.visible = True Then
            left = ctrl.left
            If left < minValue Then
                Set getMostLeftControl = ctrl
                minValue = left
            End If
        End If
    Next ctrl
End Function


Public Function getMostRightControl(container As Object, Optional onlyVisible As Boolean) As MSForms.control
    Dim ctrl As MSForms.control
    Dim right As Single
    Dim maxValue As Single
    '-------------------------------------------------------------------------------
    
    For Each ctrl In container.controls
        If Not onlyVisible Or ctrl.visible = True Then
            right = UI.controls.getRight(ctrl)
            If right > maxValue Then
                Set getMostRightControl = ctrl
                maxValue = right
            End If
        End If
    Next ctrl
End Function





'[EXTREMA COORDINATES]
Public Function getMostRightX(container As Object, Optional onlyVisible As Boolean) As Single
    Dim ctrl As MSForms.control
    Dim right As Single
    '-------------------------------------------------------------------------------
    For Each ctrl In container.controls
        If Not onlyVisible Or ctrl.visible = True Then
            right = UI.controls.getRight(ctrl)
            If right > getMostRightX Then
                getMostRightX = right
            End If
        End If
    Next ctrl
End Function

Public Function getMostLeftX(container As Object, Optional onlyVisible As Boolean) As Single
    Dim ctrl As MSForms.control
    Dim left As Single
    '-------------------------------------------------------------------------------
    left = container.width
    For Each ctrl In container.controls
        If Not onlyVisible Or ctrl.visible = True Then
            left = ctrl.left
            If left < getMostLeftX Then getMostLeftX = left
        End If
    Next ctrl
End Function

Public Function getMostBottomY(container As Object, Optional onlyVisible As Boolean) As Single
    Dim ctrl As MSForms.control
    Dim bottom As Single
    '-------------------------------------------------------------------------------
    For Each ctrl In container.controls
        If Not onlyVisible Or ctrl.visible = True Then
            bottom = UI.controls.getBottom(ctrl)
            If bottom > getMostBottomY Then getMostBottomY = bottom
        End If
    Next ctrl
End Function

Public Function getMostTopY(container As Object, Optional onlyVisible As Boolean) As Single
    Dim ctrl As MSForms.control
    Dim top As Single
    '-------------------------------------------------------------------------------
    top = container.height
    For Each ctrl In container.controls
        If Not onlyVisible Or ctrl.visible = True Then
            top = UI.controls.getBottom(ctrl)
            If top < getMostTopY Then getMostTopY = top
        End If
    Next ctrl
End Function

'
'
'Public Sub centerUserForm(form As MSForms.UserForm)
'    Dim screenArea As area
'    Dim left As Single
'    Dim top As Single
'    '----------------------------------------------------------------------------------------------------------
'
'    screenArea = F.Screen.getSize
'
'    '[Calculate left and top coordinate]
'    With screenArea
'        left = (.width - form.InsideWidth) / 2
'        top = (.height - form.InsideHeight) / 2
'    End With
'
'    '[Apply coordinates calculated above to this user form]
'    With form
'        .left = left
'        .top = top
'    End With
'
'
'End Sub
