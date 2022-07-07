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




Public Sub hideTitleBarAndBorder(frm As Object)
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



Public Sub setValidationLook(ctrl As MSForms.control, isValid As Boolean)
    With ctrl
        If isValid Then
            .backColor = VALID_BACK_COLOR
            .borderColor = VALID_BORDER_COLOR
            .ForeColor = VALID_FONT_COLOR
        Else
            .backColor = INVALID_BACK_COLOR
            .borderColor = INVALID_BORDER_COLOR
            .ForeColor = INVALID_FONT_COLOR
        End If
    End With
End Sub


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

Public Function isControlValid(ctrl As Variant) As Boolean
    Const BASE_NAME As String = "|"
    '------------------------------------------------------------------------------------------------------
    Dim name As String: name = BASE_NAME
    Dim captionLength As Long: captionLength = -1
    Dim parent As Object
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    If VBA.IsObject(ctrl) Then
        If Not ctrl Is Nothing Then
            If TypeOf ctrl Is MSForms.control Then
                name = ctrl.name
                isControlValid = (name <> BASE_NAME)
            ElseIf TypeOf ctrl Is MSForms.frame Then
                captionLength = VBA.Len(ctrl.caption)
                isControlValid = (captionLength >= 0)
            ElseIf TypeOf ctrl Is MSForms.UserForm Then
                captionLength = VBA.Len(ctrl.caption)
                isControlValid = (captionLength >= 0)
            End If
        End If
    End If
    
End Function

Public Sub clearMsFrame(frm As MSForms.frame)
    Dim ctrl As MSForms.control
    Dim controlsCounter As Long
    '------------------------------------------------------------------------------------------------------
    
    If UI.Forms.isControlValid(frm) Then
        controlsCounter = frm.controls.Count
        Do While controlsCounter
            Set ctrl = frm.controls.item(0)
            Call frm.controls.remove(ctrl.name)
            If frm.controls.Count = controlsCounter Then
                Call VBA.Err.Raise(1, description:="Error when trying to remove control " & ctrl.name)
            Else
                controlsCounter = frm.controls.Count
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
        If UI.Forms.isControlValid(ctrl) Then
            Call frame.controls.remove(ctrl.name)
        End If
    Next varControl
    
End Sub


Public Function getRight(ctrl As MSForms.control) As Single
    With ctrl
        getRight = .left + .width
    End With
End Function


Public Function getBottom(ctrl As MSForms.control) As Single
    With ctrl
        getBottom = .top + .height
    End With
End Function

Public Sub offsetControls(x As Single, y As Single, ParamArray controls() As Variant)
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
            Call customControl.offset(x, y)
        ElseIf TypeOf ctrl Is MSForms.control Then      '[MSForms.Control]
            Set msFormsControl = ctrl
            If isControlValid(msFormsControl) Then
                With msFormsControl
                    .left = F.Numerics.roundTo(.left + x, PIXEL_SIZE)
                    .top = F.Numerics.roundTo(.top + y, PIXEL_SIZE)
                End With
            End If
        End If
    Next ctrl
    
End Sub


Public Function calculateRightBottomPoint(frame As MSForms.frame, Optional onlyVisible As Boolean = True) As Coordinate
    Dim ctrl As MSForms.control
    Dim x As Single
    Dim y As Single
    '------------------------------------------------------------------------------------------------------
    
    For Each ctrl In frame.controls
        If ctrl.parent Is frame Then
            If onlyVisible = False Or ctrl.visible = True Then
                x = getRight(ctrl)
                y = getBottom(ctrl)
                
                If x > calculateRightBottomPoint.x Then
                    calculateRightBottomPoint.x = x
                End If
                
                If y > calculateRightBottomPoint.y Then
                    calculateRightBottomPoint.y = y
                End If
                
            End If
        End If
    Next ctrl

End Function


Public Sub fillListBox(listBox As MSForms.control, dict As Scripting.Dictionary)
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    If Not listBox Is Nothing Then
        If TypeOf listBox Is MSForms.listBox Or TypeOf listBox Is MSForms.comboBox Then
            Call listBox.clear
            For Each key In dict.keys
                Call listBox.addItem(key)
            Next key
        End If
    End If
    
End Sub


Public Sub locateControl(ctrl As MSForms.control, Optional left As Variant, Optional top As Variant)
    Const ADJ_OFFSET As Single = 0.083
    '------------------------------------------------------------------------------------------------------
    If isControlValid(ctrl) Then
        If Not VBA.IsMissing(left) Then
            ctrl.left = UI.Normalize.position(VBA.CDbl(left))
        End If
        If Not VBA.IsMissing(top) Then
            ctrl.top = UI.Normalize.position(VBA.CDbl(top)) + ADJ_OFFSET
        End If
    End If
End Sub


Public Sub resizeControl(ctrl As MSForms.control, Optional width As Variant, Optional height As Variant)
    If isControlValid(ctrl) Then
        If Not VBA.IsMissing(width) Then
            ctrl.width = UI.Normalize.size(VBA.CDbl(width))
        End If
        If Not VBA.IsMissing(height) Then
            ctrl.height = UI.Normalize.size(VBA.CDbl(height))
        End If
    End If
End Sub


Public Sub copySizeAndPosition(destinationControl As MSForms.control, sourceControl As MSForms.control)
    If Not isControlValid(destinationControl) Then
        'skip
    ElseIf Not isControlValid(sourceControl) Then
        'skip
    Else
        With destinationControl
            .width = sourceControl.width
            .height = sourceControl.height
            .top = sourceControl.top
            .left = sourceControl.left
        End With
    End If
End Sub


Public Function getControl(frm As MSForms.frame, controlName As String) As MSForms.control
    On Error Resume Next
    Set getControl = frm.controls(controlName)
End Function


Public Sub setFormOpacity(uf As MSForms.UserForm, opacity As Byte)
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
    
    result = SetLayeredWindowAttributes(frmHandle, 0, opacity, LWA_ALPHA)
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
    Dim windowCaption As String
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




Public Sub setControlProperty(ctrl As MSForms.control, propType As UIProp, propValue As Variant)
    If isControlValid(ctrl) Then
        On Error Resume Next
        Select Case propType.getName
            Case UI.props.backgroundColor.getName:          ctrl.backgroundColor = propValue
            Case UI.props.borderStyle.getName:              ctrl.borderStyle = propValue
            Case UI.props.ScrollBars.getName:               ctrl.ScrollBars = propValue
        End Select
    End If
End Sub



Public Sub printFrameControls(frame As MSForms.UserForm)
    Dim ctrl As MSForms.control
    Dim description As String
    '------------------------------------------------------------------------------------------------------
    If Not frame Is Nothing Then
        For Each ctrl In frame.controls
            description = ctrl.name & " | " & VBA.TypeName(ctrl) & " | Size: " & ctrl.width & "x" & ctrl.height & _
                        " | Position: (" & ctrl.left & ", " & ctrl.top & ")"
            On Error Resume Next
            description = description & " | Text: " & ctrl.caption
            description = description & " | Value: " & ctrl.value
            Debug.Print description
        Next ctrl
    End If
End Sub
