VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FControls"
'----------------------------------------------------------------------------------------------------------

Public Function isContainer(value As Variant) As Boolean
    If VBA.IsObject(value) Then
        If TypeOf value Is MSForms.UserForm Then
            isContainer = True
        ElseIf TypeOf value Is MSForms.frame Then
            isContainer = True
        End If
    End If
End Function


Public Function isValid(ctrl As Variant) As Boolean
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
                isValid = (name <> BASE_NAME)
            ElseIf TypeOf ctrl Is MSForms.frame Then
                captionLength = VBA.Len(ctrl.caption)
                isValid = (captionLength >= 0)
            ElseIf TypeOf ctrl Is MSForms.UserForm Then
                captionLength = VBA.Len(ctrl.caption)
                isValid = (captionLength >= 0)
            End If
        End If
    End If
    
End Function


Public Sub locate(ctrl As MSForms.control, Optional left As Variant, Optional top As Variant)
    Const ADJ_OFFSET As Single = 0.083
    '------------------------------------------------------------------------------------------------------
    If UI.controls.isValid(ctrl) Then
        If Not VBA.IsMissing(left) Then
            ctrl.left = UI.Normalize.position(VBA.CDbl(left))
        End If
        If Not VBA.IsMissing(top) Then
            ctrl.top = UI.Normalize.position(VBA.CDbl(top)) + ADJ_OFFSET
        End If
    End If
End Sub

Public Sub locateByDictionary(ctrl As Object, properties As Scripting.Dictionary, _
                              parentFrame As MSForms.frame, Optional padding As Single = 0)
    Dim left As Variant
    Dim right As Variant
    Dim top As Variant
    Dim bottom As Variant
    Dim width As Single
    Dim height As Single
    '------------------------------------------------------------------------------------------------------
    
    Call F.dictionaries.assignValueFromDictToVariable(properties, left, "position.left")
    Call F.dictionaries.assignValueFromDictToVariable(properties, right, "position.right")
    Call F.dictionaries.assignValueFromDictToVariable(properties, top, "position.top")
    Call F.dictionaries.assignValueFromDictToVariable(properties, bottom, "position.bottom")
    
    If TypeOf ctrl Is MSForms.label Then
        width = ctrl.width
        height = ctrl.height
    ElseIf TypeOf ctrl Is IControl Then
        width = ctrl.getWidth
        height = ctrl.getHeight
    End If
    
    With ctrl
        left = calculateX(left, right, width, parentFrame.InsideWidth, padding)
        top = calculateY(top, bottom, height, parentFrame.InsideHeight, padding)
    End With
    
    
    If TypeOf ctrl Is MSForms.label Then
        ctrl.left = left
        ctrl.top = top
    ElseIf TypeOf ctrl Is IControl Then
        Call ctrl.Move(left, top)
    End If
    
End Sub


Public Function calculateX(left As Variant, right As Variant, width As Single, _
                                    parentWidth As Single, Optional padding As Single = 0) As Single
    If Not VBA.IsEmpty(left) Then
        If VBA.IsNumeric(left) Then
            calculateX = VBA.CSng(left) + padding
        ElseIf left = "middle" Then
            calculateX = (parentWidth - width) / 2
        ElseIf left = "left" Then
            calculateX = padding
        ElseIf left = "right" Then
            calculateX = parentWidth - padding - width
        End If
    ElseIf Not VBA.IsEmpty(right) Then
        If VBA.IsNumeric(right) Then
            calculateX = parentWidth - padding - width - VBA.CSng(right)
        ElseIf left = "middle" Then
            calculateX = (parentWidth - width) / 2
        ElseIf left = "left" Then
            calculateX = padding
        ElseIf left = "right" Then
            calculateX = parentWidth - padding - width
        End If
    Else
        calculateX = padding
    End If
End Function

Public Function calculateY(top As Variant, bottom As Variant, height As Single, _
                                    parentHeight As Single, Optional padding As Single = 0) As Single
    If Not VBA.IsEmpty(top) Then
        If VBA.IsNumeric(top) Then
            calculateY = VBA.CSng(top) + padding
        ElseIf top = "middle" Then
            calculateY = (parentHeight - height) / 2
        ElseIf top = "top" Then
            calculateY = padding
        ElseIf top = "bottom" Then
            calculateY = parentHeight - padding - height
        End If
    ElseIf Not VBA.IsEmpty(bottom) Then
        If VBA.IsNumeric(bottom) Then
            calculateY = parentHeight - padding - height - VBA.CSng(bottom)
        ElseIf top = "middle" Then
            calculateY = (parentHeight - height) / 2
        ElseIf top = "top" Then
            calculateY = padding
        ElseIf top = "bottom" Then
            calculateY = parentHeight - padding - height
        End If
    Else
        calculateY = padding
    End If
End Function

Public Sub locateControlsInLine(controls As VBA.Collection, space As Single)
    Dim container As MSForms.control
    Dim ctrl As MSForms.control
    Dim totalWidth As Single
    Dim left As Single
    '------------------------------------------------------------------------------------------------------

    For Each ctrl In controls
        If container Is Nothing Then Set container = ctrl.parent
        If Not container Is ctrl.parent Then GoTo DifferentContainerException
    Next ctrl

    For Each ctrl In controls
        If totalWidth > 0 Then totalWidth = totalWidth + space
        totalWidth = totalWidth + ctrl.width
    Next ctrl

    left = (container.width - totalWidth) / 2
    For Each ctrl In controls
        ctrl.left = left
        left = left + ctrl.width + space
    Next ctrl

'===========================================================================================================
ExitPoint:
    Exit Sub

'----------------------------------------------------------------------------------------------------------
DifferentContainerException:
    VBA.Err.Raise 1
    GoTo ExitPoint
End Sub








Public Sub resize(ctrl As MSForms.control, Optional width As Variant, Optional height As Variant)
    If UI.controls.isValid(ctrl) Then
        If Not VBA.IsMissing(width) Then
            ctrl.width = UI.Normalize.size(VBA.CDbl(width))
        End If
        If Not VBA.IsMissing(height) Then
            ctrl.height = UI.Normalize.size(VBA.CDbl(height))
        End If
    End If
End Sub









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



Public Sub copySizeAndPosition(destinationControl As MSForms.control, sourceControl As MSForms.control)
    If Not isValid(destinationControl) Then
        'skip
    ElseIf isValid(sourceControl) Then
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

Public Sub setProperty(ctrl As MSForms.control, propType As UIProp, propValue As Variant)
    If isValid(ctrl) Then
        On Error Resume Next
        Select Case propType.getName
            Case UI.props.backgroundColor.getName:          ctrl.backgroundColor = propValue
            Case UI.props.borderStyle.getName:              ctrl.borderStyle = propValue
            Case UI.props.ScrollBars.getName:               ctrl.ScrollBars = propValue
        End Select
    End If
End Sub

