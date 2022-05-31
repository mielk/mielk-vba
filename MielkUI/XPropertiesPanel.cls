VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XPropertiesPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPanel

Private Const CLASS_NAME As String = "XPropertiesPanel"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_LABEL_WIDTH As Single = 100
Private Const DEFAULT_VALUE_WIDTH As Single = 400
Private Const DEFAULT_LABEL_VALUE_SPACE As Single = 0
Private Const DEFAULT_LINE_HEIGHT As Single = 24
'----------------------------------------------------------------------------------------------------------
Private pWindow As WTemplate
Private pListener As IListener
Private pParentFrame As MSForms.frame
Private pValidationFrame As MSForms.frame
Private pInnerFrame As MSForms.frame
Private pLines As Scripting.Dictionary
Private pTopBorder As MSForms.label
Private pBottomBorder As MSForms.label
'----------------------------------------------------------------------------------------------------------
Private pEditItem As IEditItem
Private pLinesDefinitions As VBA.Collection
'[UI properties] ------------------------------------------------------------------------------------------
Private pLabelWidth As Single
Private pValueWidth As Single
Private pLabelValueSpace As Single
Private pLineHeight As Single
Private pTop As Single
Private pLeft As Single
Private pBottom As Single
Private pRight As Single
Private pWidth As Single
Private pMinWidth As Single
Private pMaxWidth As Single
Private pHeight As Single
Private pMinHeight As Single
Private pMaxHeight As Single
Private pHorizontalPadding As Single
Private pVerticalPadding As Single
'[State] --------------------------------------------------------------------------------------------------
Private pName As String
Private pIndex As Long
Private pIsRendered As Boolean
Private pValidation As ValidationObject
'----------------------------------------------------------------------------------------------------------
Event AfterChange()
Event AfterValidationChange(validation As ValidationObject)
'----------------------------------------------------------------------------------------------------------





Private Sub Class_Initialize()
    Set pLines = F.Dictionaries.Create(False)
    pLabelWidth = DEFAULT_LABEL_WIDTH
    pValueWidth = DEFAULT_VALUE_WIDTH
    pLabelValueSpace = DEFAULT_LABEL_VALUE_SPACE
    pLineHeight = DEFAULT_LINE_HEIGHT
    
    '[Size]
    pWidth = -1
    pMinWidth = -1
    pMaxWidth = -1
    pHeight = -1
    pMinHeight = -1
    pMaxHeight = -1
    
End Sub



'[SETTERS]
Public Function setName(value As String) As XPropertiesPanel
    Set setName = Me
    pName = value
End Function

Public Function setListener(value As IListener) As XPropertiesPanel
    Set setListener = Me
    Set pListener = value
End Function

Public Function setWindow(value As WTemplate) As XPropertiesPanel
    Set setWindow = Me
    Set pWindow = value
End Function

Public Function setParentFrame(value As MSForms.frame) As XPropertiesPanel
    Set setParentFrame = Me
    Set pParentFrame = value
End Function

Public Function setValidationFrame(value As MSForms.frame) As XPropertiesPanel
    Set setValidationFrame = Me
    Set pValidationFrame = value
End Function

Public Function setEditItem(value As IEditItem) As XPropertiesPanel
    Set setEditItem = Me
    Set pEditItem = value
End Function

Public Function setLabelWidth(value As Single) As XPropertiesPanel
    Set setLabelWidth = Me
    pLabelWidth = value
End Function

Public Function setValueWidth(value As Single) As XPropertiesPanel
    Set setValueWidth = Me
    pValueWidth = value
End Function

Public Function setLinesDefinitions(value As VBA.Collection) As XPropertiesPanel
    Set setLinesDefinitions = Me
    Set pLinesDefinitions = value
End Function

Public Function setTop(value As Single) As XPropertiesPanel
    Set setTop = Me
    pTop = value
End Function

Public Function setLeft(value As Single) As XPropertiesPanel
    Set setLeft = Me
    pLeft = value
End Function

Public Function setIndex(value As Long) As XPropertiesPanel
    Set setIndex = Me
    pIndex = value
End Function

Public Function setUiProperty(prop As UIProp, value As Variant) As XPropertiesPanel
    Set setUiProperty = Me
    Select Case prop.getId
        Case UI.props.top.getId:                    Call setTop(VBA.CSng(value))
        Case UI.props.left.getId:                   Call setLeft(VBA.CSng(value))
        Case UI.props.width.getId:                  pWidth = VBA.CSng(value)
        Case UI.props.MinWidth.getId:               pMinWidth = VBA.CSng(value)
        Case UI.props.maxWidth.getId:               pMaxWidth = VBA.CSng(value)
        Case UI.props.height.getId:                 pHeight = VBA.CSng(value)
        Case UI.props.minHeight.getId:              pMinHeight = VBA.CSng(value)
        Case UI.props.maxHeight.getId:              pMaxHeight = VBA.CSng(value)
        Case UI.props.padding.getId:
                                                    pHorizontalPadding = VBA.CSng(value)
                                                    pVerticalPadding = VBA.CSng(value)
        Case UI.props.paddingHorizontal.getId:      pHorizontalPadding = VBA.CSng(value)
        Case UI.props.paddingVertical.getId:        pVerticalPadding = VBA.CSng(value)
    End Select
    
    If pIsRendered Then
        If prop.isLayoutProperty Then Call arrangeLines
        If prop.isPositionProperty Then Call updatePosition
        If prop.isSizeProperty Then Call updateSize
    End If
    
End Function



'[GETTERS]
Public Function getName() As String
    getName = pName
End Function

Public Function getWindow() As WTemplate
    Set getWindow = pWindow
End Function

Public Function getInnerFrame() As MSForms.frame
    Set getInnerFrame = pInnerFrame
End Function

Public Function getValidationFrame() As MSForms.frame
    Set getValidationFrame = pValidationFrame
End Function

Public Function getEditItem() As IEditItem
    Set getEditItem = pEditItem
End Function

Public Function getBottom() As Single
    getBottom = pBottom
End Function

Public Function getWidth() As Single
    getWidth = pInnerFrame.width
End Function

Public Function getLabelWidth() As Single
    getLabelWidth = pLabelWidth
End Function

Public Function getValueWidth() As Single
    getValueWidth = pValueWidth
End Function

Public Function getLabelValueSpace() As Single
    getLabelValueSpace = pLabelValueSpace
End Function

Public Function getLineHeight() As Single
    getLineHeight = pLineHeight
End Function

Public Function getIndex() As Long
    getIndex = pIndex
End Function

Public Function getUiProperty(prop As UIProp) As Variant
    If UI.Forms.isControlValid(pInnerFrame) Then
        With pInnerFrame
            Select Case prop.getId
                Case UI.props.top.getId:                getUiProperty = .top
                Case UI.props.left.getId:               getUiProperty = .left
                Case UI.props.right.getId:              getUiProperty = .left + .width
                Case UI.props.bottom.getId:             getUiProperty = .top + .height
            End Select
        End With
    End If
End Function




'[RENDERING]
Public Sub render()
    If Not pIsRendered Then
        Call insertInnerFrame
        Call insertLines
        Call arrangeLines
        Call updatePosition
        Call updateValidation
        pIsRendered = True
    End If
End Sub

Public Sub insertInnerFrame()
    Set pInnerFrame = pParentFrame.controls.Add(CREATE_FRAME_ID)
    With pInnerFrame
        .borderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Public Sub insertLines()
    Dim dict As Scripting.Dictionary
    Dim line As xxPropertyLine
    '----------------------------------------------------------------------------------------------------------
    
    Set pLines = F.Dictionaries.Create(False)
    For Each dict In pLinesDefinitions
        Set line = UI.Create.PropertyLine(Me, dict)
        Call pLines.Add(line.getProperty, line)
        Call line.render
    Next dict
    
End Sub

Public Sub insertBorder(index As XlBordersIndex, color As Long)
    Dim dict As Scripting.Dictionary
    Dim line As xxPropertyLine
    '----------------------------------------------------------------------------------------------------------

    If index = xlEdgeTop Then
        Call updateBorderVariable(pTopBorder, color)
    ElseIf index = xlEdgeBottom Then
        Call updateBorderVariable(pBottomBorder, color)
    End If

    Call updateBordersSizeAndPosition

End Sub

Private Sub updateBorderVariable(ByRef borderLabel As MSForms.label, color As Long)
    If borderLabel Is Nothing Then
        Set borderLabel = pInnerFrame.controls.Add(CREATE_LABEL_ID)
    End If
    With borderLabel
        .borderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleOpaque
        .BackColor = color
        .height = PIXEL_SIZE
    End With
End Sub

Private Sub updateBordersSizeAndPosition()
    If UI.Forms.isControlValid(pTopBorder) Then
        With pTopBorder
            .left = 0
            .top = PIXEL_SIZE
            .width = pInnerFrame.InsideWidth
            .height = PIXEL_SIZE
        End With
    End If
    
    If UI.Forms.isControlValid(pBottomBorder) Then
        With pBottomBorder
            .left = 0
            .top = pParentFrame.InsideHeight - PIXEL_SIZE
            .width = pParentFrame.InsideWidth
            .height = PIXEL_SIZE
        End With
    End If
    
End Sub



Private Sub arrangeLines()
    Dim key As Variant
    Dim line As xxPropertyLine
    Dim top As Single:              top = pVerticalPadding
    Dim right As Single
    '----------------------------------------------------------------------------------------------------------
    
    If F.Collections.isNonEmptyCollection(pLinesDefinitions) Then
        For Each key In pLines.Keys
            Set line = pLines.Keys
            Call line.Move(pHorizontalPadding, top)
            right = line.getRight
            If right > pRight Then pRight = right
            top = top + pLineHeight
        Next key
        pBottom = top
    End If
    
    Call updateSize
    Call updateScrollBars
    
    
End Sub

Private Sub updateSize()
    Dim lastPoint As Coordinate
    '----------------------------------------------------------------------------------------------------------
    
    lastPoint = UI.Forms.calculateRightBottomPoint(pInnerFrame, True)
    
    With pInnerFrame
        '[Width]
        If pWidth >= 0 Then
            .width = pWidth
        ElseIf pMinWidth > 0 And lastPoint.x < pMinWidth Then
            .width = pMinWidth
        ElseIf pMaxWidth > 0 And lastPoint.x < pMaxWidth Then
            .width = pMaxWidth
        Else
            .width = lastPoint.x + pHorizontalPadding
        End If
        
        '[Height]
        If pHeight >= 0 Then
            .height = pHeight
        ElseIf pMinHeight > 0 And lastPoint.y < pMinHeight Then
            .height = pMinHeight
        ElseIf pMaxHeight > 0 And lastPoint.y < pMaxHeight Then
            .height = pMaxHeight
        Else
            .height = lastPoint.y + pVerticalPadding
        End If
        
    End With
    
    Call updateScrollBars
    Call updateBordersSizeAndPosition
    
End Sub

Private Sub updateScrollBars()
    Dim lastPoint As Coordinate
    '----------------------------------------------------------------------------------------------------------
    
    lastPoint = UI.Forms.calculateRightBottomPoint(pInnerFrame, True)
    With pInnerFrame
        .ScrollBars = fmScrollBarsNone
        
        '[Vertical scroll bar]
        .ScrollHeight = lastPoint.y
        If lastPoint.y > .InsideHeight Then
            .ScrollBars = .ScrollBars + fmScrollBarsVertical
            .scrollTop = 0
        End If
        
        '[Horizontal scroll bar]
        .ScrollWidth = lastPoint.x
        If lastPoint.x > .InsideWidth Then
            .ScrollBars = .ScrollBars + fmScrollBarsHorizontal
            .ScrollLeft = 0
        End If
        
    End With
     
End Sub

Private Sub updatePosition()
    If UI.Forms.isControlValid(pInnerFrame) Then
        With pInnerFrame
            .left = pLeft
            .top = pTop
        End With
    End If
End Sub


'[VALIDATION]
Public Sub updateValidation()
    Dim key As Variant
    Dim line As xxPropertyLine
    Dim status As Boolean
    Dim previousStatus As Boolean
    Dim validation As ValidationObject
    Dim hasValidationErrors As Boolean
    '----------------------------------------------------------------------------------------------------------
    
    If Not pValidation Is Nothing Then
        previousStatus = pValidation.getStatus
    End If
    
    status = True
    For Each key In pLines.Keys
        Set line = pLines.item(key)
        Set validation = line.getValidation
        If validation Is Nothing Then
            hasValidationErrors = True
        ElseIf Not validation.getStatus Then
            hasValidationErrors = True
        End If
        
        If hasValidationErrors Then
            status = False
            Exit For
        End If
        
    Next key
    Set pValidation = F.Create.ValidationObject(status)
    
    If previousStatus <> status Then
        RaiseEvent AfterValidationChange(pValidation)
    End If
    
    RaiseEvent AfterChange
    
End Sub

Public Function getValidation() As ValidationObject
    Set getValidation = pValidation
End Function

Public Sub validateLineByProperty(propertyName As String)
    Dim line As xxPropertyLine
    '----------------------------------------------------------------------------------------------------------
    
    Set line = F.Dictionaries.getObject(pLines, propertyName)
    If Not line Is Nothing Then
        Call line.validate
    End If
    
End Sub







Public Sub setFocus()
    Dim line As xxPropertyLine
    '----------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    Set line = pLines.items(0)
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If Not line Is Nothing Then
        Call line.setFocus
    End If
    
    
ErrHandler:
    Stop
    
End Sub



Public Sub triggerConfirm()
    If pValidation.getStatus Then
        Call pWindow.triggerConfirm
    End If
End Sub








'[Implementation of IPanel interface]

Private Function IPanel_setName(value As String) As IPanel
    Set IPanel_setName = setName(value)
End Function

Private Function IPanel_getName() As String
    IPanel_getName = getName
End Function


Private Function IPanel_setIndex(value As Long) As IPanel
    Set IPanel_setIndex = setIndex(value)
End Function

Private Function IPanel_getIndex() As Long
    IPanel_getIndex = getIndex
End Function


Private Sub IPanel_render()
    Call render
End Sub

Private Function IPanel_setParentFrame(value As MSForms.IOptionFrame) As IPanel
    Set IPanel_setParentFrame = setParentFrame(value)
End Function

Private Function IPanel_setUiProperty(prop As MielkCore.UIProp, value As Variant) As IPanel
    Set IPanel_setUiProperty = setUiProperty(prop, value)
End Function

Private Function IPanel_getUiProperty(prop As MielkCore.UIProp) As Variant
    Call F.Variables.assign(IPanel_getUiProperty, getUiProperty(prop))
End Function

Private Sub IPanel_insertBorder(index As XlBordersIndex, color As Long)
    Call insertBorder(index, color)
End Sub

Private Sub IPanel_addValidationLabel(value As UIValidationLabel)
    Stop
End Sub

Private Function IPanel_setListener(value As IListener) As IPanel
    Set IPanel_setListener = setListener(value)
End Function

Private Function IPanel_getValidation() As MielkCore.ValidationObject
    Set IPanel_getValidation = getValidation
End Function

Private Function IPanel_getValue() As Variant
    Set IPanel_getValue = pEditItem
End Function
