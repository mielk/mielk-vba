VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XOptionsPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPanel
Implements IControl

Private Const CLASS_NAME As String = "XOptionsPanel"
'[Captions] -----------------------------------------------------------------------------------------------
Private Const WARNING_TOO_LESS_ITEMS_SELECTED As String = "OptionsContainer.Warnings.TooLessItemsSelected"
Private Const WARNING_TOO_MANY_ITEMS_SELECTED As String = "OptionsContainer.Warnings.TooManyItemsSelected"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_CHECKBOX_WIDTH As Single = 15
Private Const DEFAULT_LABEL_WIDTH As Single = 400
Private Const DEFAULT_CHECKBOX_LABEL_SPACE As Single = 1.5
Private Const DEFAULT_LINE_HEIGHT As Single = 18
Private Const DEFAULT_MIN_ITEMS_SELECTED As Long = 0
Private Const DEFAULT_MAX_ITEMS_SELECTED As Long = 0
'----------------------------------------------------------------------------------------------------------
Private pParentFrame As MSForms.frame
Private pFrame As MSForms.frame
Private pListener As IListener
Private pValidationLabel As UIValidationLabel
Private pLines As Scripting.Dictionary
Private pTopBorder As MSForms.label
Private pBottomBorder As MSForms.label
'[UI properties] ------------------------------------------------------------------------------------------
Private pCheckboxWidth As Single
Private pLabelWidth As Single
Private pCheckboxValueSpace As Single
Private pLineHeight As Single
Private pTop As Single
Private pLeft As Single
Private pBottom As Single
Private pRight As Single
Private pWidth As Single
Private pHeight As Single
Private pMaxHeight As Single
Private pMinHeight As Single
Private pMargin As Single
'[State] --------------------------------------------------------------------------------------------------
Private pName As String
Private pIndex As Long
Private pItems As Scripting.Dictionary
Private pInitialSelection As Scripting.Dictionary
Private pFixedItems As Scripting.Dictionary
Private pMinItemsSelected As Long
Private pMaxItemsSelected As Long
Private pValidation As ValidationObject
Private pIsRendered As Boolean
Private pDisableEvents As Boolean
'----------------------------------------------------------------------------------------------------------
Event AfterChange()
Event AfterResize()
Event AfterMoved()
Event AfterValidationChange(validation As ValidationObject)
'----------------------------------------------------------------------------------------------------------





Private Sub Class_Initialize()
    Set pItems = F.dictionaries.Create(False)
    Set pLines = F.dictionaries.Create(False)
    pCheckboxWidth = DEFAULT_CHECKBOX_WIDTH
    pLabelWidth = DEFAULT_LABEL_WIDTH
    pCheckboxValueSpace = DEFAULT_CHECKBOX_LABEL_SPACE
    pLineHeight = DEFAULT_LINE_HEIGHT
    pMinItemsSelected = DEFAULT_MIN_ITEMS_SELECTED
    pMaxItemsSelected = DEFAULT_MAX_ITEMS_SELECTED
End Sub



'[SETTERS]
Public Function setName(value As String) As XOptionsPanel
    Set setName = Me
    pName = value
End Function

Public Function setListener(value As IListener) As XOptionsPanel
    Set setListener = Me
    Set pListener = value
End Function

Public Function setMargin(value As Single) As XOptionsPanel
    Set setMargin = Me
    pMargin = value
End Function

Public Function setParentFrame(value As MSForms.frame) As XOptionsPanel
    Set setParentFrame = Me
    Set pParentFrame = value
End Function

Public Function setCheckboxWidth(value As Single) As XOptionsPanel
    Set setCheckboxWidth = Me
    pCheckboxWidth = value
End Function

Public Function setLabelWidth(value As Single) As XOptionsPanel
    Set setLabelWidth = Me
    pLabelWidth = value
End Function

Public Function setItems(value As Scripting.Dictionary) As XOptionsPanel
    Set setItems = Me
    Set pItems = value
End Function

Public Function setInitialSelection(value As Scripting.Dictionary) As XOptionsPanel
    Set setInitialSelection = Me
    Set pInitialSelection = value
End Function

Public Function setFixedItems(value As Scripting.Dictionary) As XOptionsPanel
    Set setFixedItems = Me
    Set pFixedItems = value
End Function

Public Function setMinItemsSelected(value As Long) As XOptionsPanel
    Set setMinItemsSelected = Me
    pMinItemsSelected = value
End Function

Public Function setMaxItemsSelected(value As Long) As XOptionsPanel
    Set setMaxItemsSelected = Me
    pMaxItemsSelected = value
End Function

Public Function setTop(value As Single) As XOptionsPanel
    Set setTop = Me
    pTop = value
    Call adjustPosition
End Function

Public Function setLeft(value As Single) As XOptionsPanel
    Set setLeft = Me
    pLeft = value
    Call adjustPosition
End Function

Public Function setIndex(value As Long) As XOptionsPanel
    Set setIndex = Me
    pIndex = value
End Function


Public Function setUiProperty(prop As UIProp, value As Variant, _
                                    Optional skipRefreshingView As Boolean = False) As XOptionsPanel
    Set setUiProperty = Me
    
    Select Case prop.getId
        Case UI.props.width.getId:              pWidth = VBA.CSng(value)
        Case UI.props.height.getId:             pHeight = VBA.CSng(value)
        Case UI.props.minHeight.getId:          pMinHeight = VBA.CSng(value)
        Case UI.props.maxHeight.getId:          pMaxHeight = VBA.CSng(value)
        Case UI.props.top.getId:                pTop = VBA.CSng(value)
        Case UI.props.left.getId:               pLeft = VBA.CSng(value)
        Case UI.props.right.getId:              pRight = VBA.CSng(value)
        Case UI.props.bottom.getId:             pBottom = VBA.CSng(value)
        Case UI.props.margin.getId:             pMargin = VBA.CSng(value)
    End Select
    
    If Not skipRefreshingView And pIsRendered Then
        Select Case prop.getId
            Case UI.props.top.getId, UI.props.left.getId, UI.props.right.getId, UI.props.bottom.getId:
                Call adjustPosition
            Case Else
                Call adjustSize
                Call arrangeLines
        End Select
    End If
    
End Function












'[GETTERS]
Public Function getName() As String
    getName = pName
End Function

Public Function getParentFrame() As MSForms.frame
    Set getParentFrame = pParentFrame
End Function

Public Function getItemsFrame() As MSForms.frame
    Set getItemsFrame = pFrame
End Function

Public Function getItems() As Scripting.Dictionary
    Set getItems = pItems
End Function

Public Function getIndex() As Long
    getIndex = pIndex
End Function

Public Function getMargin() As Single
    getMargin = pMargin
End Function

Public Function getBottom() As Single
    getBottom = pBottom
End Function

Public Function getWidth() As Single
    getWidth = pRight - pLeft
End Function

Public Function getCheckboxWidth() As Single
    getCheckboxWidth = pCheckboxWidth
End Function

Public Function getLabelWidth() As Single
    getLabelWidth = pLabelWidth
End Function

Public Function getCheckboxValueSpace() As Single
    getCheckboxValueSpace = pCheckboxValueSpace
End Function

Public Function getLineHeight() As Single
    getLineHeight = pLineHeight
End Function

Public Function getUiProperty(prop As UIProp) As Variant
    Select Case prop.getId
        Case UI.props.width.getId:              getUiProperty = pWidth
        Case UI.props.height.getId:             getUiProperty = pHeight
        Case UI.props.minHeight.getId:          getUiProperty = pMinHeight
        Case UI.props.maxHeight.getId:          getUiProperty = pMaxHeight
        Case UI.props.top.getId:                getUiProperty = pFrame.top
        Case UI.props.left.getId:               getUiProperty = pFrame.left
        Case UI.props.right.getId:              getUiProperty = pFrame.left + pFrame.width
        Case UI.props.bottom.getId:             getUiProperty = pFrame.top + pFrame.height
    End Select
End Function





'[RENDERING]
Public Sub render()
    If Not pIsRendered Then
        pDisableEvents = True
        Call insertFrame
        Call insertLines
        pIsRendered = True
        Call adjustSize
        Call adjustPosition
        Call arrangeLines
        pDisableEvents = False
        Call updateValidation
    End If
    
    
ExitPoint:
    pDisableEvents = False
    
End Sub

Private Sub insertFrame()
    If pFrame Is Nothing Then
        Set pFrame = pParentFrame.controls.Add(CREATE_FRAME_ID)
        With pFrame
            .SpecialEffect = fmSpecialEffectFlat
            .borderStyle = fmBorderStyleNone
            .backColor = pParentFrame.backColor
            .width = 100
            .height = 100
        End With
    End If
End Sub

Private Sub insertLines()
    Dim key As Variant
    Dim value As Variant
    Dim line As xxOptionLine
    '------------------------------------------------------------------------------------------------------
    
    Set pLines = F.dictionaries.Create(False)
    
    If F.dictionaries.isNonEmptyDictionary(pItems) Then
        For Each key In pItems.keys
            Set line = UI.Create.OptionLine(Me, key, pItems.item(key))
            Call line.render
            
            '[Select if applies]
            If isItemInitiallySelected(key) Then Call line.setState(True)
            '[Lock if applies]
            If isItemFixed(key) Then Call line.setAsFixed(True)
            
            Call pLines.Add(key, line)
            
        Next key
    End If
    
End Sub


Private Function isItemInitiallySelected(key As Variant) As Boolean
    If pInitialSelection Is Nothing Then
    Else
        isItemInitiallySelected = pInitialSelection.Exists(key)
    End If
End Function

Private Function isItemFixed(key As Variant) As Boolean
    If pFixedItems Is Nothing Then
    Else
        isItemFixed = pFixedItems.Exists(key)
    End If
End Function

Public Sub insertBorder(index As XlBordersIndex, color As Long)
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If index = xlEdgeTop Then
        Call updateBorderVariable(pTopBorder, color)
    ElseIf index = xlEdgeBottom Then
        Call updateBorderVariable(pBottomBorder, color)
    End If
    
    Call updateBorderSizeAndPosition
    
End Sub


Private Sub updateBorderVariable(ByRef borderLabel As MSForms.label, color As Long)
    If borderLabel Is Nothing Then
        Set borderLabel = pFrame.controls.Add(CREATE_LABEL_ID)
        With borderLabel
            .borderStyle = fmBorderStyleNone
            .BackStyle = fmBackStyleOpaque
            .backColor = color
            .height = PIXEL_SIZE
        End With
    End If
End Sub

Private Sub updateBorderSizeAndPosition()
    If UI.Forms.isControlValid(pTopBorder) Then
        With pTopBorder
            .left = 0
            .top = PIXEL_SIZE
            .width = pFrame.InsideWidth
            .height = PIXEL_SIZE
        End With
    End If
    
    If UI.Forms.isControlValid(pBottomBorder) Then
        With pBottomBorder
            .left = 0
            .top = pFrame.InsideHeight - PIXEL_SIZE
            .width = pFrame.InsideWidth
            .height = PIXEL_SIZE
        End With
    End If
    
End Sub

Public Sub addValidationLabel(value As UIValidationLabel)
    Set pValidationLabel = value
    Call pValidationLabel.setHeight(pLineHeight)
End Sub




Private Sub adjustSize()
    Dim width As Single
    Dim height As Single
    '------------------------------------------------------------------------------------------------------
    
    If pWidth > 0 Then width = pWidth Else width = pParentFrame.InsideWidth - 2 * pMargin
    If pHeight > 0 Then height = pHeight Else height = pMinHeight
    
    If pIsRendered Then
        With pFrame
            If width >= 0 Then .width = width
            If height >= 0 Then .height = height
        End With
        
        If UI.Forms.isControlValid(pTopBorder) Then pTopBorder.width = width
        If UI.Forms.isControlValid(pBottomBorder) Then pBottomBorder.width = width
        
    End If
    
    If Not pDisableEvents Then RaiseEvent AfterResize
    
End Sub

Private Sub adjustPosition()
    If pIsRendered Then
        With pFrame
            .top = pTop
            .left = pLeft
        End With
        
        If Not pValidationLabel Is Nothing Then
            Call pValidationLabel.setTop(pTop + pMargin)
        End If
        
    End If
    
    If Not pDisableEvents Then RaiseEvent AfterMoved
    
End Sub

Private Sub arrangeLines()
    Dim key As Variant
    Dim line As xxOptionLine
    Dim top As Single:                  top = pMargin
    Dim right As Single
    '------------------------------------------------------------------------------------------------------
    
    If F.dictionaries.isNonEmptyDictionary(pLines) Then
        For Each key In pLines.keys
            Set line = pLines.item(key)
            Call line.setTop(top)
            top = top + pLineHeight
        Next key
    End If
    
    Call updatePanelHeightAndScrollBar
    
End Sub


Private Sub updatePanelHeightAndScrollBar()
    Dim lastLineBottom As Single
    '------------------------------------------------------------------------------------------------------
    
    lastLineBottom = UI.Forms.calculateRightBottomPoint(pFrame, True).y
    
    With pFrame
        '[Panel height]
        If pHeight <> 0 Then
            pFrame.height = pHeight
        ElseIf lastLineBottom < pMinHeight Then
            pFrame.height = pMinHeight
        ElseIf lastLineBottom <= pMaxHeight Then
            pFrame.height = lastLineBottom
        Else
            pFrame.height = pMaxHeight
        End If
        
        '[Vertical scroll bar]
        If lastLineBottom > pFrame.height Then
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = lastLineBottom
            .scrollTop = 0
        Else
            .ScrollBars = fmScrollBarsNone
            .ScrollHeight = pFrame.height
            .scrollTop = 0
        End If
    End With
    
End Sub

Public Function getSelectedItems() As Scripting.Dictionary
    Stop
End Function




'[VALIDATION]
Public Sub updateValidation()
    Dim key As Variant
    Dim line As xxOptionLine
    Dim previousStatus As Boolean
    Dim countSelected As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not pDisableEvents Then
        If Not pValidation Is Nothing Then
            previousStatus = pValidation.getStatus
        End If
        
        For Each key In pLines.keys
            Set line = pLines.item(key)
            If line.isSelected Then countSelected = countSelected + 1
        Next key
        
        If pMinItemsSelected > 0 And pMinItemsSelected > countSelected Then
            Set pValidation = F.Create.ValidationObject(False, _
                                    F.Strings.format(MsgService.GetText(WARNING_TOO_LESS_ITEMS_SELECTED), pMinItemsSelected))
        ElseIf pMaxItemsSelected > 0 And pMaxItemsSelected < countSelected Then
            Set pValidation = F.Create.ValidationObject(False, _
                                    F.Strings.format(MsgService.GetText(WARNING_TOO_MANY_ITEMS_SELECTED), pMaxItemsSelected))
        Else
            Set pValidation = F.Create.ValidationObject(True)
        End If
        
        
        If previousStatus <> pValidation.getStatus Then
            RaiseEvent AfterValidationChange(pValidation)
            If Not pListener Is Nothing Then Call pListener.catchEvent(Me, Events.ValidationChange.getName)
        End If
        
        Call updateValidationLabelView
        
        RaiseEvent AfterChange
        If Not pListener Is Nothing Then Call pListener.catchEvent(Me, Events.change.getName)
        
    End If
    
End Sub



Public Sub updateValidationLabelView()
    If pValidation Is Nothing Then
        'Cannot read status, because there is no validation object.
    ElseIf pValidationLabel Is Nothing Then
        'Cannot update label, because it doesn't exist.
        Stop
    ElseIf pValidation.getStatus Then
        Call pValidationLabel.setVisible(False)
    Else
        With pValidationLabel
            Call .setVisible(True)
            Call .setCaption(pValidation.getMessage)
        End With
    End If
End Sub

Public Function getValidation() As ValidationObject
    Set getValidation = pValidation
End Function




Public Function changeSelectionForAllItems(state As Boolean) As ValidationObject
    Dim key As Variant
    Dim line As xxOptionLine
    '------------------------------------------------------------------------------------------------------
    
    pDisableEvents = True
    For Each key In pLines.keys
        Set line = pLines.item(key)
        If Not line.isFixed Then
            Call line.setState(state)
        End If
    Next key
    pDisableEvents = False
    
    Call updateValidation
    
ExitPoint:
    pDisableEvents = False
    
End Function



Public Function getSelectedItemsDictionary() As Scripting.Dictionary
    Dim key As Variant
    Dim line As xxOptionLine
    Dim name As String
    '------------------------------------------------------------------------------------------------------
    
    Set getSelectedItemsDictionary = F.dictionaries.Create(False)
    
    If F.dictionaries.isNonEmptyDictionary(pLines) Then
        For Each key In pLines.keys
            Set line = pLines.item(key)
            If line.isSelected Then
                Call F.dictionaries.addItem(getSelectedItemsDictionary, key, line.getItem)
            End If
        Next key
    End If
    
End Function


Public Sub setFocus()
    Dim line As xxOptionLine
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    Set line = pLines.items(0)
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If Not line Is Nothing Then
        'Call line.setFocus
        Stop
    End If
    
ErrHandler:
End Sub







'[Implementation of IControl interface]
Private Sub IControl_setVisibility(value As Boolean)
    Stop
End Sub

Private Sub IControl_setEnability(value As Boolean)
    Stop
End Sub

Private Sub IControl_setWidth(value As Single)
    Stop
End Sub

Private Function IControl_getHeight() As Single
    Stop
End Function

Private Function IControl_getWidth() As Single
    Stop
End Function

Private Function IControl_getLeft() As Single
    Stop
End Function

Private Function IControl_getTop() As Single
    Stop
End Function

Private Function IControl_getRectangle() As MielkCore.RECT
    Stop
End Function

Private Sub IControl_resize(Optional width As Variant, Optional height As Variant)
    Stop
End Sub

Private Sub IControl_alignVertically(baseCtrl As MSForms.IControl)
    Stop
End Sub

Private Function IControl_move(Optional x As Variant, Optional y As Variant) As IControl
    Stop
End Function

Private Function IControl_offset(Optional x As Single, Optional y As Single) As IControl
    Stop
End Function

Private Sub IControl_triggerEvent(eventId As Long)
    Stop
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

Private Function IPanel_setParentFrame(value As MSForms.frame) As IPanel
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
    Call addValidationLabel(value)
End Sub

Private Function IPanel_setListener(value As IListener) As IPanel
    Set IPanel_setListener = setListener(value)
End Function

Private Function IPanel_getValidation() As MielkCore.ValidationObject
    Set IPanel_getValidation = getValidation
End Function

Private Function IPanel_getValue() As Variant
    Set IPanel_getValue = getSelectedItemsDictionary
End Function

Private Function IPanel_setWindow(value As WTemplate) As IPanel
    
End Function

Private Function IPanel_isVisible() As Boolean
    If UI.Forms.isControlValid(pFrame) Then
        IPanel_isVisible = pFrame.visible
    End If
End Function

