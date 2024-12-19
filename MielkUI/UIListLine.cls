VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIListLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IControl

Private Enum ControlTypeEnum
    ControlType_Background = 1
    ControlType_Caption = 2
    ControlType_RemoveIcon = 3
    ControlType_EditIcon = 4
    ControlType_PreviewIcon = 5
    ControlType_ActivenessBorder = 6
End Enum

Private Const CLASS_NAME As String = "UIListLine"
'[UI constants] -------------------------------------------------------------------------------------------
Private Const DEFAULT_ICON_SIZE As Single = 10.5
Private Const DEFAULT_LINE_HEIGHT As Single = 15
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_FONT_COLOR As Long = VBA.vbBlack
Private Const DEFAULT_FONT_FAMILY As String = APP_FONT_FAMILY
Private Const DEFAULT_PADDING As Single = 6
Private Const DEFAULT_BACK_COLOR As Long = VBA.vbWhite
Private Const DEFAULT_CURRENT_ITEM_BACK_COLOR As Long = 16053492
Private Const DEFAULT_SELECTED_ITEM_BACK_COLOR As Long = 13273922
Private Const DEFAULT_CURRENT_SELECTED_BACK_COLOR As Long = 14597527
Private Const DEFAULT_CURRENT_ITEM_BORDER_COLOR As Long = 9868950
Private Const DEFAULT_BORDER_COLOR As Long = VBA.vbWhite
Private Const DEFAULT_SELECTION_BORDER_COLOR As Long = &H80000006
Private Const DEFAULT_VISIBILITY As Boolean = True
Private Const DEFAULT_SELECTED_ITEM_FONT_COLOR As Long = vbWhite
Private Const DEFAULT_CURRENT_ITEM_FONT_COLOR As Long = vbBlack
Private Const DEFAULT_CURRENT_SELECTED_FONT_COLOR As Long = vbWhite
'[UI components] ------------------------------------------------------------------------------------------
Private pList As UIList
Private WithEvents pBackgroundLabel As MSForms.label
Attribute pBackgroundLabel.VB_VarHelpID = -1
Private WithEvents pCaptionLabel As MSForms.label
Attribute pCaptionLabel.VB_VarHelpID = -1
Private WithEvents pActivenessBorder As MSForms.label
Attribute pActivenessBorder.VB_VarHelpID = -1
Private WithEvents pRemoveIcon As MSForms.label
Attribute pRemoveIcon.VB_VarHelpID = -1
Private WithEvents pEditIcon As MSForms.label
Attribute pEditIcon.VB_VarHelpID = -1
Private WithEvents pPreviewIcon As MSForms.label
Attribute pPreviewIcon.VB_VarHelpID = -1
'[UI properties] ------------------------------------------------------------------------------------------
Private pTop As Single
Private pHeight As Single
Private pPadding As Single
Private pFontSize As Single
Private pFontFamily As String
Private pFontColor As Long
Private pSelectedFontColor As Long
Private pCurrentFontColor As Long
Private pCurrentSelectedFontColor As Long
Private pBackColor As Long
Private pSelectedBackColor As Long
Private pCurrentBackColor As Long
Private pCurrentSelectedBackColor As Long
Private pBorderColor As Long
Private pSelectedBorderColor As Long
Private pActiveBorderColor As Long
Private pIconSize As Single
'[Special formats]
Private pSpecialFormat As Scripting.Dictionary
Private pSpecialForeColor As Long                       'Cached for better performance
Private pSpecialFontBold As Boolean
'[State] --------------------------------------------------------------------------------------------------
Private pUuid As String
Private pIndex As Long
Private pItem As Variant
Private pKey As Variant
Private pCaption As String
Private pIsRendered As Boolean
Private pVisible As Boolean
Private pSelected As Boolean
Private pActive As Boolean
Private pPreviewMode As Boolean
'----------------------------------------------------------------------------------------------------------





'[INITIALIZE]
Private Sub Class_Initialize()
    Call initializeDefaultValues
End Sub

Private Sub initializeDefaultValues()
    pVisible = DEFAULT_VISIBILITY
    pHeight = UI.Normalize.size(DEFAULT_LINE_HEIGHT)
    pFontSize = DEFAULT_FONT_SIZE
    pFontFamily = DEFAULT_FONT_FAMILY
    pFontColor = DEFAULT_FONT_COLOR
    pSelectedFontColor = DEFAULT_SELECTED_ITEM_FONT_COLOR
    pCurrentFontColor = DEFAULT_CURRENT_ITEM_FONT_COLOR
    pCurrentSelectedFontColor = DEFAULT_CURRENT_SELECTED_FONT_COLOR
    pPadding = DEFAULT_PADDING
    pBackColor = DEFAULT_BACK_COLOR
    pSelectedBackColor = DEFAULT_SELECTED_ITEM_BACK_COLOR
    pCurrentBackColor = DEFAULT_CURRENT_ITEM_BACK_COLOR
    pCurrentSelectedBackColor = DEFAULT_CURRENT_SELECTED_BACK_COLOR
    pBorderColor = DEFAULT_BORDER_COLOR
    pActiveBorderColor = DEFAULT_CURRENT_ITEM_BORDER_COLOR
    pIconSize = DEFAULT_ICON_SIZE
    pSpecialForeColor = -1
    pUuid = F.Crypto.createUUID
End Sub






'[SETTERS]
Public Function setParent(value As UIList) As UIListLine
    Set setParent = Me
    Set pList = value
End Function


'[Values]
Public Function setIndex(value As Long) As UIListLine
    Set setIndex = Me
    pIndex = value
End Function

Public Function setItem(value As Variant) As UIListLine
    Set setItem = Me
    Call F.Variables.assign(pItem, value)
End Function

Friend Function setKey(value As Variant) As UIListLine
    Set setKey = Me
    Call F.Variables.assign(pKey, value)
End Function

Public Function setCaption(value As String) As UIListLine
    Set setCaption = Me
    pCaption = value
    If pIsRendered Then Call refreshCaptionLabel
End Function

Public Function setSpecialFormat(value As Scripting.Dictionary) As UIListLine
    Dim changed As Boolean
    '------------------------------------------------------------------------------------------------------
    Set setSpecialFormat = Me
    If Not pSpecialFormat Is Nothing Then
        changed = Not value Is pSpecialFormat
    Else
        changed = True
    End If
    
    If changed Then
        Set pSpecialFormat = value
        Call applySpecialFormat
    End If
End Function




'[UI Properties]
Public Function setActiveState(value As Boolean) As UIListLine
    Set setActiveState = Me
    pActive = value
    If pIsRendered Then Call refreshViewAfterStateChange
End Function

Public Function setSelection(value As Boolean) As UIListLine
    Set setSelection = Me
    pSelected = value
    If pIsRendered Then Call refreshViewAfterStateChange
End Function

Public Function setVisibility(value As Boolean) As UIListLine
    Set setVisibility = Me
    pVisible = value
    If pIsRendered Then Call updateVisibility
End Function
    
Public Function setFontSize(value As Single) As UIListLine
    Set setFontSize = Me
    pFontSize = value
    If pIsRendered Then Call refreshCaptionLabel
End Function

Public Function setFontFamily(value As String) As UIListLine
    Set setFontFamily = Me
    pFontFamily = value
    If pIsRendered Then Call refreshCaptionLabel
End Function

Public Function setFontColor(value As Long) As UIListLine
    Set setFontColor = Me
    pFontColor = value
    If pIsRendered Then Call refreshCaptionLabel
End Function



'[Position]
Public Function setTop(value As Single) As UIListLine
    Set setTop = Me
    pTop = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBottom(value As Single) As UIListLine
    Set setBottom = Me
    pTop = value - pHeight
    If pIsRendered Then Call arrangeControls
End Function


'[Size]
Public Function setHeight(value As Single) As UIListLine
    Set setHeight = Me
    pHeight = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setPadding(value As Single) As UIListLine
    Set setPadding = Me
    pPadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBackColor(value As Long) As UIListLine
    Set setBackColor = Me
    pBackColor = value
    If pIsRendered Then Call refreshBackgroundLabel
End Function

Public Function setSelectionBackColor(value As Long) As UIListLine
    Set setSelectionBackColor = Me
    pSelectedBackColor = value
    If pIsRendered Then Call refreshBackgroundLabel
End Function

Public Function setBorderColor(value As Long) As UIListLine
    Set setBorderColor = Me
    pBorderColor = value
    If pIsRendered Then Call refreshBackgroundLabel
End Function

Public Function setSelectionBorderColor(value As Long) As UIListLine
    Set setSelectionBorderColor = Me
    pSelectedBorderColor = value
    If pIsRendered Then Call refreshBackgroundLabel
End Function

Friend Sub setPreviewMode(value As Boolean)
    pPreviewMode = value
    pCaptionLabel.font.Bold = VBA.IIf(pPreviewMode Or pSpecialFontBold, 1, False)
    Call refreshPreviewIcon
End Sub





'[GETTERS]
Public Function getUuid() As String
    getUuid = pUuid
End Function

Public Function getItem() As Variant
    Call F.Variables.assign(getItem, pItem)
End Function

Friend Function getKey() As Variant
    Call F.Variables.assign(getKey, pKey)
End Function

Public Function isValid() As Boolean
    isValid = UI.controls.isValid(pBackgroundLabel)
End Function

Public Function getLeft() As Single
    getLeft = pBackgroundLabel.left
End Function

Public Function getRight() As Single
    getRight = UI.controls.getRight(pBackgroundLabel)
End Function

Public Function getTop() As Single
    getTop = pBackgroundLabel.top
End Function

Public Function getBottom() As Single
    getBottom = UI.controls.getBottom(pBackgroundLabel)
End Function

Public Function getRectangle() As RECT
    With getRectangle
        .top = getTop
        .bottom = getBottom
        .left = getLeft
        .right = getRight
    End With
End Function

Public Function getWidth() As Single
    getWidth = pBackgroundLabel.width
End Function

Public Function getHeight() As Single
    getHeight = pBackgroundLabel.height
End Function

Public Function getCanvas() As MSForms.UserForm
    Set getCanvas = pList.getCanvas
End Function

Public Function getCaption() As String
    getCaption = pCaption
End Function

Public Function getIndex() As Long
    getIndex = pIndex
End Function

Public Function isVisible() As Boolean
    isVisible = pVisible
End Function

Public Function isSelected() As Boolean
    isSelected = pSelected
End Function

Public Function isActive() As Boolean
    isActive = pActive
End Function




'[API]
Public Sub render()
    Call insertBackgroundLabel
    Call insertActivenessBorder
    Call insertCaptionLabel
    Call insertRemoveIcon
    Call insertEditIcon
    Call insertPreviewIcon
    pIsRendered = True
End Sub

Public Function Move(Optional left As Variant, Optional top As Variant) As UIListLine
    Set Move = Me
    'Left cannot be modified - line always fills the whole frame, therefore Left = 0.
    If F.Variables.isOfNumericType(top) Then pTop = top
    Call arrangeControls
End Function

Public Function offset(Optional left As Variant, Optional top As Variant) As UIListLine
    Set offset = Me
    'Left cannot be modified - line always fills the whole frame, therefore Left = 0.
    If F.Variables.isOfNumericType(top) Then pTop = pTop + top
    Call arrangeControls
End Function

Public Function resize(Optional width As Variant, Optional height As Variant) As UIListLine
    Set resize = Me
    'Width cannot be modified - line always fills the whole frame.
    If F.Math.isNotNegative(height) Then pHeight = height
    Call arrangeControls
End Function

Public Sub destroy()
    Call removeControls
End Sub

Friend Sub clear()
    Set pItem = Nothing
    Call setSpecialFormat(Nothing)
    pCaption = vbNullString
    Call resetCaptionLabel
    Call setVisibility(False)
End Sub





'[RENDERING]
Private Function getLabelName(controlType As ControlTypeEnum) As String
    Dim tag As String
    '----------------------------------------------------------------------------------------------------------
    tag = CLASS_NAME & "|" & pUuid & "|" & controlType
    getLabelName = "ctrl_" & VBA.left$(F.Crypto.SHA1HASH(tag), 25)
End Function

Private Sub insertBackgroundLabel()
    Set pBackgroundLabel = getCanvas.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_Background), False)
    Call refreshBackgroundLabel
End Sub

Private Sub insertActivenessBorder()
    Set pActivenessBorder = getCanvas.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_ActivenessBorder), False)
    Call refreshActivenessBorder
End Sub

Private Sub insertCaptionLabel()
    Set pCaptionLabel = getCanvas.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_Caption), False)
    Call refreshCaptionLabel
    Call resizeCaptionLabel
End Sub

Private Sub insertRemoveIcon()
'    Stop 'xxx
    If pList.areItemsRemovable Then
        Set pRemoveIcon = getCanvas.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_RemoveIcon), False)
        Call UI.controls.resize(pRemoveIcon, pIconSize, pIconSize)
        Call refreshRemoveIcon
    End If
End Sub

Private Sub insertEditIcon()
'    Stop 'xxx
    If pList.areItemsEditable Then
        Set pEditIcon = getCanvas.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_EditIcon), False)
        Call UI.controls.resize(pEditIcon, pIconSize, pIconSize)
        Call refreshEditIcon
    End If
End Sub

Private Sub insertPreviewIcon()
    If pList.areItemsPreviewable Then
        Set pPreviewIcon = getCanvas.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_PreviewIcon), False)
        Call UI.controls.resize(pPreviewIcon, pIconSize, pIconSize)
        Call refreshPreviewIcon
    End If
End Sub

Private Sub refreshBackgroundLabel()
    With pBackgroundLabel
        .BackStyle = MSForms.fmBackStyleOpaque
        .backColor = getProperBackColor
        .borderStyle = MSForms.fmBorderStyleNone
    End With
End Sub

Private Sub refreshActivenessBorder()
    With pActivenessBorder
        .BackStyle = MSForms.fmBackStyleTransparent
        .borderColor = pActiveBorderColor
        .borderStyle = MSForms.fmBorderStyleSingle
    End With
End Sub

Private Sub refreshCaptionLabel()
    With pCaptionLabel
        .BackStyle = 0 'MSForms.fmBackStyleTransparent
        .borderStyle = MSForms.fmBorderStyleNone
        .font.size = pFontSize
        .font.name = pFontFamily
        .ForeColor = getProperFontColor
        .caption = pCaption
        .width = getCanvas.InsideWidth
        .AutoSize = True
        .AutoSize = False
        .width = getCanvas.InsideWidth
        If Not pSpecialFormat Is Nothing Then Call applySpecialFormat
        Call arrangeControls
    End With
End Sub

Private Sub resetCaptionLabel()
    With pCaptionLabel
        .BackStyle = 0 'MSForms.fmBackStyleTransparent
        .borderStyle = MSForms.fmBorderStyleNone
        .font.Italic = False
        .font.Bold = False
        .font.size = pFontSize
        .font.name = pFontFamily
        .ForeColor = getProperFontColor
        .caption = pCaption
        .width = getCanvas.InsideWidth
        .AutoSize = True
        .AutoSize = False
        .width = getCanvas.InsideWidth
        If Not pSpecialFormat Is Nothing Then Call applySpecialFormat
        Call arrangeControls
    End With
End Sub

Private Sub refreshRemoveIcon()
    If UI.controls.isValid(pRemoveIcon) Then
        With pRemoveIcon
            .BackStyle = fmBackStyleTransparent
            .borderStyle = fmBorderStyleNone
            .caption = vbNullString
            .MousePointer = fmMousePointerCustom
            .MouseIcon = Icons.getMousePointer
            Set .picture = getProperRemoveIcon
        End With
    End If
End Sub

Private Sub refreshEditIcon()
    If UI.controls.isValid(pEditIcon) Then
        With pEditIcon
            .BackStyle = fmBackStyleTransparent
            .borderStyle = fmBorderStyleNone
            .caption = vbNullString
            .MousePointer = fmMousePointerCustom
            .MouseIcon = Icons.getMousePointer
            Set .picture = getProperEditIcon
        End With
    End If
End Sub

Private Sub refreshPreviewIcon()
    If UI.controls.isValid(pPreviewIcon) Then
        With pPreviewIcon
            .BackStyle = fmBackStyleTransparent
            .borderStyle = fmBorderStyleNone
            .caption = vbNullString
            .MousePointer = fmMousePointerCustom
            .MouseIcon = Icons.getMousePointer
            Set .picture = getProperPreviewIcon
        End With
    End If
End Sub

Private Sub refreshViewAfterStateChange()
    pBackgroundLabel.backColor = getProperBackColor
    pCaptionLabel.ForeColor = getProperFontColor
    If UI.controls.isValid(pRemoveIcon) Then Set pRemoveIcon.picture = getProperRemoveIcon
    If UI.controls.isValid(pEditIcon) Then Set pEditIcon.picture = getProperEditIcon
    If UI.controls.isValid(pPreviewIcon) Then Set pPreviewIcon.picture = getProperPreviewIcon
    Call updateActivenessBorder
End Sub

Private Function getProperFontColor() As Long
    If pSpecialForeColor >= 0 Then
        getProperFontColor = VBA.IIf(pSelected, pCurrentSelectedFontColor, pSpecialForeColor)
    ElseIf pSelected Then
        getProperFontColor = pCurrentSelectedFontColor
    Else
        getProperFontColor = pFontColor
    End If
End Function

Private Function getProperBackColor() As Long
    If pSelected Then
        getProperBackColor = pCurrentSelectedBackColor
    Else
        getProperBackColor = pBackColor
    End If
End Function

Private Function getProperRemoveIcon() As Variant
    If pSelected Then
        Set getProperRemoveIcon = Icons.getCancelListItem_LightBlue
    Else
        Set getProperRemoveIcon = Icons.getCancelListItem_White
    End If
End Function

Private Function getProperEditIcon() As Variant
    If pSelected Then
        Set getProperEditIcon = Icons.getEditListItem_LightBlue
    Else
        Set getProperEditIcon = Icons.getEditListItem_White
    End If
End Function

Private Function getProperPreviewIcon() As Variant
    If pPreviewMode Then
        If pSelected Then
            Set getProperPreviewIcon = Icons.getEye_Active_LightBlue
        Else
            Set getProperPreviewIcon = Icons.getEye_Active_White
        End If
    Else
        If pSelected Then
            Set getProperPreviewIcon = Icons.getEye_Inactive_LightBlue
        Else
            Set getProperPreviewIcon = Icons.getEye_Inactive_White
        End If
    End If

End Function

Private Sub updateActivenessBorder()
    pActivenessBorder.visible = pActive
End Sub

Private Sub removeControls()
    On Error Resume Next
    With getCanvas.controls
        Call .remove(getLabelName(ControlType_Background))
        Call .remove(getLabelName(ControlType_Caption))
        Call .remove(getLabelName(ControlType_EditIcon))
        Call .remove(getLabelName(ControlType_RemoveIcon))
        Call .remove(getLabelName(ControlType_ActivenessBorder))
        Call .remove(getLabelName(ControlType_PreviewIcon))
    End With
End Sub

Private Sub arrangeControls()
    Dim left As Single:                         left = pPadding / 2
    Dim topCaption As Single
    Dim topButton As Single
    Dim icon As Variant
    '------------------------------------------------------------------------------------------------------
    
    Call checkControlsReferences
    Call moveControlsOutOfView
    Call updateVisibility
    Call resizeControls
    Call UI.controls.locate(pBackgroundLabel, -PIXEL_SIZE, pTop)
    With pBackgroundLabel
        Call UI.controls.locate(pActivenessBorder, .left + 3 * PIXEL_SIZE, .top + 2 * PIXEL_SIZE)
        Call UI.controls.resize(pActivenessBorder, .width - 6 * PIXEL_SIZE, .height - 4 * PIXEL_SIZE)
    End With
    
    topCaption = pTop + (pBackgroundLabel.height - pCaptionLabel.height) / 2
    topButton = pTop + (pBackgroundLabel.height - pIconSize) / 2
    
    If pList.actionButtonsToLeft Then
        For Each icon In VBA.Array(pRemoveIcon, pEditIcon, pPreviewIcon)
            Call locateIcon(icon, left, topButton)
        Next icon
        Call UI.controls.locate(pCaptionLabel, left + pPadding / 2, topCaption)
    Else
        Call UI.controls.locate(pCaptionLabel, left, topCaption)
        For Each icon In VBA.Array(pRemoveIcon, pEditIcon, pPreviewIcon)
            Call locateIcon(icon, left, topButton)
        Next icon
    End If
    
    DoEvents
End Sub

Private Sub locateIcon(icon As Variant, ByRef left As Single, ByVal top As Single)
    Dim label As MSForms.label
    '------------------------------------------------------------------------------------------------------
    If UI.controls.isValid(icon) Then
        Set label = icon
        Call UI.controls.locate(label, left, top)
        left = left + pIconSize + pPadding / 2
    End If
End Sub

Private Sub checkControlsReferences()
    On Error Resume Next
    With pList.getCanvas
        Set pBackgroundLabel = .controls(getLabelName(ControlType_Background))
        Set pCaptionLabel = .controls(getLabelName(ControlType_Caption))
        Set pRemoveIcon = .controls(getLabelName(ControlType_RemoveIcon))
        Set pEditIcon = .controls(getLabelName(ControlType_EditIcon))
        Set pPreviewIcon = .controls(getLabelName(ControlType_PreviewIcon))
    End With
End Sub

Private Sub moveControlsOutOfView()
    On Error Resume Next
    Call UI.controls.locate(pBackgroundLabel, -1, -pHeight - 1)
    Call UI.controls.locate(pCaptionLabel, 0, -pHeight - 1)
    Call UI.controls.locate(pRemoveIcon, 0, -pHeight - 1)
    Call UI.controls.locate(pEditIcon, 0, -pHeight - 1)
    Call UI.controls.locate(pPreviewIcon, 0, -pHeight - 1)
End Sub

Private Function getRemoveIconLeftCoordinate() As Single
    If UI.controls.isValid(pBackgroundLabel) Then
        getRemoveIconLeftCoordinate = UI.controls.getRight(pBackgroundLabel) - pIconSize - pPadding / 2 - VBA.IIf(pList.hasVerticalScrollBar, SCROLL_BAR_WIDTH, 0)
    End If
End Function

Private Function getEditIconLeftCoordinate() As Single
    If UI.controls.isValid(pRemoveIcon) Then
        getEditIconLeftCoordinate = pRemoveIcon.left - pIconSize - pPadding / 2
    Else
        getEditIconLeftCoordinate = UI.controls.getRight(pBackgroundLabel) - SCROLL_BAR_WIDTH - pIconSize - pPadding / 2
    End If
End Function

Private Sub resizeControls()
    Call UI.controls.resize(pBackgroundLabel, getCanvas.InsideWidth + 2, pHeight)
    Call resizeCaptionLabel
    Call UI.controls.resize(pRemoveIcon, pIconSize, pIconSize)
    Call UI.controls.resize(pEditIcon, pIconSize, pIconSize)
    Call UI.controls.resize(pPreviewIcon, pIconSize, pIconSize)
End Sub

Private Sub resizeCaptionLabel()
    With pCaptionLabel
        .width = pBackgroundLabel.width
        .AutoSize = True
        .AutoSize = False
        .width = pBackgroundLabel.width
    End With
End Sub

Private Sub updateVisibility()
    On Error Resume Next
    pBackgroundLabel.visible = pVisible
    pCaptionLabel.visible = pVisible
    pActivenessBorder.visible = pVisible And pActive
    If UI.controls.isValid(pRemoveIcon) Then pRemoveIcon.visible = pVisible
    If UI.controls.isValid(pEditIcon) Then pEditIcon.visible = pVisible
    If UI.controls.isValid(pPreviewIcon) Then pPreviewIcon.visible = pVisible
End Sub

Friend Sub applySpecialFormat()
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    If Not pSpecialFormat Is Nothing Then

        '[Assigned cacheable values to variables]
        If pSpecialFormat.Exists(UI.props.fontColor.getName) Then
            pSpecialForeColor = pSpecialFormat.item(UI.props.fontColor.getName)
        End If

        If pSpecialFormat.Exists(UI.props.fontBold.getName) Then
            pSpecialFontBold = pSpecialFormat.item(UI.props.fontBold.getName)
        End If

        '[Apply properties]
        For Each key In pSpecialFormat.keys
            Call applyProperty(VBA.CStr(key), pSpecialFormat.item(key))
        Next key
    
    Else
        pSpecialForeColor = -1
        Call resetCaptionLabel
    End If
    
End Sub

Private Sub applyProperty(propName As String, value As Variant)
    Dim fontBold As Long
    '------------------------------------------------------------------------------------------------------
    
    If UI.controls.isValid(pCaptionLabel) Then
        With pCaptionLabel
            Select Case VBA.LCase$(propName)
                Case UI.props.fontColor.getLowerCaseName
                        .ForeColor = value
                Case UI.props.FontItalic.getLowerCaseName
                        fontBold = .font.Bold
                        .font.Italic = VBA.IIf(value, 1, False)
                        .font.Bold = VBA.IIf(fontBold, 1, False)
                Case UI.props.fontBold.getLowerCaseName
                        .font.Bold = VBA.IIf(value, 1, False)
            End Select
        End With
    End If
End Sub






'[Control events]
Private Sub pBackgroundLabel_Click()
    Call pList.triggerLineClick(Me)
End Sub

Private Sub pBackgroundLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call pList.triggerLineDoubleClick(Me)
End Sub

Private Sub pCaptionLabel_Click()
    Call pList.triggerLineClick(Me)
End Sub

Private Sub pCaptionLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call pList.triggerLineDoubleClick(Me)
    Call pList.changeSelection(Me)
End Sub

Private Sub pActivenessBorder_Click()
    Call pList.triggerLineClick(Me)
End Sub

Private Sub pActivenessBorder_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call pList.triggerLineDoubleClick(Me)
End Sub

Private Sub pEditIcon_Click()
    DoEvents
    Call pList.editItem(Me)
End Sub

Private Sub pPreviewIcon_Click()
    Call pList.setPreviewLine(Me)
End Sub

Private Sub pRemoveIcon_Click()
    DoEvents
    Call pList.removeItem(Me)
End Sub




'[Implementation of UIControl]
Private Sub IControl_setVisibility(value As Boolean)
    Call setVisibility(value)
End Sub

Private Sub IControl_setEnability(value As Boolean)
    'N/A
End Sub

Private Sub IControl_setWidth(value As Single)
    'N/A
End Sub

Private Function IControl_getHeight() As Single
    IControl_getHeight = getHeight
End Function

Private Function IControl_getWidth() As Single
    IControl_getWidth = getWidth
End Function

Private Function IControl_getLeft() As Single
    IControl_getLeft = getLeft
End Function

Private Function IControl_getTop() As Single
    IControl_getTop = getTop
End Function

Private Function IControl_getRectangle() As RECT
    IControl_getRectangle = getRectangle
End Function

Private Sub IControl_resize(Optional width As Variant, Optional height As Variant)
    Call resize(width, height)
End Sub

Private Sub IControl_alignVertically(baseCtrl As MSForms.IControl)
    'N/A
End Sub

Private Function IControl_move(Optional x As Variant, Optional y As Variant) As IControl
    Set IControl_move = Move(x, y)
End Function

Private Function IControl_offset(Optional x As Single, Optional y As Single) As IControl
    Set IControl_offset = offset(x, y)
End Function

Private Sub IControl_triggerEvent(eventId As Long)
    'N/A
End Sub
