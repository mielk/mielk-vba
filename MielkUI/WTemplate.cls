VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IListener
Implements IGridParentWindow

Private Type MoveCache
    isActive As Boolean
    x As Single
    Y As Single
End Type

Private Type FormPosition
    x As Single
    Y As Single
End Type


Private Const CLASS_NAME As String = "WTemplate"
'----------------------------------------------------------------------------------------------------------
Private Const CANVAS_FRAME_NAME As String = "frmCanvas"
Private Const BODY_FRAME_NAME As String = "frmBody"
Private Const CONTENT_FRAME_NAME As String = "frmContent"
Private Const VALIDATION_FRAME_NAME As String = "frmValidation"
'[Config tags] --------------------------------------------------------------------------------------------
Private Const TITLE_PANEL_TAG As String = "titlePanel"
'----------------------------------------------------------------------------------------------------------
Private Const WINDOW_PADDING As Single = 6
Private Const DEFAULT_WIDTH As Single = 400
Private Const DEFAULT_MIN_WIDTH As Single = 200
Private Const DEFAULT_MAX_WIDTH As Single = 2560
Private Const DEFAULT_MIN_HEIGHT As Single = 60
Private Const DEFAULT_MAX_HEIGHT As Single = 600
Private Const DEFAULT_CONTENT_HEIGHT As Single = 24
Private Const DEFAULT_VALIDATION_PANEL_WIDTH As Single = 150
Private Const DEFAULT_CONTENT_FRAME_HORIZONTAL_MARGIN As Single = 6
Private Const DEFAULT_CONTENT_FRAME_VERTICAL_MARGIN As Single = 6
Private Const DEFAULT_HAS_TITLE_BAR As Boolean = True
Private Const DEFAULT_HAS_BUTTONS_PANEL As Boolean = True
Private Const DEFAULT_DISPLAY_WARNINGS_FOR_UNSAVED_DATA As Boolean = True
'[Captions]
Private Const DEFAULT_OK_CAPTION As String = "ok"
Private Const DEFAULT_CANCEL_CAPTION As String = "cancel"
'[Validation]
Private Const DEFAULT_ERROR_BACK_COLOR As Long = vbRed
Private Const DEFAULT_ERROR_FONT_COLOR As Long = vbWhite
Private Const DEFAULT_ERROR_FONT_BOLD As Boolean = True
Private Const DEFAULT_ERROR_PADDING As Single = 3
'[Extra frame default props]
Private Const EXTRA_FRAME_DEFAULT_BACK_COLOR As Long = vbWhite
Private Const EXTRA_FRAME_DEFAULT_BORDER_STYLE As Long = fmBorderStyleNone
Private Const EXTRA_FRAME_DEFAULT_SCROLL_BARS As Long = fmScrollBarsVertical
'[State] --------------------------------------------------------------------------------------------------
Private pId As String
Private pUuid As String
Private pIsRendered As Boolean
Private pIsCancelled As Boolean
Private pValidation As ValidationObject
Private pDisplayWarningsForUnsavedData As Boolean
Private pHasUnsavedData As Boolean
Private pHelpFileTag As String
Private pBackgroundActive As Boolean
'[UI components] ------------------------------------------------------------------------------------------
Private WithEvents pForm As ufTemplate
Attribute pForm.VB_VarHelpID = -1
Private pBackgroundWindow As WBackground
Private pCanvas As MSForms.frame
Private pBody As MSForms.frame
Private WithEvents pTitlePanel As UITitlePanel
Attribute pTitlePanel.VB_VarHelpID = -1
Private pContentFrame As MSForms.frame
Private pValidationFrame As MSForms.frame
Private WithEvents pButtonsPanel As UIButtonsPanel
Attribute pButtonsPanel.VB_VarHelpID = -1
Private pExtraFrames As Scripting.Dictionary
'[UI properties] ------------------------------------------------------------------------------------------
Private pProperties As Scripting.Dictionary
Private pHasBorder As Boolean
Private pBorderColor As Long
Private pCaption As String
Private pMinWidth As Single
Private pMaxWidth As Single
Private pMinHeight As Single
Private pMaxHeight As Single
Private pHasTitleBar As Boolean
Private pHasButtonsPanel As Boolean
Private pHasValidationPanel As Boolean
Private pValidationPanelWidth As Single
Private pContentFrameHorizontalMargin As Single
Private pContentFrameVerticalMargin As Single
Private pContentFrameTopMargin As Single
'[Buttons]
Private pOkCaption As String
Private pCancelCaption As String
Private pHasCancelButton As Boolean
Private pHasOkButton As Boolean
'[Validation]
Private pErrorBackColor As Long
Private pErrorFontColor As Long
Private pErrorFontBold As Boolean
Private pErrorPadding As Single
Private pValidationLabels As Scripting.Dictionary
'[Position] -----------------------------------------------------------------------------------------------
Private pMoveCache As MoveCache
Private pPosition As FormPosition
'[Events] -------------------------------------------------------------------------------------------------
Event AfterDisplayed()
Event Cancel()
Event Confirm()
Event CustomAction(actionTag As String, params As Scripting.Dictionary)
Event HelpFileButtonClicked()
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    Call setDefaultValues
    Call initializeContainers
End Sub


Private Sub setDefaultValues()
    Set pForm = New ufTemplate
    pUuid = F.Crypto.createUUID
    pHasBorder = True
    pBorderColor = MIELK_COLOR
    pMinWidth = DEFAULT_MIN_WIDTH
    pMaxWidth = DEFAULT_MAX_WIDTH
    pMinHeight = DEFAULT_MIN_HEIGHT
    pMaxHeight = DEFAULT_MAX_HEIGHT
    pContentFrameHorizontalMargin = DEFAULT_CONTENT_FRAME_HORIZONTAL_MARGIN
    pContentFrameVerticalMargin = DEFAULT_CONTENT_FRAME_VERTICAL_MARGIN
    pHasTitleBar = DEFAULT_HAS_TITLE_BAR
    pHasButtonsPanel = DEFAULT_HAS_BUTTONS_PANEL
    pDisplayWarningsForUnsavedData = DEFAULT_DISPLAY_WARNINGS_FOR_UNSAVED_DATA
    
    '[Captions]
    pOkCaption = DEFAULT_OK_CAPTION
    pCancelCaption = DEFAULT_CANCEL_CAPTION
    pHasOkButton = True
    pHasCancelButton = True
    
    With pForm
        .caption = pUuid
        .backColor = TRANSPARENCY_LAYER_COLOR
    End With
    
    '[Temp]
    pHasValidationPanel = True
    pValidationPanelWidth = DEFAULT_VALIDATION_PANEL_WIDTH
    
End Sub

Private Sub initializeContainers()
    Set pExtraFrames = F.dictionaries.Create(False)
    Set pValidationLabels = F.dictionaries.Create(False)
End Sub

Private Sub Class_Terminate()
    Call destroy
End Sub

Public Sub destroy()
    On Error Resume Next
    Call pForm.hide
    Call VBA.unload(pForm)
End Sub



'[SETTERS]
Public Function setId(value As String) As WTemplate
    Set setId = Me
    pId = value
End Function

Public Function setCaption(value As String) As WTemplate
    Set setCaption = Me
    pCaption = value
    Call updateCaption
End Function

Public Function setBorderColor(value As Long) As WTemplate
    Set setBorderColor = Me
    pBorderColor = value
    Call updateBorderColor
End Function

Public Function setHasTitleBar(value As Boolean) As WTemplate
    Set setHasTitleBar = Me
    pHasTitleBar = value
End Function

Public Function setHasButtonsPanel(value As Boolean) As WTemplate
    Set setHasButtonsPanel = Me
    pHasButtonsPanel = value
End Function

Public Function setValidation(value As ValidationObject) As WTemplate
    Set setValidation = Me
    Set pValidation = value
    If Not pButtonsPanel Is Nothing Then
        Call pButtonsPanel.setStatus(pValidation.getStatus)
    End If
End Function

Public Function setValidationPanelWidth(value As Single) As WTemplate
    Set setValidationPanelWidth = Me
    pValidationPanelWidth = value
    Call updateValidationPanelWidth
End Function

Public Function setContentFrameHorizontalMargin(value As Single) As WTemplate
    Set setContentFrameHorizontalMargin = Me
    pContentFrameHorizontalMargin = value
End Function

Public Function setContentFrameVerticalMargin(value As Single) As WTemplate
    Set setContentFrameVerticalMargin = Me
    pContentFrameVerticalMargin = value
End Function

Public Function setContentFrameTopMargin(value As Single) As WTemplate
    Set setContentFrameTopMargin = Me
    pContentFrameTopMargin = value
End Function

Public Function setMinWidth(value As Single) As WTemplate
    Set setMinWidth = Me
    pMinWidth = value
End Function

Public Function setMaxWidth(value As Single) As WTemplate
    Set setMaxWidth = Me
    pMaxWidth = value
End Function

Public Function setMinHeight(value As Single) As WTemplate
    Set setMinHeight = Me
    pMinHeight = value
End Function

Public Function setMaxHeight(value As Single) As WTemplate
    Set setMaxHeight = Me
    pMaxHeight = value
End Function

Public Function setHelpFileTag(value As String) As WTemplate
    Set setHelpFileTag = Me
    pHelpFileTag = value
End Function

Public Function setHasUnsavedData(value As Boolean) As WTemplate
    Set setHasUnsavedData = Me
    pHasUnsavedData = value
    If pDisplayWarningsForUnsavedData Then
        Call pButtonsPanel.setHasUnsavedData(value)
    End If
End Function

Public Function setWarningsForUnsavedData(value As Boolean) As WTemplate
    Set setWarningsForUnsavedData = Me
    pDisplayWarningsForUnsavedData = value
End Function

Public Function setOkCaption(value As String) As WTemplate
    Set setOkCaption = Me
    If VBA.Len(value) Then
        pOkCaption = value
        If Not pButtonsPanel Is Nothing Then
            Call pButtonsPanel.setOkCaption(pOkCaption)
        End If
    End If
End Function

Public Function setCancelCaption(value As String) As WTemplate
    Set setCancelCaption = Me
    If VBA.Len(value) Then
        pCancelCaption = value
        If Not pButtonsPanel Is Nothing Then
            Call pButtonsPanel.setCancelCaption(pCancelCaption)
        End If
    End If
End Function

Public Function setHasCancelButton(value As Boolean) As WTemplate
    Set setHasCancelButton = Me
    pHasCancelButton = value
End Function

Public Function setHasOkButton(value As Boolean) As WTemplate
    Set setHasOkButton = Me
    pHasOkButton = value
End Function

Public Sub configure(properties As Scripting.Dictionary)
    If F.dictionaries.isNonEmptyDictionary(properties) Then
        
        Set pProperties = properties
        
        '[Size and layout]
        Call F.dictionaries.assignValueFromDictToVariable(properties, pMinWidth, "size.minWidth")
        Call F.dictionaries.assignValueFromDictToVariable(properties, pMaxHeight, "size.maxHeight")
        Call F.dictionaries.assignValueFromDictToVariable(properties, pContentFrameHorizontalMargin, "padding")
        Call F.dictionaries.assignValueFromDictToVariable(properties, pContentFrameVerticalMargin, "padding")
        
        '[Borders]
        Call F.dictionaries.assignValueFromDictToVariable(properties, pHasBorder, "border.visible")
        Call F.dictionaries.assignValueFromDictToVariable(properties, pBorderColor, "border.color")
        
    End If
End Sub






'[GETTERS]
Public Function getId() As String
    getId = pId
End Function

Friend Function getForm() As ufTemplate
    Set getForm = pForm
End Function

Public Function isVisible() As Boolean
    If Not pForm Is Nothing Then
        isVisible = pForm.isVisible
    End If
End Function

Public Function getSizeAndLocation(Optional onlyVisiblePart As Boolean = False) As area
    With getSizeAndLocation
        If onlyVisiblePart Then
            .left = pBody.left + pForm.getLeft
            .top = pBody.top + pForm.getTop
            .width = pBody.width
            .height = pBody.height
        Else
            .left = pForm.getLeft
            .top = pForm.getTop
            .width = pForm.getWidth
            .height = pForm.getHeight
        End If
    End With
End Function

Public Function getDefaultLocationProps() As DLocation
    Set getDefaultLocationProps = UI.Create.locationProps
    With getDefaultLocationProps
        Call .setAnchorPoint(AnchorPoint_TopLeft)
        Call .setAnchorParent(Me)
        Call .setOffsetLeft(DEFAULT_SUBFORM_OFFSET)
        Call .setOffsetTop(DEFAULT_SUBFORM_OFFSET)
    End With
End Function

Public Function getPadding() As Single
    getPadding = WINDOW_PADDING
End Function

Public Function getBody() As MSForms.frame
    Set getBody = pBody
End Function

Public Function getContentFrame() As MSForms.frame
    Set getContentFrame = pContentFrame
End Function

Public Function getExtraFrame(tag As String) As MSForms.frame
    Set getExtraFrame = F.dictionaries.getObject(pExtraFrames, tag)
End Function

Public Function getValidationFrame() As MSForms.frame
    Set getValidationFrame = pValidationFrame
End Function

Public Function getValidationPanelWidth() As Single
    getValidationPanelWidth = pValidationPanelWidth
End Function

Public Function getContentFrameHorizontalMargin() As Single
    getContentFrameHorizontalMargin = pContentFrameHorizontalMargin
End Function

Public Function getContentFrameWorkingAreaWidth() As Single
    getContentFrameWorkingAreaWidth = pContentFrame.InsideWidth - 2 * pContentFrameHorizontalMargin
End Function

Public Function getContentFrameVerticalMargin() As Single
    getContentFrameVerticalMargin = pContentFrameVerticalMargin
End Function

Public Function getHelpFileTag() As String
    getHelpFileTag = pHelpFileTag
End Function

Public Function hasDocumentation() As Boolean
    hasDocumentation = (VBA.Len(pHelpFileTag) > 0)
End Function

Public Function hasUnsavedChanges() As Boolean
    hasUnsavedChanges = pHasUnsavedData
End Function

Public Function displayWarningWhenClosingWithUnsavedData() As Boolean
    displayWarningWhenClosingWithUnsavedData = pDisplayWarningsForUnsavedData
End Function

Public Function isCancelled() As Boolean
    isCancelled = pIsCancelled
End Function






'[RENDERING]
Public Sub display(Optional mode As FormShowConstants = vbModal, Optional hasBackground As Boolean = True)
    If mode = vbModal And hasBackground Then
        pBackgroundActive = True
        Call pBackgroundWindow.show
    Else
        Call pForm.show(mode)
    End If
End Sub

Public Sub render()
    If Not pIsRendered Then
        Set pBackgroundWindow = createBackgroundWindow
        Call insertCanvas
        Call insertBody
        Call insertTitlePanel
        Call insertContentFrame
        Call insertExtraFrames
        Call insertButtonsPanel
        Call insertValidationPanel
        Call arrangeComponents
        pIsRendered = True
    End If
End Sub

Private Function createBackgroundWindow() As WBackground
    Set createBackgroundWindow = New WBackground
    With createBackgroundWindow
        Call .setWindow(Me)
    End With
End Function

Private Sub insertCanvas()
    Set pCanvas = pForm.controls.add(CREATE_FRAME_ID, CANVAS_FRAME_NAME, True)
    With pCanvas
        .backColor = TRANSPARENCY_LAYER_COLOR
        .SpecialEffect = fmSpecialEffectFlat
        .borderStyle = fmBorderStyleNone
        .top = 0
        .left = 0
        .width = pForm.getWidth
        .height = pForm.getHeight
    End With
End Sub

Private Sub insertBody()
    Set pBody = pCanvas.controls.add(CREATE_FRAME_ID, BODY_FRAME_NAME, True)
    With pBody
        .backColor = vbWhite
        .SpecialEffect = fmSpecialEffectFlat
        .borderStyle = fmBorderStyleSingle
        .borderColor = pBorderColor
        .top = 0
        .left = 0
        .width = pForm.getWidth
        .height = pForm.getHeight
    End With
End Sub

Private Sub insertTitlePanel()
    If pHasTitleBar Then
        'UI.Create.[] is not used in this case, because for unknown reasons if this component is
        'created by UI.Create.[] it doesn't handle controls events.
        Set pTitlePanel = New UITitlePanel
        With pTitlePanel
            Call .setWindow(Me)
            Call .setWidth(VBA.IIf(pHasTitleBar, DEFAULT_WIDTH, 0))
            Call .setBackColor(pBorderColor)
            Call .setCaption(pCaption)
            Call .configure(F.dictionaries.getObject(pProperties, TITLE_PANEL_TAG))
            Call .render
        End With
    End If
End Sub


Private Sub insertContentFrame()
    Set pContentFrame = pBody.controls.add(CREATE_FRAME_ID, CONTENT_FRAME_NAME, True)
    With pContentFrame
        .backColor = vbWhite
        .borderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .width = pMaxWidth
        .height = DEFAULT_CONTENT_HEIGHT
    End With
End Sub

Private Sub insertExtraFrames()
    Dim key As Variant
    Dim frame As MSForms.frame
    Dim props As Scripting.Dictionary
    '----------------------------------------------------------------------------------------------------------
    
    For Each key In pExtraFrames.keys
        Set props = pExtraFrames.item(key)
        Set frame = pBody.controls.add(CREATE_FRAME_ID, VBA.CStr(key), True)
        Call applyPropsForExtraFrame(frame, props)
        Set pExtraFrames.item(key) = frame
    Next key
    
End Sub

Private Sub applyPropsForExtraFrame(frame As MSForms.frame, props As Scripting.Dictionary)
    Dim propTypes As Variant
    Dim prop As Variant
    Dim value As Variant
    '----------------------------------------------------------------------------------------------------------
    
    '[Default props]
    With frame
        .backColor = EXTRA_FRAME_DEFAULT_BACK_COLOR
        .borderStyle = EXTRA_FRAME_DEFAULT_BORDER_STYLE
        .ScrollBars = EXTRA_FRAME_DEFAULT_SCROLL_BARS
    End With
    
    '[Custom props]
    propTypes = VBA.Array(UI.props.backgroundColor, UI.props.borderStyle, UI.props.ScrollBars)
    If F.dictionaries.isNonEmptyDictionary(props) Then
        For Each prop In propTypes
            If props.Exists(prop) Then
                Call UI.controls.setProperty(frame, UI.Cast.toProp(prop), props.item(prop))
            End If
        Next prop
    End If
    
End Sub

Private Sub insertButtonsPanel()
    'UI.Create.[] is not used in this case, because for unknown reasons if this component is
    'created by UI.Create.[] it doesn't handle controls events.
    Set pButtonsPanel = New UIButtonsPanel
    With pButtonsPanel
        Call .setWindow(Me)
        Call .setWidth(DEFAULT_WIDTH)
        Call .setOkCaption(pOkCaption)
        Call .setCancelCaption(pCancelCaption)
        Call .setHasCancelButton(pHasCancelButton)
        Call .setHasOkButton(pHasOkButton)
        Call .render
    End With
End Sub

Private Sub insertValidationPanel()
    If pHasValidationPanel Then
        Set pValidationFrame = pCanvas.controls.add(CREATE_FRAME_ID, VALIDATION_FRAME_NAME, True)
        With pValidationFrame
            .backColor = TRANSPARENCY_LAYER_COLOR
            .borderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .width = pValidationPanelWidth
            .height = 0
        End With
    End If
End Sub




Public Sub arrangeComponents()
    Dim top As Single:                          top = 0
    Dim width As Single
    Dim extraFramesBottomRightPoint As Coordinate
    '----------------------------------------------------------------------------------------------------------
    
    If Not pContentFrame Is Nothing Then
        Call adjustContentFrameSize
    End If
    
    extraFramesBottomRightPoint = getExtraFramesBottomRightPoint
    width = calculateWidth(extraFramesBottomRightPoint)
    
    '[Title]
    If Not pTitlePanel Is Nothing Then
        With pTitlePanel
            If .isVisible Then
                Call .setTop(top)
                Call .setWidth(width + PIXEL_SIZE)
                top = .getBottom
            End If
        End With
    End If
    
    '[Content]
    If Not pContentFrame Is Nothing Then
        With pContentFrame
            If .visible Then
                .top = top + pContentFrameTopMargin
                If .width > width Then
                    .ScrollBars = VBA.IIf(.ScrollBars <= fmScrollBarsHorizontal, fmScrollBarsHorizontal, fmScrollBarsBoth)
                    .ScrollWidth = .width
                    .width = width
                Else
                    If .width < width And .width > extraFramesBottomRightPoint.x Then
                        .width = width
                    End If
                    .ScrollBars = VBA.IIf(.ScrollBars <= fmScrollBarsHorizontal, fmScrollBarsNone, fmScrollBarsVertical)
                End If
                top = .top + .height
            End If
        End With
    End If
    
    '[Extra frames]
    If extraFramesBottomRightPoint.Y > top Then
        top = UI.Normalize.size(extraFramesBottomRightPoint.Y)
    End If
    
    '[Validation]
    If Not pValidationFrame Is Nothing Then
        With pValidationFrame
            .height = pContentFrame.height
            .width = pValidationPanelWidth
            .top = pContentFrame.top
            .left = width + WINDOW_PADDING
        End With
    End If
    
    '[Buttons]
    If Not pButtonsPanel Is Nothing Then
        Call pButtonsPanel.setIsVisible(pHasButtonsPanel)
        If pHasButtonsPanel Then
            With pButtonsPanel
                Call .setTop(top - PIXEL_SIZE)
                Call .setWidth(width)
                top = .getBottom
            End With
        End If
    End If
    
    With pBody
        .height = top + PIXEL_SIZE
        .width = width
        .top = PIXEL_SIZE
        .left = PIXEL_SIZE
    End With
    
    With pCanvas
        .height = pBody.height + pBody.top
        .width = pValidationFrame.left + pValidationFrame.width
        .top = 0
        .left = 0
    End With
    
    Call updateUserFormSize
    
End Sub



Private Sub adjustContentFrameSize()
    Dim rightBottomPoint As Coordinate
    Dim width As Single
    '----------------------------------------------------------------------------------------------------------
    
    rightBottomPoint = UI.Forms.calculateRightBottomPoint(pContentFrame)
    With pContentFrame
        .height = rightBottomPoint.Y + pContentFrameVerticalMargin
        
        '[Vertical scroll bar]
        If .height < pMinHeight Then
            .height = pMinHeight
            .ScrollBars = VBA.IIf(.ScrollBars = fmScrollBarsNone Or .ScrollBars = fmScrollBarsVertical, fmScrollBarsNone, fmScrollBarsHorizontal)
        ElseIf .height > pMaxHeight Then
            .ScrollBars = VBA.IIf(.ScrollBars > fmScrollBarsNone, fmScrollBarsBoth, fmScrollBarsVertical)
            .ScrollHeight = .height
            .height = pMaxHeight
        Else
            .ScrollBars = VBA.IIf(.ScrollBars = fmScrollBarsNone Or .ScrollBars = fmScrollBarsVertical, fmScrollBarsNone, fmScrollBarsHorizontal)
            .scrollTop = 0
        End If
        
        width = rightBottomPoint.x + pContentFrameHorizontalMargin + VBA.IIf(.ScrollBars >= fmScrollBarsVertical, SCROLL_BAR_WIDTH, 0)
        
        If pMinWidth > 0 And width < pMinWidth Then
            width = pMinWidth
        ElseIf pMaxWidth > 0 And width > pMaxWidth Then
            width = pMaxWidth
        End If
        
        .width = width
        
    End With
     
End Sub

Private Function calculateWidth(extraFramesBottomRightPoint As Coordinate) As Single
    calculateWidth = F.Numerics.getMax(pContentFrame.width, extraFramesBottomRightPoint.x)
    If calculateWidth < pMinWidth Then
        calculateWidth = pMinWidth
    ElseIf calculateWidth > pMaxWidth Then
        calculateWidth = pMaxWidth
    End If
End Function

Private Function getExtraFramesBottomRightPoint() As Coordinate
    Dim key As Variant
    Dim frame As MSForms.frame
    Dim width As Single
    Dim height As Single
    '----------------------------------------------------------------------------------------------------------
    
    With getExtraFramesBottomRightPoint
        For Each key In pExtraFrames.keys
            Set frame = pExtraFrames.item(key)
            If UI.controls.isValid(frame) Then
                width = frame.left + frame.width
                height = frame.top + frame.height
                If width > .x Then .x = width
                If height > .Y Then .Y = height
            End If
        Next key
        
        .x = .x + pContentFrameHorizontalMargin
        .Y = .Y + pContentFrameVerticalMargin
        
    End With
    
End Function

Private Sub updateValidationPanelWidth()
    If UI.controls.isValid(pValidationFrame) Then
        pValidationFrame.width = pValidationPanelWidth
    End If
End Sub

Private Sub updateUserFormSize()
    With pForm
        Call .setHeight(.getHeight - .InsideHeight + pCanvas.height + 4 * PIXEL_SIZE)
        Call .setWidth(.getWidth - .InsideWidth + pCanvas.width + 2 * PIXEL_SIZE)
    End With
End Sub

Private Sub updateBorderColor()
    If pIsRendered Then
        pBody.borderColor = pBorderColor
    End If
End Sub

Public Function getInsideSectionTop() As Single
    If Not pTitlePanel Is Nothing Then
        getInsideSectionTop = pTitlePanel.getBottom + pContentFrameTopMargin
    Else
        getInsideSectionTop = pContentFrameTopMargin
    End If
End Function

Private Sub updateCaption()
    If Not pTitlePanel Is Nothing Then
        Call pTitlePanel.setCaption(pCaption)
    End If
End Sub




'[POSITION]
Private Sub setMoveModeOn(ByVal x As Single, ByVal Y As Single)
    With pMoveCache
        .isActive = True
        .x = x
        .Y = Y
    End With
End Sub

Private Sub moveForm(ByVal x As Single, ByVal Y As Single)
    Dim xOffset As Single
    Dim yOffset As Single
    Dim left As Single
    Dim top As Single
    '----------------------------------------------------------------------------------------------------------
    
    With pMoveCache
        If .isActive Then
            xOffset = x - .x
            yOffset = Y - .Y
            With pForm
                left = .getLeft + xOffset
                top = .getTop + yOffset
                Call .setLeft(left)
                Call .setTop(top)
            End With
            
            If pBackgroundActive Then
                With pBackgroundWindow
                    Call .setLeft(left + PIXEL_SIZE)
                    Call .setTop(top + PIXEL_SIZE)
                End With
            End If
        End If
    End With
End Sub

Private Sub setMoveModeOff()
    With pMoveCache
        .isActive = False
    End With
End Sub



'[EXTRA CONTENTS]
Public Sub addExtraFrame(tagName As String, Optional frameProps As Scripting.Dictionary)
    Call F.dictionaries.addItem(pExtraFrames, tagName, frameProps)
End Sub

Public Sub addExtraButton(id As String, caption As String, Optional backgroundColor As Long = 13273922, _
                            Optional fontColor As Long = vbWhite)
    If Not pButtonsPanel Is Nothing Then
        Call pButtonsPanel.addExtraButton(id, caption, backgroundColor, fontColor)
    End If
End Sub


'[VISIBILITY]
Public Sub hide()
    Call pForm.hide
End Sub




Private Sub pForm_AfterDisplayed()
    RaiseEvent AfterDisplayed
End Sub

'[Events]
Private Sub pTitlePanel_AfterClose()
    Call ErrorManager.clear
    Call triggerCancel
End Sub

Private Sub pTitlePanel_HelpFileButtonClicked()
    Call ErrorManager.clear
    RaiseEvent HelpFileButtonClicked
End Sub

Private Sub pTitlePanel_MoveModeOn(ByVal x As Single, ByVal Y As Single)
    Call ErrorManager.clear
    Call setMoveModeOn(x, Y)
End Sub

Private Sub pTitlePanel_Move(ByVal x As Single, ByVal Y As Single)
    Call ErrorManager.clear
    Call moveForm(x, Y)
End Sub

Private Sub pTitlePanel_MoveModeOff(ByVal x As Single, ByVal Y As Single)
    Call ErrorManager.clear
    Call setMoveModeOff
End Sub

Public Sub triggerConfirm()
    pIsCancelled = False
    Call pForm.hide
    If pBackgroundActive Then
        Call pBackgroundWindow.hide
    End If
    RaiseEvent Confirm
End Sub

Public Sub triggerCancel()
    pIsCancelled = True
    Call pForm.hide
    If pBackgroundActive Then
        Call pBackgroundWindow.hide
    End If
    RaiseEvent Cancel
    Call unload(pForm)
End Sub

Public Sub triggerCustomAction(actionTag As String, Optional params As Scripting.Dictionary)
    RaiseEvent CustomAction(actionTag, params)
End Sub




'[VALIDATION]
Public Function insertValidationLabel(key As Variant) As UIValidationLabel
    Set insertValidationLabel = UI.Create.validationLabel(pValidationFrame)
    Call F.dictionaries.addItem(pValidationLabels, key, insertValidationLabel)
    Call insertValidationLabel.render
End Function




'[PANEL EVENTS]
Private Sub pButtonsPanel_Cancel()
    Call ErrorManager.clear
    Call triggerCancel
End Sub

Private Sub pButtonsPanel_Confirm()
    Call ErrorManager.clear
    Call triggerConfirm
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, closeMode As Integer)
    Call ErrorManager.clear
    If pForm.isVisible Then
        Call triggerCancel
    End If
End Sub



Public Sub sendToBackground()
    Dim width As Single
    Dim height As Single
    '----------------------------------------------------------------------------------------------------------
    
    width = UI.controls.getRight(pBody)
    height = UI.controls.getBottom(pBody)
    
    If pBackgroundActive Then
        With pBackgroundWindow
            Call .setWidth(width - PIXEL_SIZE)
            Call .setHeight(height - PIXEL_SIZE)
            Call .makeOpaque
        End With
    End If
    
    With pForm
        Call .setWidth(width)
        Call .setHeight(height)
    End With
    
    Call UI.Forms.SetFormOpacity(pForm, BACKGROUND_OPACITY)
    
End Sub

Public Sub restoreFromBackground()
    If pBackgroundActive Then
        Call pBackgroundWindow.makeTransparent
    End If
    Call UI.Forms.makeUserFormTransparent(pForm)
    Call updateUserFormSize
End Sub

Public Sub locateRelativeTo(locationProps As DLocation)
    Dim topLeftPoint As Coordinate
    '----------------------------------------------------------------------------------------------------------
    
    '[Calculate left and top coordinate]
    If Not locationProps Is Nothing Then
        With locationProps
            If .isDefined Then
                topLeftPoint = .getTopLeftPointForUserForm(Me)
                
                '[Apply coordinates calculates above to this user form]
                With pForm
                    .StartUpPosition = 0
                    .left = topLeftPoint.x
                    .top = topLeftPoint.Y
                End With
                
            End If
        End With
    End If
    
End Sub

Public Sub centerOnScreen()
    Dim screenArea As area
    Dim left As Single
    Dim top As Single
    '----------------------------------------------------------------------------------------------------------
    
    screenArea = F.Screen.getSize
    
    '[Calculate left and top coordinate]
    With screenArea
        left = (.width - pForm.getWidth) / 2
        top = (.height - pForm.getHeight) / 2
    End With
    
    '[Apply coordinates calculated above to this user form]
    With pForm
        Call .setLeft(left)
        Call .setTop(top)
    End With
    
End Sub

Public Sub toFront()
    Call pForm.toFront
End Sub




'[Implementation of IGridParentWindow]

Private Function IGridParentWindow_getDefaultLocationProps() As DLocation
    Set IGridParentWindow_getDefaultLocationProps = getDefaultLocationProps
End Function

Private Sub IGridParentWindow_restoreFromBackground()
    Call restoreFromBackground
End Sub

Private Sub IGridParentWindow_sendToBackground()
    Call sendToBackground
End Sub





'[Implementation of IListener interface]

Private Sub IListener_catchEvent(src As Variant, ev As Variant, Optional properties As Scripting.IDictionary)
    Stop
End Sub
