VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UITitlePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "UITitlePanel"
'[Components names] ---------------------------------------------------------------------------------------
Private Const FRAME_NAME As String = "frmTitle"
Private Const CAPTION_LABEL_NAME As String = "lblCaption"
Private Const CLOSE_BUTTON_NAME As String = "lblCloseButton"
Private Const DOCUMENTATION_BUTTON_NAME As String = "lblDocumentationButton"
Private Const MOVE_LAYER_NAME As String = "lblMoveLayer"
Private Const UNSAVED_DATA_QUESTION As String = "Windows.Warnings.CancelWithoutSaving"
'[UI default values] --------------------------------------------------------------------------------------
Private Const DEFAULT_IS_VISIBLE As Boolean = True
Private Const DEFAULT_HEIGHT As Single = 32
Private Const DEFAULT_PADDING As Single = 12
Private Const DEFAULT_MOVE_ENABLED As Boolean = True
Private Const DEFAULT_CLOSE_ENABLED As Boolean = True
Private Const DEFAULT_BACKGROUND_COLOR As Long = 11630657
Private Const DEFAULT_TITLE_BAR_FONT_FAMILY As String = APP_FONT_FAMILY
Private Const DEFAULT_TITLE_BAR_FONT_SIZE As Single = 11
Private Const DEFAULT_TITLE_BAR_FONT_COLOR As Long = vbWhite
Private Const DEFAULT_TITLE_BAR_FONT_BOLD As Boolean = True
'[Close button]
Private Const CLOSE_BUTTON_SYMBOL As String = "Î"
Private Const CLOSE_BUTTON_FONT_FAMILY As String = "Wingdings 2"
Private Const CLOSE_BUTTON_HEIGHT As Single = 13.5
Private Const CLOSE_BUTTON_WIDTH As Single = 14.25
Private Const CLOSE_BUTTON_FONT_SIZE As Single = 13
Private Const DEFAULT_CLOSE_BUTTON_MARGIN As Single = 3
'[Documentation button]
Private Const DOCUMENTATION_BUTTON_SYMBOL As String = "?"
Private Const DOCUMENTATION_BUTTON_FONT_FAMILY As String = "Arial"
Private Const DOCUMENTATION_BUTTON_FONT_SIZE As Single = 11
'[UI components] ------------------------------------------------------------------------------------------
Private pWindow As WTemplate
Private pFrame As MSForms.frame
Private pContainer As MSForms.frame
Private WithEvents pCaptionLabel As MSForms.label
Attribute pCaptionLabel.VB_VarHelpID = -1
Private WithEvents pMoveLayer As MSForms.label
Attribute pMoveLayer.VB_VarHelpID = -1
Private WithEvents pCloseButton As MSForms.label
Attribute pCloseButton.VB_VarHelpID = -1
Private WithEvents pHelpFileButton As MSForms.label
Attribute pHelpFileButton.VB_VarHelpID = -1
Private pImages As Scripting.Dictionary
Private pActionIcons As Scripting.Dictionary
'[UI properties] ------------------------------------------------------------------------------------------
Private pIsVisible As Boolean
Private pHeight As Single
Private pWidth As Single
Private pTop As Single
Private pLeft As Single
Private pPadding As Single
Private pCaptionIndent As Single
Private pMoveEnabled As Boolean
Private pCloseEnabled As Boolean
Private pBackgroundColor As Long
Private pTitleBarFontFamily As String
Private pTitleBarFontSize As Single
Private pTitleBarFontColor As Long
Private pTitleBarFontBold As Boolean
Private pCloseButtonMargin As Single
Private pCaption As String
Private pImagesDefinitions As VBA.Collection
Private pActionIconsDefinitions As VBA.Collection
'[Events] -------------------------------------------------------------------------------------------------
Event AfterClose()
Event MoveModeOn(ByVal x As Single, ByVal y As Single)
Event Move(ByVal x As Single, ByVal y As Single)
Event MoveModeOff(ByVal x As Single, ByVal y As Single)
Event HelpFileButtonClicked()
'----------------------------------------------------------------------------------------------------------







Private Sub Class_Initialize()
    Call setDefaultValues
End Sub

Private Sub setDefaultValues()
    pIsVisible = DEFAULT_IS_VISIBLE
    pHeight = DEFAULT_HEIGHT
    pPadding = DEFAULT_PADDING
    pMoveEnabled = DEFAULT_MOVE_ENABLED
    pCloseEnabled = DEFAULT_CLOSE_ENABLED
    pBackgroundColor = DEFAULT_BACKGROUND_COLOR
    pTitleBarFontFamily = DEFAULT_TITLE_BAR_FONT_FAMILY
    pTitleBarFontSize = DEFAULT_TITLE_BAR_FONT_SIZE
    pTitleBarFontColor = DEFAULT_TITLE_BAR_FONT_COLOR
    pTitleBarFontBold = DEFAULT_TITLE_BAR_FONT_BOLD
    pCloseButtonMargin = DEFAULT_CLOSE_BUTTON_MARGIN
End Sub




'[SETTERS]

Public Function setWindow(value As WTemplate) As UITitlePanel
    Set setWindow = Me
    Set pWindow = value
    Set pFrame = pWindow.getBody
End Function

Public Function setHeight(value As Single) As UITitlePanel
    Set setHeight = Me
    pHeight = value
    Call updateSize
End Function

Public Function setWidth(value As Single) As UITitlePanel
    Set setWidth = Me
    pWidth = value
    Call updateSize
End Function

Public Function setTop(value As Single) As UITitlePanel
    Set setTop = Me
    pTop = value
    Call updatePosition
End Function

Public Function setLeft(value As Single) As UITitlePanel
    Set setLeft = Me
    pLeft = value
    Call updatePosition
End Function

Public Function setCaption(value As String) As UITitlePanel
    Set setCaption = Me
    pCaption = value
    Call updateCaption
End Function

Public Function setBackColor(value As Long) As UITitlePanel
    Set setBackColor = Me
    pBackgroundColor = value
    Call updateBackColor
End Function

Public Function configure(properties As Scripting.Dictionary) As UITitlePanel
    Set configure = Me
    Call F.dictionaries.assignValueFromDictToVariable(properties, pIsVisible, "visible")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pHeight, "height")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pMoveEnabled, "moveEnabled")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pCloseEnabled, "closeEnabled")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pBackgroundColor, "backgroundColor")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pTitleBarFontFamily, "font.family")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pTitleBarFontSize, "font.size")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pTitleBarFontColor, "font.color")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pTitleBarFontBold, "font.bold")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pCaption, "caption")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pCaptionIndent, "captionIndent")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pCloseButtonMargin, "closeButtonMargin")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pImagesDefinitions, "images")
    Call F.dictionaries.assignValueFromDictToVariable(properties, pActionIconsDefinitions, "actionIcons")
End Function




'[GETTERS]
Public Function isVisible() As Boolean
    isVisible = pIsVisible
End Function

Public Function getBottom() As Single
    If UI.controls.isValid(pContainer) Then
        With pContainer
            getBottom = .top + .height
        End With
    End If
End Function





'[RENDERING]
Public Function render() As UITitlePanel
    Call insertContainer
    Call insertTitleLabel
    Call insertImages
    Call insertMoveLayer
    Call insertCloseButton
    Call insertDocumentationButton
    Call insertActionIcons
End Function


Private Sub insertContainer()
    If pContainer Is Nothing Then
        Set pContainer = pFrame.controls.add(CREATE_FRAME_ID, FRAME_NAME, True)
    End If
    With pContainer
        .SpecialEffect = fmSpecialEffectFlat
        .backColor = pBackgroundColor
        .borderStyle = fmBorderStyleNone
        .left = 0
        .top = 0
        .width = pWidth
        .height = pHeight
        .MousePointer = fmMousePointerSizeAll
    End With
End Sub

Private Sub insertTitleLabel()
    If pCaptionLabel Is Nothing Then
        Set pCaptionLabel = pContainer.controls.add(CREATE_LABEL_ID, CAPTION_LABEL_NAME, True)
    End If
    With pCaptionLabel
        .caption = MsgService.getText(pCaption)
        With .font
            .name = pTitleBarFontFamily
            .Bold = pTitleBarFontBold
            .size = pTitleBarFontSize
        End With
        .ForeColor = pTitleBarFontColor
        .left = pPadding + pCaptionIndent
        .TextAlign = fmTextAlignLeft
        Call centerTitleVertically
        .MousePointer = fmMousePointerSizeAll
    End With
End Sub

Private Sub centerTitleVertically()
    With pCaptionLabel
        .AutoSize = False
        .width = F.Math.getMax(pContainer.InsideWidth - 2 * pPadding, 0)
        .AutoSize = True
        .height = .height + PIXEL_SIZE
        .top = (pContainer.InsideHeight - .height) / 2
        .width = F.Math.getMax(pContainer.InsideWidth - 2 * pPadding, 0)
    End With
End Sub


Private Sub insertMoveLayer()
    If pMoveLayer Is Nothing Then
        Set pMoveLayer = pContainer.add(CREATE_LABEL_ID, MOVE_LAYER_NAME, True)
    End If
    With pMoveLayer
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .width = pContainer.InsideWidth
        .height = pContainer.InsideHeight
        .left = 0
        .top = 0
        .MousePointer = fmMousePointerSizeAll
        .visible = pMoveEnabled
    End With
End Sub


Private Sub insertCloseButton()

    If pCloseButton Is Nothing Then
        Set pCloseButton = pContainer.add(CREATE_LABEL_ID, CLOSE_BUTTON_NAME, True)
    End If
    With pCloseButton
        .visible = pCloseEnabled
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .width = CLOSE_BUTTON_WIDTH
        .height = CLOSE_BUTTON_HEIGHT
        .left = calculateCloseButtonLeft
        .top = pCloseButtonMargin
        .caption = CLOSE_BUTTON_SYMBOL
        .ForeColor = pTitleBarFontColor
        With .font
            .name = "Arial"
            .name = CLOSE_BUTTON_FONT_FAMILY
            .size = CLOSE_BUTTON_FONT_SIZE
            .Bold = True
        End With
        .MousePointer = fmMousePointerCustom
        On Error Resume Next
        .MouseIcon = Icons.getMousePointer
    End With
End Sub

Private Function calculateCloseButtonLeft() As Single
    calculateCloseButtonLeft = pContainer.InsideWidth - pCloseButtonMargin - pCloseButton.width + PIXEL_SIZE
End Function

Private Sub insertDocumentationButton()
    If pHelpFileButton Is Nothing Then
        Set pHelpFileButton = pContainer.add(CREATE_LABEL_ID, CLOSE_BUTTON_NAME, False)
    End If
    With pHelpFileButton
        .visible = pWindow.hasDocumentation
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .width = CLOSE_BUTTON_WIDTH
        .height = CLOSE_BUTTON_HEIGHT
        .left = calculateDocumentationButtonLeft
        .top = pCloseButtonMargin
        .caption = DOCUMENTATION_BUTTON_SYMBOL
        .ForeColor = pTitleBarFontColor
        .TextAlign = fmTextAlignCenter
        With .font
            .name = "Arial"
            .name = DOCUMENTATION_BUTTON_FONT_FAMILY
            .size = DOCUMENTATION_BUTTON_FONT_SIZE
            .Bold = True
        End With
        .MousePointer = fmMousePointerCustom
        On Error Resume Next
        .MouseIcon = Icons.getMousePointer
        .visible = pWindow.hasDocumentation
    End With
End Sub

Private Function calculateDocumentationButtonLeft() As Single
    Dim left As Single
    '----------------------------------------------------------------------------------------------------------
    
    If UI.controls.isValid(pCloseButton) Then
        left = pCloseButton.left
    Else
        left = pContainer.InsideWidth
    End If
    calculateDocumentationButtonLeft = left - pCloseButtonMargin - pCloseButton.width + PIXEL_SIZE
    
End Function

Private Function calculateCaptionLabelWidth() As Single
    Dim left As Single
    '----------------------------------------------------------------------------------------------------------
    
    If UI.controls.isValid(pHelpFileButton) Then
        calculateCaptionLabelWidth = pHelpFileButton.left - pCaptionLabel.left - pPadding
    ElseIf UI.controls.isValid(pCloseButton) Then
        calculateCaptionLabelWidth = pCloseButton.left - pCaptionLabel.left - pPadding
    Else
        calculateCaptionLabelWidth = pContainer.InsideWidth - 2 * pPadding
    End If
    
End Function



Private Sub insertImages()
    Dim dto As Scripting.Dictionary
    Dim name As String
    Dim img As MSForms.label
    Dim imageProps As Scripting.Dictionary
    '----------------------------------------------------------------------------------------------------------
    
    Set pImages = F.dictionaries.Create(False)
    If F.Collections.isNonEmptyCollection(pImagesDefinitions) Then
        For Each dto In pImagesDefinitions
            name = F.dictionaries.getItem(dto, props.name.getName)
            Set img = insertImage(dto)
            If Not img Is Nothing Then
                Set imageProps = F.dictionaries.Clone(dto)
                Call imageProps.add(props.control.getName, img)
                Call pImages.add(name, imageProps)
            End If
        Next dto
    End If
    
End Sub

Private Function insertImage(dto As Scripting.Dictionary) As MSForms.label
    Dim name As String:             name = "img"
    Dim height As Single:           height = 15
    Dim width As Single:            width = 15
    '----------------------------------------------------------------------------------------------------------
    
    Call F.dictionaries.assignValueFromDictToVariable(dto, name, "name")
    Call F.dictionaries.assignValueFromDictToVariable(dto, height, "size.height")
    Call F.dictionaries.assignValueFromDictToVariable(dto, width, "size.width")
    
    Set insertImage = pContainer.controls.add(CREATE_LABEL_ID, name, True)
    With insertImage
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .width = width
        .height = height
        On Error Resume Next
        Set .picture = Icons.getIcon(name)
        On Error GoTo 0
    End With

    Call UI.controls.locateByDictionary(insertImage, dto, pContainer, pPadding)

End Function





Private Sub insertActionIcons()
    Dim dto As Scripting.Dictionary
    Dim actionIcon As UIIcon
    Dim name As String
    Dim iconProps As Scripting.Dictionary
    '----------------------------------------------------------------------------------------------------------
    
    Set pActionIcons = F.dictionaries.Create(False)
    If F.Collections.isNonEmptyCollection(pActionIconsDefinitions) Then
        For Each dto In pActionIconsDefinitions
            name = F.dictionaries.getItem(dto, props.name.getName)
            Set actionIcon = insertActionIcon(dto)
            If Not actionIcon Is Nothing Then
                Set iconProps = F.dictionaries.Clone(dto)
                Call iconProps.add(props.control.getName, actionIcon)
                Call pActionIcons.add(name, iconProps)
            End If
        Next dto
    End If
End Sub

Private Function insertActionIcon(dto As Scripting.Dictionary) As UIIcon
    Dim name As String:             name = "img"
    Dim height As Single:           height = 15
    Dim width As Single:            width = 15
    Dim tag As String
    '----------------------------------------------------------------------------------------------------------
    
    Call F.dictionaries.assignValueFromDictToVariable(dto, name, "name")
    Call F.dictionaries.assignValueFromDictToVariable(dto, height, "size.height")
    Call F.dictionaries.assignValueFromDictToVariable(dto, width, "size.width")
    Call F.dictionaries.assignValueFromDictToVariable(dto, tag, "tag")
    
    Set insertActionIcon = UI.Create.UIIcon(pContainer)
    With insertActionIcon
        Call .setTag(tag)
        Call .setPicture(Icons.getIcon(name))
        Call .setListener(pWindow)
        Call .setWidth(width)
        Call .setHeight(height)
        Call .render
    End With
    
    Call UI.controls.locateByDictionary(insertActionIcon, dto, pContainer, pPadding)
    
End Function





'[SIZE & POSITION]
Private Sub updatePosition()
    If UI.controls.isValid(pContainer) Then
        With pContainer
            .left = pLeft
            .top = pTop
        End With
    End If
End Sub

Private Sub updateSize()
    If UI.controls.isValid(pContainer) Then
        pContainer.width = F.Math.getMax(pWidth, 0)
        Call centerTitleVertically
        pMoveLayer.width = F.Math.getMax(pContainer.width, 0)
        pCloseButton.left = calculateCloseButtonLeft
        pHelpFileButton.left = calculateDocumentationButtonLeft
        pCaptionLabel.width = F.Math.getMax(calculateCaptionLabelWidth, 0)
        Call locateControls(pImages)
        Call locateControls(pActionIcons)
    End If
End Sub

Private Sub updateCaption()
    If UI.controls.isValid(pCaptionLabel) Then
        pCaptionLabel.caption = MsgService.getText(pCaption)
    End If
End Sub

Private Sub updateBackColor()
    If UI.controls.isValid(pContainer) Then
        pContainer.backColor = pBackgroundColor
    End If
End Sub

Private Sub locateControls(controlsDict As Scripting.Dictionary)
    Dim varKey As Variant
    Dim ctrl As Object
    Dim properties As Scripting.Dictionary
    '----------------------------------------------------------------------------------------------------------
    For Each varKey In controlsDict.keys
        Set properties = controlsDict.item(varKey)
        Set ctrl = F.dictionaries.getObject(properties, props.control.getName)
        If UI.controls.isValid(ctrl) Then
            Call UI.controls.locateByDictionary(ctrl, properties, pContainer, pPadding)
        End If
    Next varKey
End Sub







'[EVENTS]
Private Sub pCloseButton_Click()
    Call ErrorManager.clear

    If confirmSkippingNotSavedData Then
        RaiseEvent AfterClose
    End If

End Sub

Private Function confirmSkippingNotSavedData() As Boolean
    If Not pWindow.displayWarningWhenClosingWithUnsavedData Then
        confirmSkippingNotSavedData = True
    ElseIf Not pWindow.hasUnsavedChanges Then
        confirmSkippingNotSavedData = F.Dialogs.getUserConfirmation(MsgService.getText(UNSAVED_DATA_QUESTION))
    Else
        confirmSkippingNotSavedData = True
    End If
End Function

Private Sub pHelpFileButton_Click()
    Call ErrorManager.clear
    RaiseEvent HelpFileButtonClicked
End Sub

Private Sub pMoveLayer_MouseDown(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ErrorManager.clear
    If Button = 1 Then RaiseEvent MoveModeOn(x, y)
End Sub

Private Sub pMoveLayer_MouseMove(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ErrorManager.clear
    If Button = 1 Then RaiseEvent Move(x, y)
End Sub

Private Sub pMoveLayer_MouseUp(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ErrorManager.clear
    If Button = 1 Then RaiseEvent MoveModeOff(x, y)
End Sub
