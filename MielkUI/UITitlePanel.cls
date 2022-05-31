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
Private Const UNSAVED_DATA_QUESTION As String = "ConfirmCancellingWindowWithUnsavedData"
'[UI default values] --------------------------------------------------------------------------------------
Private Const DEFAULT_IS_VISIBLE As Boolean = True
Private Const DEFAULT_HEIGHT As Single = 32
Private Const DEFAULT_PADDING As Single = 12
Private Const DEFAULT_MOVE_ENABLED As Boolean = True
Private Const DEFAULT_CLOSE_ENABLED As Boolean = True
Private Const DEFAULT_BACKGROUND_COLOR As Long = 11630657
Private Const DEFAULT_TITLE_BAR_FONT_FAMILY As String = "Segoe UI"
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




'[GETTERS]
Public Function isVisible() As Boolean
    isVisible = pIsVisible
End Function

Public Function getBottom() As Single
    If UI.Forms.isControlValid(pContainer) Then
        With pContainer
            getBottom = .top + .height
        End With
    End If
End Function





'[RENDERING]
Public Function render() As UITitlePanel
    Call insertContainer
    Call insertTitleLabel
    Call insertMoveLayer
    Call insertCloseButton
    Call insertDocumentationButton
End Function


Private Sub insertContainer()
    If pContainer Is Nothing Then
        Set pContainer = pFrame.controls.Add(CREATE_FRAME_ID, FRAME_NAME, True)
    End If
    With pContainer
        .SpecialEffect = fmSpecialEffectFlat
        .BackColor = pBackgroundColor
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
        Set pCaptionLabel = pContainer.controls.Add(CREATE_LABEL_ID, CAPTION_LABEL_NAME, True)
    End If
    With pCaptionLabel
        .caption = getTranslatedContent
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
        .width = pContainer.InsideWidth - 2 * pPadding
        .AutoSize = True
        .height = .height + PIXEL_SIZE
        .top = (pContainer.InsideHeight - .height) / 2
        .width = pContainer.InsideWidth - 2 * pPadding
    End With
End Sub

Private Function getTranslatedContent() As String
    If F.Strings.startsWith(pCaption, "#") Then
        getTranslatedContent = MsgService.getText(VBA.Mid$(pCaption, 2))
    Else
        getTranslatedContent = pCaption
    End If
End Function


Private Sub insertMoveLayer()
    If pMoveLayer Is Nothing Then
        Set pMoveLayer = pContainer.Add(CREATE_LABEL_ID, MOVE_LAYER_NAME, True)
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
        Set pCloseButton = pContainer.Add(CREATE_LABEL_ID, CLOSE_BUTTON_NAME, True)
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
        Set pHelpFileButton = pContainer.Add(CREATE_LABEL_ID, CLOSE_BUTTON_NAME, False)
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
    
    If UI.Forms.isControlValid(pCloseButton) Then
        left = pCloseButton.left
    Else
        left = pContainer.InsideWidth
    End If
    calculateDocumentationButtonLeft = left - pCloseButtonMargin - pCloseButton.width + PIXEL_SIZE
    
End Function

Private Function calculateCaptionLabelWidth() As Single
    Dim left As Single
    '----------------------------------------------------------------------------------------------------------
    
    If UI.Forms.isControlValid(pHelpFileButton) Then
        calculateCaptionLabelWidth = pHelpFileButton.left - pCaptionLabel.left - pPadding
    ElseIf UI.Forms.isControlValid(pHelpFileButton) Then
        calculateCaptionLabelWidth = pCloseButton.left - pCaptionLabel.left - pPadding
    Else
        calculateCaptionLabelWidth = pContainer.InsideWidth - 2 * pPadding
    End If
    
End Function



'[SIZE & POSITION]
Private Sub updatePosition()
    If UI.Forms.isControlValid(pContainer) Then
        With pContainer
            .left = pLeft
            .top = F.Numerics.roundTo(VBA.CDbl(pTop), PIXEL_SIZE)
        End With
    End If
End Sub

Private Sub updateSize()
    If UI.Forms.isControlValid(pContainer) Then
        pContainer.width = pWidth
        Call centerTitleVertically
        pMoveLayer.width = pContainer.width
        pCloseButton.left = calculateCloseButtonLeft
        pHelpFileButton.left = calculateDocumentationButtonLeft
        pCaptionLabel.width = calculateCaptionLabelWidth
    End If
End Sub

Private Sub updateCaption()
    If UI.Forms.isControlValid(pCaptionLabel) Then
        pCaptionLabel.caption = getTranslatedContent
    End If
End Sub

Private Sub updateBackColor()
    If UI.Forms.isControlValid(pContainer) Then
        pContainer.BackColor = pBackgroundColor
    End If
End Sub








'[EVENTS]
Private Sub pCloseButton_Click()
    Call ErrorManager.clear

    If confirmSkippingNotSavedData Then
        RaiseEvent AfterClose
    End If

End Sub

Private Function confirmSkippingNotSavedData() As Boolean
    If Not pWindow.hasUnsavedChanges Then
        confirmSkippingNotSavedData = F.Dialogs.getUserConfirmation(MsgService.getText(UNSAVED_DATA_QUESTION))
    Else
        confirmSkippingNotSavedData = True
    End If
End Function

Private Sub pHelpFileButton_Click()
    Call ErrorManager.clear
    RaiseEvent HelpFileButtonClicked
End Sub

Private Sub pMoveLayer_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ErrorManager.clear
    If button = 1 Then RaiseEvent MoveModeOn(x, y)
End Sub

Private Sub pMoveLayer_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ErrorManager.clear
    If button = 1 Then RaiseEvent Move(x, y)
End Sub

Private Sub pMoveLayer_MouseUp(ByVal button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ErrorManager.clear
    If button = 1 Then RaiseEvent MoveModeOff(x, y)
End Sub
