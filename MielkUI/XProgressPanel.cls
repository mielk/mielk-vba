VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XProgressPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum ControlTypeEnum
    ControlType_Container = 1
    ControlType_MainProgressBarHeader = 2
    ControlType_MainProgressBarFrame = 3
    ControlType_MainProgressBarProgress = 4
    ControlType_MainProgressBarCaption = 5
    ControlType_DetailProgressBarHeader = 6
    ControlType_DetailProgressBarFrame = 7
    ControlType_DetailProgressBarProgress = 8
    ControlType_DetailProgressBarCaption = 9
    ControlType_BottomBorder = 10
End Enum

Private Const CLASS_NAME As String = "XProgressPanel"
'[UI constants] -------------------------------------------------------------------------------------------
Private Const DEFAULT_MAIN_PROGRESS_BAR_HEADER_CAPTION As String = "Item {0} ({1} of {2})"
Private Const DEFAULT_DETAIL_PROGRESS_BAR_HEADER_CAPTION As String = "Processing item {0}"
Private Const DEFAULT_WIDTH As Single = 402
Private Const HEADER_LABEL_FONT_FAMILY As String = APP_FONT_FAMILY ' APP_FONT_FAMILY
Private Const HEADER_LABEL_FONT_SIZE As Long = 9
Private Const HEADER_LABEL_HEIGHT As Single = 12
Private Const HEADER_LABEL_WIDTH As Single = 100
Private Const HEADER_LABEL_LEFT_OFFSET As Single = 0
Private Const MAIN_HEADER_LABEL_TOP_OFFSET As Single = 0
Private Const DETAIL_HEADER_LABEL_TOP_OFFSET As Single = 0
Private Const DEFAULT_PROGRESS_BAR_HEIGHT As Single = 18
Private Const DEFAULT_PROGRESS_BACK_COLOR As Long = vbWhite
Private Const DEFAULT_PROGRESS_FRONT_COLOR As Long = 12300693
Private Const DEFAULT_PROGRESS_BORDER_COLOR As Long = &H80000006
Private Const DEFAULT_TOP_PADDING As Single = 6
Private Const DEFAULT_BOTTOM_PADDING As Single = 12
Private Const DEFAULT_INSIDE_PADDING As Single = 3
Private Const DEFAULT_BACK_STYLE As Variant = MSForms.fmBackStyleTransparent
Private Const DEFAULT_BACK_COLOR As Long = &H8000000F
Private Const DEFAULT_BORDER_STYLE As Variant = MSForms.fmBorderStyleSingle
Private Const DEFAULT_BORDER_COLOR As Long = &H80000006
Private Const DEFAULT_PROGRESS_CAPTION_FONT_FAMILY As String = APP_FONT_FAMILY ' APP_FONT_FAMILY
Private Const DEFAULT_PROGRESS_CAPTION_FONT_SIZE As Long = 8
Private Const DEFAULT_PROGRESS_CAPTION_FONT_COLOR As Long = vbBlack
Private Const DEFAULT_HAS_BOTTOM_BORDER_ONLY As Boolean = True
Private Const DEFAULT_HAS_DETAILS_PROGRESS_BAR As Boolean = True
Private Const PERCENT_LABEL_FORMAT As String = "#,#0%"
'[External UI components] ---------------------------------------------------------------------------------
Private pWindow As WTemplate
Private pController As IProgressControlController
'[Internal UI components] ---------------------------------------------------------------------------------
Private pContainerFrame As MSForms.frame
Private pMainProgressBarHeaderLabel As MSForms.label
Private pMainProgressBarFrameLabel As MSForms.label
Private pMainProgressBarProgressLabel As MSForms.label
Private pMainProgressBarCaptionLabel As MSForms.label
Private pDetailProgressBarHeaderLabel As MSForms.label
Private pDetailProgressBarFrameLabel As MSForms.label
Private pDetailProgressBarProgressLabel As MSForms.label
Private pDetailProgressBarCaptionLabel As MSForms.label
Private pBottomBorder As MSForms.label
'[UI properties] ------------------------------------------------------------------------------------------
Private pHasDetailsProgressBar As Boolean
Private pTop As Single
Private pLeft As Single
Private pWidth As Single
Private pProgressBarHeight As Single
Private pProgressBackColor As Long
Private pProgressFrontColor As Long
Private pProgressBorderColor As Long
Private pProgressCaptionFontFamily As String
Private pProgressCaptionFontSize As Single
Private pProgressCaptionFontColor As Long
Private pMainProgressBarCaption As String
Private pDetailProgressBarCaption As String
Private pTopPadding As Single
Private pBottomPadding As Single
Private pInsidePadding As Single
Private pBackStyle As MSForms.fmBackStyle
Private pBackColor As Long
Private pBorderStyle As MSForms.fmBorderStyle
Private pBorderColor As Long
Private pHasBottomBorderOnly As Boolean
Private pHeaderLabelHeight As Single
Private pHeaderFontFamily As String
Private pHeaderFontSize As Single
'[State] --------------------------------------------------------------------------------------------------
Private pUuid As String
Private pIsRendered As Boolean
Private pCurrentItemName As String
Private pCurrentItemIndex As Long
Private pCurrentItemDone As Boolean
Private pTotalItems As Long
Private pDetailAction As String
Private pDetailProgress As Single
'----------------------------------------------------------------------------------------------------------



'[INITIALIZE]
Private Sub Class_Initialize()
    Call initializeContainers
    Call initializeDefaultValues
End Sub

Private Sub initializeContainers()
End Sub

Private Sub initializeDefaultValues()
    pUuid = F.Crypto.createUUID
    pHasDetailsProgressBar = DEFAULT_HAS_DETAILS_PROGRESS_BAR
    pWidth = DEFAULT_WIDTH
    pProgressBarHeight = DEFAULT_PROGRESS_BAR_HEIGHT
    pProgressBackColor = DEFAULT_PROGRESS_BACK_COLOR
    pProgressFrontColor = DEFAULT_PROGRESS_FRONT_COLOR
    pProgressBorderColor = DEFAULT_PROGRESS_BORDER_COLOR
    pProgressCaptionFontFamily = DEFAULT_PROGRESS_CAPTION_FONT_FAMILY
    pProgressCaptionFontSize = DEFAULT_PROGRESS_CAPTION_FONT_SIZE
    pProgressCaptionFontColor = DEFAULT_PROGRESS_CAPTION_FONT_COLOR
    pMainProgressBarCaption = DEFAULT_MAIN_PROGRESS_BAR_HEADER_CAPTION
    pDetailProgressBarCaption = DEFAULT_DETAIL_PROGRESS_BAR_HEADER_CAPTION
    pTopPadding = DEFAULT_TOP_PADDING
    pBottomPadding = DEFAULT_BOTTOM_PADDING
    pInsidePadding = DEFAULT_INSIDE_PADDING
    pBackStyle = DEFAULT_BACK_STYLE
    pBackColor = DEFAULT_BACK_COLOR
    pBorderStyle = DEFAULT_BORDER_STYLE
    pBorderColor = DEFAULT_BORDER_COLOR
    pHasBottomBorderOnly = DEFAULT_HAS_BOTTOM_BORDER_ONLY
    pHeaderLabelHeight = HEADER_LABEL_HEIGHT
    pHeaderFontSize = HEADER_LABEL_FONT_SIZE
    pHeaderFontFamily = HEADER_LABEL_FONT_FAMILY
End Sub



'[SETTERS]
Public Function setWindow(value As WTemplate) As XProgressPanel
    Set setWindow = Me
    Set pWindow = value
End Function

Public Function setFrame(value As MSForms.frame) As XProgressPanel
    Set setFrame = Me
    Set pContainerFrame = value
End Function

Public Function setController(value As IProgressControlController) As XProgressPanel
    Set setController = Me
    Set pController = value
End Function

Public Function setHasDetailsProgressBar(value As Boolean) As XProgressPanel
    Set setHasDetailsProgressBar = Me
    pHasDetailsProgressBar = value
End Function



'[Position]
Public Function setTop(value As Single) As XProgressPanel
    Set setTop = Me
    pTop = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBottom(value As Single) As XProgressPanel
    Set setBottom = Me
    pTop = value - getHeight
    If pIsRendered Then Call arrangeControls
End Function

Public Function setLeft(value As Single) As XProgressPanel
    Set setLeft = Me
    pLeft = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setRight(value As Single) As XProgressPanel
    Set setRight = Me
    pLeft = value - pWidth
    If pIsRendered Then Call arrangeControls
End Function


'[Size]
Public Function setWidth(value As Single) As XProgressPanel
    Set setWidth = Me
    pWidth = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setHeight(value As Single) As XProgressPanel
    Set setHeight = Me
    Call VBA.Err.Raise(1)
End Function


'[Captions]
Public Function setMainHeaderCaption(value As String) As XProgressPanel
    Set setMainHeaderCaption = Me
    pMainProgressBarCaption = value
    If pIsRendered Then Call refreshProgressBarHeaderLabel(ControlType_MainProgressBarHeader)
End Function

Public Function setDetailHeaderCaption(value As String) As XProgressPanel
    Set setDetailHeaderCaption = Me
    pDetailProgressBarCaption = value
    If pIsRendered Then Call refreshProgressBarHeaderLabel(ControlType_DetailProgressBarHeader)
End Function


'[Headers]
Public Function setHeaderLabelHeight(value As Single) As XProgressPanel
    Set setHeaderLabelHeight = Me
    pHeaderLabelHeight = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setHeaderFontFamily(value As String) As XProgressPanel
    Set setHeaderFontFamily = Me
    pHeaderFontFamily = value
    If pIsRendered Then
        Call refreshProgressBarHeaderLabel(ControlType_MainProgressBarHeader)
        Call refreshProgressBarHeaderLabel(ControlType_DetailProgressBarHeader)
        Call arrangeControls
    End If
End Function

Public Function setHeaderFontSize(value As Single) As XProgressPanel
    Set setHeaderFontSize = Me
    pHeaderFontSize = value
    If pIsRendered Then
        Call refreshProgressBarHeaderLabel(ControlType_MainProgressBarHeader)
        Call refreshProgressBarHeaderLabel(ControlType_DetailProgressBarHeader)
        Call arrangeControls
    End If
End Function


'[Backcolor and borders]
Public Function setBackStyle(value As MSForms.fmBackStyle) As XProgressPanel
    Set setBackStyle = Me
    pBackStyle = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function

Public Function setBackColor(value As Long) As XProgressPanel
    Set setBackColor = Me
    pBackColor = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function

Public Function setBorderStyle(value As MSForms.fmBorderStyle) As XProgressPanel
    Set setBorderStyle = Me
    pBorderStyle = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function

Public Function setBorderColor(value As Long) As XProgressPanel
    Set setBorderColor = Me
    pBorderColor = value
    If pIsRendered Then Call refreshBackgroundFrame
End Function



'[Actual progress bar layout]
Public Function setProgressBarHeight(value As Single) As XProgressPanel
    Set setProgressBarHeight = Me
    pProgressBarHeight = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setProgressBackColor(value As Long) As XProgressPanel
    Set setProgressBackColor = Me
    pProgressBackColor = value
    If pIsRendered Then
        Call refreshProgressBarFrameLabel(ControlType_MainProgressBarFrame)
        Call refreshProgressBarFrameLabel(ControlType_DetailProgressBarFrame)
    End If
End Function

Public Function setProgressFrontColor(value As Long) As XProgressPanel
    Set setProgressFrontColor = Me
    pProgressFrontColor = value
    If pIsRendered Then
        Call refreshProgressBarProgressLabel(ControlType_MainProgressBarProgress)
        Call refreshProgressBarProgressLabel(ControlType_DetailProgressBarProgress)
    End If
End Function

Public Function setProgressBorderColor(value As Long) As XProgressPanel
    Set setProgressBorderColor = Me
    pProgressBorderColor = value
    If pIsRendered Then
        Call refreshProgressBarFrameLabel(ControlType_MainProgressBarFrame)
        Call refreshProgressBarFrameLabel(ControlType_DetailProgressBarFrame)
    End If
End Function

Public Function setProgressCaptionFontFamily(value As String) As XProgressPanel
    Set setProgressCaptionFontFamily = Me
    pProgressCaptionFontFamily = value
    If pIsRendered Then
        Call refreshProgressBarCaptionLabel(ControlType_MainProgressBarCaption)
        Call refreshProgressBarCaptionLabel(ControlType_DetailProgressBarCaption)
    End If
End Function

Public Function setProgressCaptionFontSize(value As Single) As XProgressPanel
    Set setProgressCaptionFontSize = Me
    pProgressCaptionFontSize = value
    If pIsRendered Then
        Call refreshProgressBarCaptionLabel(ControlType_MainProgressBarCaption)
        Call refreshProgressBarCaptionLabel(ControlType_DetailProgressBarCaption)
    End If
End Function

Public Function setProgressCaptionFontColor(value As Long) As XProgressPanel
    Set setProgressCaptionFontColor = Me
    pProgressCaptionFontColor = value
    If pIsRendered Then
        Call refreshProgressBarCaptionLabel(ControlType_MainProgressBarCaption)
        Call refreshProgressBarCaptionLabel(ControlType_DetailProgressBarCaption)
    End If
End Function



'[Layout]
Public Function setTopPadding(value As Single) As XProgressPanel
    Set setTopPadding = Me
    pTopPadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBottomPadding(value As Single) As XProgressPanel
    Set setBottomPadding = Me
    pBottomPadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setInsidePadding(value As Single) As XProgressPanel
    Set setInsidePadding = Me
    pInsidePadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setHasBottomBorderOnly(value As Boolean) As XProgressPanel
    Set setHasBottomBorderOnly = Me
    pHasBottomBorderOnly = value
    If pIsRendered Then Call arrangeControls
End Function



'[GETTERS]
Public Function isValid() As Boolean
    isValid = UI.controls.isValid(pContainerFrame)
End Function

Public Function hasDetailsProgressBar() As Boolean
    hasDetailsProgressBar = pHasDetailsProgressBar
End Function

'[Position]
Public Function getLeft() As Single
    getLeft = pContainerFrame.left
End Function

Public Function getRight() As Single
    getRight = UI.controls.getRight(pContainerFrame)
End Function

Public Function getTop() As Single
    getTop = pContainerFrame.top
End Function

Public Function getBottom() As Single
    getBottom = UI.controls.getBottom(pContainerFrame)
End Function


'[Size]
Public Function getWidth() As Single
    getWidth = pContainerFrame.width
End Function

Public Function getHeight() As Single
    getHeight = pContainerFrame.height
End Function


'[Other properties]
Private Function getMainProgressBarHeader() As String
    getMainProgressBarHeader = F.Strings.format(pMainProgressBarCaption, pCurrentItemName, pCurrentItemIndex, pTotalItems)
End Function

Private Function getMainProgressNumeric() As String
    getMainProgressNumeric = VBA.format(getCurrentMainProgress, PERCENT_LABEL_FORMAT)
End Function

Private Function getMainProgressBarWidth() As Single
    getMainProgressBarWidth = getCurrentMainProgress * pMainProgressBarFrameLabel.width
End Function

Private Function getCurrentMainProgress() As Single
    If pTotalItems Then
        getCurrentMainProgress = (pCurrentItemIndex - VBA.IIf(pCurrentItemDone, 0, 1)) / pTotalItems
    End If
End Function

Private Function getDetailProgressBarHeader() As String
    getDetailProgressBarHeader = pDetailAction
End Function

Private Function getDetailProgressNumeric() As String
    getDetailProgressNumeric = VBA.format(pDetailProgress, PERCENT_LABEL_FORMAT)
End Function

Private Function getDetailProgressBarWidth() As Single
    getDetailProgressBarWidth = pDetailProgress * pDetailProgressBarFrameLabel.width
End Function



'[API]
Public Sub render()
    Call refreshBackgroundFrame
    Call insertMainProgressBar
    Call insertDetailProgressBar
    Call insertBottomBorder
    Call arrangeControls
    pIsRendered = True
End Sub

Public Sub Move(Optional left As Variant, Optional top As Variant)
    If left <> pLeft Or top <> pTop Then
        If F.Variables.isOfNumericType(left) Then pLeft = left
        If F.Variables.isOfNumericType(top) Then pTop = top
        Call arrangeControls
    End If
End Sub

Public Sub resize(Optional width As Variant, Optional height As Variant)
    If pWidth <> width Then
        If width >= 0 Then pWidth = width
        Call arrangeControls
    End If
End Sub

Public Sub startNewMainItem(itemName As String)
    pCurrentItemName = itemName
    pCurrentItemIndex = pCurrentItemIndex + 1
    pCurrentItemDone = False
    pTotalItems = pController.getTotalItemsCounter
    pDetailProgress = 0
    Call updateMainProgressBar
    Call updateDetailProgressBar
End Sub

Public Sub endCurrentMainItem()
    pCurrentItemDone = True
    pDetailProgress = 1
    Call updateMainProgressBar
    Call updateDetailProgressBar
End Sub

Public Function updateDetailProgress(action As String, progress As Single) As XProgressPanel
    If VBA.Len(action) Then pDetailAction = action
    If F.Math.isInRange(progress, 0, 1) Then pDetailProgress = progress
    Call updateDetailProgressBar
End Function




'[RENDERING]
Private Function getLabelName(controlType As ControlTypeEnum) As String
    Dim tag As String
    '----------------------------------------------------------------------------------------------------------
    tag = CLASS_NAME & "|" & pUuid & "|" & controlType
    getLabelName = "ctrl_" & VBA.left$(F.Crypto.SHA1HASH(tag), 25)
End Function

Private Function getControl(ctrlType As ControlTypeEnum) As MSForms.control
    Dim name As String
    '----------------------------------------------------------------------------------------------------------
    On Error Resume Next
    If ctrlType = ControlType_Container Then
        Set getControl = pContainerFrame
    Else
        name = getLabelName(ctrlType)
        Set getControl = pContainerFrame.controls(name)
    End If
End Function


'[Inserting controls]


Private Sub insertMainProgressBar()
    Set pMainProgressBarHeaderLabel = insertProgressBarHeaderLabel(ControlType_MainProgressBarHeader, getMainProgressBarHeader)
    Set pMainProgressBarFrameLabel = insertProgressBarFrameLabel(ControlType_MainProgressBarFrame)
    Set pMainProgressBarProgressLabel = insertProgressBarProgressLabel(ControlType_MainProgressBarProgress)
    Set pMainProgressBarCaptionLabel = insertProgressBarCaptionLabel(ControlType_MainProgressBarCaption)
End Sub

Private Sub insertDetailProgressBar()
    Set pDetailProgressBarHeaderLabel = insertProgressBarHeaderLabel(ControlType_DetailProgressBarHeader, getDetailProgressBarHeader)
    Set pDetailProgressBarFrameLabel = insertProgressBarFrameLabel(ControlType_DetailProgressBarFrame)
    Set pDetailProgressBarProgressLabel = insertProgressBarProgressLabel(ControlType_DetailProgressBarProgress)
    Set pDetailProgressBarCaptionLabel = insertProgressBarCaptionLabel(ControlType_DetailProgressBarCaption)
End Sub

Private Function insertProgressBarHeaderLabel(ctrlType As ControlTypeEnum, caption As String) As MSForms.label
    Set insertProgressBarHeaderLabel = pContainerFrame.controls.add(CREATE_LABEL_ID, getLabelName(ctrlType), True)
    Call refreshProgressBarHeaderLabel(ctrlType, caption)
End Function

Private Function insertProgressBarFrameLabel(ctrlType As ControlTypeEnum) As MSForms.label
    Set insertProgressBarFrameLabel = pContainerFrame.controls.add(CREATE_LABEL_ID, getLabelName(ctrlType), True)
    Call refreshProgressBarFrameLabel(ctrlType)
End Function

Private Function insertProgressBarProgressLabel(ctrlType As ControlTypeEnum) As MSForms.label
    Set insertProgressBarProgressLabel = pContainerFrame.controls.add(CREATE_LABEL_ID, getLabelName(ctrlType), True)
    Call refreshProgressBarProgressLabel(ctrlType)
End Function

Private Function insertProgressBarCaptionLabel(ctrlType As ControlTypeEnum) As MSForms.label
    Set insertProgressBarCaptionLabel = pContainerFrame.controls.add(CREATE_LABEL_ID, getLabelName(ctrlType), True)
    Call refreshProgressBarCaptionLabel(ctrlType)
End Function

Private Sub insertBottomBorder()
    Set pBottomBorder = pContainerFrame.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_BottomBorder), True)
    Call refreshBottomBorder
End Sub



'[Refreshing controls]
Private Sub refreshBackgroundFrame()
    With pContainerFrame
        .backColor = pBackColor
        .borderStyle = VBA.IIf(pHasBottomBorderOnly, MSForms.fmBorderStyleNone, pBorderStyle)
        .borderColor = pBorderColor
        .SpecialEffect = fmSpecialEffectFlat
        .caption = vbNullString
    End With
End Sub

Private Sub refreshProgressBarHeaderLabel(ctrlType As ControlTypeEnum, Optional caption As Variant)
    Dim ctrl As MSForms.control
    '------------------------------------------------------------------------------------------------------
    Set ctrl = getControl(ctrlType)
    With ctrl
        .BackStyle = MSForms.fmBackStyleTransparent
        .borderStyle = MSForms.fmBorderStyleNone
        .font.name = pHeaderFontFamily
        .font.size = pHeaderFontSize
        If Not VBA.IsMissing(caption) Then
            .caption = F.Strings.stringify(caption)
        End If
    End With
End Sub

Private Sub refreshProgressBarFrameLabel(ctrlType As ControlTypeEnum)
    Dim ctrl As MSForms.control
    '------------------------------------------------------------------------------------------------------
    Set ctrl = getControl(ctrlType)
    With ctrl
        .BackStyle = MSForms.fmBackStyleOpaque
        .backColor = pProgressBackColor
        .borderStyle = MSForms.fmBorderStyleSingle
        .borderColor = pProgressBorderColor
    End With
End Sub

Private Sub refreshProgressBarProgressLabel(ctrlType As ControlTypeEnum)
    Dim ctrl As MSForms.control
    '------------------------------------------------------------------------------------------------------
    Set ctrl = getControl(ctrlType)
    With ctrl
        .BackStyle = MSForms.fmBackStyleOpaque
        .backColor = pProgressFrontColor
        .borderStyle = MSForms.fmBorderStyleSingle
        .borderColor = pProgressBorderColor
    End With
End Sub

Private Sub refreshProgressBarCaptionLabel(ctrlType As ControlTypeEnum)
    Dim ctrl As MSForms.control
    '------------------------------------------------------------------------------------------------------
    Set ctrl = getControl(ctrlType)
    With ctrl
        .BackStyle = MSForms.fmBackStyleTransparent
        .borderStyle = MSForms.fmBorderStyleNone
        .borderColor = pProgressBorderColor
        .font.name = pProgressCaptionFontFamily
        .font.size = pProgressCaptionFontSize
        .ForeColor = pProgressCaptionFontColor
        .TextAlign = fmTextAlignCenter
    End With
End Sub

Private Sub refreshBottomBorder()
    With pBottomBorder
        .BackStyle = MSForms.fmBackStyleOpaque
        .backColor = pBorderColor
        .borderStyle = MSForms.fmBorderStyleNone
        .visible = pHasBottomBorderOnly
    End With
End Sub



'[Arranging controls]
Private Sub arrangeControls()
    Dim top As Single
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    
    Call resizeControls
    
    Call UI.controls.locate(pMainProgressBarHeaderLabel, pInsidePadding, pTopPadding)
    top = UI.controls.getBottom(pMainProgressBarHeaderLabel) + pInsidePadding
    Call UI.controls.locate(pMainProgressBarFrameLabel, pInsidePadding, top)
    Call UI.controls.locate(pMainProgressBarProgressLabel, pInsidePadding, top)
    Call UI.controls.locate(pMainProgressBarCaptionLabel, pInsidePadding, pMainProgressBarFrameLabel.top + (pProgressBarHeight - pMainProgressBarCaptionLabel.height) / 2)
    
    Call UI.controls.locate(pDetailProgressBarHeaderLabel, pInsidePadding, UI.controls.getBottom(pMainProgressBarFrameLabel) + pInsidePadding)
    top = UI.controls.getBottom(pDetailProgressBarHeaderLabel) + pInsidePadding
    Call UI.controls.locate(pDetailProgressBarFrameLabel, pInsidePadding, top)
    Call UI.controls.locate(pDetailProgressBarProgressLabel, pInsidePadding, top)
    Call UI.controls.locate(pDetailProgressBarCaptionLabel, pInsidePadding, pDetailProgressBarFrameLabel.top + (pProgressBarHeight - pDetailProgressBarCaptionLabel.height) / 2)
    
    Call UI.controls.locate(pBottomBorder, 0, UI.controls.getBottom(pDetailProgressBarFrameLabel) + pBottomPadding)
    Call UI.controls.locate(pContainerFrame, pLeft, pTop)
    Call autoSizeBackgroundFrame
    
End Sub

Private Sub resizeControls()
    Call UI.controls.resize(pContainerFrame, pWidth, 100)
    Call UI.controls.resize(pMainProgressBarHeaderLabel, pWidth - 2 * pInsidePadding, pHeaderLabelHeight)
    Call UI.controls.resize(pDetailProgressBarHeaderLabel, pWidth - 2 * pInsidePadding, pHeaderLabelHeight)
    Call UI.controls.resize(pMainProgressBarFrameLabel, pWidth - 2 * pInsidePadding, pProgressBarHeight)
    Call UI.controls.resize(pDetailProgressBarFrameLabel, pWidth - 2 * pInsidePadding, pProgressBarHeight)
    Call UI.controls.resize(pMainProgressBarProgressLabel, getMainProgressBarWidth, pProgressBarHeight)
    Call UI.controls.resize(pDetailProgressBarProgressLabel, getDetailProgressBarWidth, pProgressBarHeight)
    Call resizeProgressCaptionLabel(pMainProgressBarCaptionLabel)
    Call resizeProgressCaptionLabel(pDetailProgressBarCaptionLabel)
    Call UI.controls.resize(pBottomBorder, pWidth, 1)
End Sub

Private Sub resizeProgressCaptionLabel(ctrl As MSForms.control)
    With ctrl
        .width = pWidth - 2 * pInsidePadding
        .AutoSize = True
        .AutoSize = False
        .width = pWidth - 2 * pInsidePadding
    End With
End Sub

Private Sub autoSizeBackgroundFrame()
    Dim endPoint As Coordinate
    '----------------------------------------------------------------------------------------------------------
    Call hideShowDetailsProgressBar
    endPoint = UI.Forms.calculateRightBottomPoint(pContainerFrame, True)
    pContainerFrame.height = endPoint.Y + pTopPadding
End Sub

Private Sub hideShowDetailsProgressBar()
    On Error Resume Next
    pDetailProgressBarHeaderLabel.visible = pHasDetailsProgressBar
    pDetailProgressBarFrameLabel.visible = pHasDetailsProgressBar
    pDetailProgressBarProgressLabel.visible = pHasDetailsProgressBar
    pDetailProgressBarCaptionLabel.visible = pHasDetailsProgressBar
End Sub



'[Clearing]
Private Sub clear()
    Call UI.Forms.clearMsFrame(pContainerFrame)
End Sub



'[PROGRESS UPDATE]
Private Sub updateProgressBars()
    Call updateMainProgressBar
    Call updateDetailProgressBar
End Sub

Private Sub updateMainProgressBar()
    pMainProgressBarHeaderLabel.caption = getMainProgressBarHeader
    pMainProgressBarCaptionLabel.caption = getMainProgressNumeric
    pMainProgressBarProgressLabel.width = getMainProgressBarWidth
End Sub

Private Sub updateDetailProgressBar()
    pDetailProgressBarHeaderLabel.caption = getDetailProgressBarHeader
    pDetailProgressBarCaptionLabel.caption = getDetailProgressNumeric
    pDetailProgressBarProgressLabel.width = getDetailProgressBarWidth
End Sub
