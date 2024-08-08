VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIListHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IControl

Private Enum ControlTypeEnum
    ControlType_Background = 1
    ControlType_Caption = 2
    ControlType_Sorter = 3
    ControlType_Filter = 4
    ControlType_ToFirst = 5
    ControlType_ToPrevious = 6
    ControlType_CurrentPage = 7
    ControlType_TotalPages = 8
    ControlType_ToNext = 9
    ControlType_ToLast = 10
End Enum

Private Const CLASS_NAME As String = "UIListHeader"
'[Captions] ----------------------------------------------------------------------------------------------
Private Const PAGINATION_STATUS As String = "UiList.Pagination.Status"
'[External references] -----------------------------------------------------------------------------------
Private pParent As UIList
Private WithEvents pDefinition As DListHeader
Attribute pDefinition.VB_VarHelpID = -1
'[UI properties] -----------------------------------------------------------------------------------------
Private Const PAGINATION_BUTTON_SIZE As Single = 9
Private Const PAGINATION_CONTROLS_SPACE As Single = 3
Private Const PAGINATION_LABEL_FONT_SIZE As Single = 8
'[UI components] -----------------------------------------------------------------------------------------
Private pBackgroundLabel As MSForms.label
Private pCaptionLabel As MSForms.label
Private pSortButton As MSForms.label
'[UI pagination] -----------------------------------------------------------------------------------------
Private WithEvents pToFirstButton As MSForms.label
Attribute pToFirstButton.VB_VarHelpID = -1
Private WithEvents pToPreviousButton As MSForms.label
Attribute pToPreviousButton.VB_VarHelpID = -1
'Private pCurrentPage As MSForms.textBox
Private pPaginationStatus As MSForms.label
Private WithEvents pToNextButton As MSForms.label
Attribute pToNextButton.VB_VarHelpID = -1
Private WithEvents pToLastButton As MSForms.label
Attribute pToLastButton.VB_VarHelpID = -1
'[Size & location] ---------------------------------------------------------------------------------------
Private pLeft As Single
Private pWidth As Single
'[State] -------------------------------------------------------------------------------------------------
Private pUuid As String
Private pIsRendered As Boolean
Private pVisible As Boolean
'---------------------------------------------------------------------------------------------------------
Private pTotalPages As Long
Private pCurrentPage As Long
'---------------------------------------------------------------------------------------------------------



'[INITIALIZE]
Private Sub Class_Initialize()
    pUuid = F.Crypto.createUUID
    pVisible = True
End Sub




'[SETTERS]
Public Function setParent(value As UIList) As UIListHeader
    Set setParent = Me
    Set pParent = value
End Function

Public Function setDefinition(value As DListHeader) As UIListHeader
    Set setDefinition = Me
    Set pDefinition = value
End Function

Public Function setWidth(value As Single) As UIListHeader
    Set setWidth = Me
    pWidth = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setWidthShare(value As Single) As UIListHeader
    Set setWidthShare = Me
    pWidth = value * getFrame.InsideWidth
    If pIsRendered Then Call arrangeControls
End Function

Public Function setLeft(value As Single) As UIListHeader
    Set setLeft = Me
    pLeft = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setVisibility(value As Boolean) As UIListHeader
    Set setVisibility = Me
    pVisible = value
    If pIsRendered Then Call updateVisibility
End Function

Public Function setTotalPages(value As Long) As UIListHeader
    Set setTotalPages = Me
    pTotalPages = value
    Call updatePaginationControlsLayout
End Function

Public Function setCurrentPage(value As Long) As UIListHeader
    Set setCurrentPage = Me
    pCurrentPage = value
    Call updatePaginationControlsLayout
End Function

Public Sub updatePaginationStatus(currentPage As Long, pagesCounter As Long)
    If currentPage Then pCurrentPage = currentPage
    pTotalPages = pagesCounter
    Call updatePaginationControlsLayout
End Sub




'[GETTERS]
Public Function getParent() As UIList
    Set getParent = pParent
End Function

Public Function getDefinition() As UIListHeader
    Set getDefinition = pDefinition
End Function

Public Function getFrame() As MSForms.frame
    Set getFrame = pParent.getHeaderFrame
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

Public Function getWidth() As Single
    getWidth = pBackgroundLabel.width
End Function

Public Function getHeight() As Single
    getHeight = pBackgroundLabel.height
End Function

Public Function getRectangle() As RECT
    With getRectangle
        .top = getTop
        .bottom = getBottom
        .left = getLeft
        .right = getRight
    End With
End Function

Public Function getCaption() As String
    getCaption = pDefinition.getCaption
End Function

Public Function getProperty() As Variant
    getProperty = pDefinition.getProperty
End Function



'[RENDERING]
Public Sub render()
    Call insertBackgroundLabel
    Call insertCaptionLabel
    Call insertPaginationLabels
    Call arrangeControls
    pIsRendered = True
End Sub

Public Function Move(Optional left As Variant, Optional top As Variant) As UIListHeader
    Set Move = Me
    If F.Variables.isOfNumericType(left) Then pLeft = left
    '[top cannot be changed from this level - it depends on parental list settings]
    Call arrangeControls
End Function

Public Function offset(Optional left As Variant, Optional top As Variant) As UIListHeader
    Set offset = Me
    If F.Variables.isOfNumericType(left) Then pLeft = pLeft + left
    '[top cannot be changed from this level - it depends on parental list settings]
    Call arrangeControls
End Function

Public Function resize(Optional width As Variant, Optional height As Variant) As UIListHeader
    Set resize = Me
    If F.Math.isNotNegative(width) Then pWidth = width
    '[height cannot be changed from this level - it depends on parental list settings]
    Call arrangeControls
End Function

Public Sub remove()
    Call clear
End Sub



'[RENDERING]
Private Function getLabelName(controlType As ControlTypeEnum) As String
    Dim tag As String
    '----------------------------------------------------------------------------------------------------------
    tag = CLASS_NAME & "|" & pUuid & "|" & controlType
    getLabelName = "ctrl_" & VBA.left$(F.Crypto.SHA1HASH(tag), 25)
End Function

Private Sub insertBackgroundLabel()
    Set pBackgroundLabel = getFrame.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_Background), True)
    Call refreshBackgroundLabel
End Sub

Private Sub insertCaptionLabel()
    Set pCaptionLabel = getFrame.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_Caption), True)
    Call refreshCaptionLabel
End Sub

Private Sub insertPaginationLabels()
    Dim frame As MSForms.frame:                     Set frame = getFrame
    '----------------------------------------------------------------------------------------------------------
    
    Set pToFirstButton = frame.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_ToFirst), True)
    Set pToPreviousButton = frame.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_ToPrevious), True)
    'Set pCurrentPage = frame.controls.Add(CREATE_TEXTBOX_ID, getLabelName(ControlType_CurrentPage), False)
    Set pPaginationStatus = frame.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_TotalPages), True)
    Set pToNextButton = frame.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_ToNext), True)
    Set pToLastButton = frame.controls.add(CREATE_LABEL_ID, getLabelName(ControlType_ToLast), True)
    
    Call formatPaginationButton(pToFirstButton, UI.Icons.getArrowDoubleLeft)
    Call formatPaginationButton(pToPreviousButton, UI.Icons.getArrowLeft)
    Call formatPaginationButton(pToNextButton, UI.Icons.getArrowRight)
    Call formatPaginationButton(pToLastButton, UI.Icons.getArrowDoubleRight)
'
'    With pCurrentPage
'        .height = 32
'        .width = 100
'        .backColor = vbRed
'    End With
'
    With pPaginationStatus
        .font.size = PAGINATION_BUTTON_SIZE
        .font.name = APP_FONT_FAMILY
        .ForeColor = vbBlack
        .font.Bold = 1
    End With
    
End Sub

Private Sub formatPaginationButton(btn As MSForms.label, picture As Variant)
    With btn
        .width = PAGINATION_BUTTON_SIZE
        .height = PAGINATION_BUTTON_SIZE
        .borderStyle = fmBorderStyleNone
        Set .picture = picture
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = UI.Icons.getMouseIcon
    End With
End Sub




Private Sub refreshBackgroundLabel()
    With pBackgroundLabel
        .BackStyle = MSForms.fmBackStyleTransparent
        .borderStyle = MSForms.fmBorderStyleNone
    End With
End Sub

Private Sub refreshCaptionLabel()
    With pCaptionLabel
        .BackStyle = MSForms.fmBackStyleTransparent
        .borderStyle = MSForms.fmBorderStyleNone
        .font.size = pParent.getHeaderFontSize
        .font.name = pParent.getHeaderFontFamily
        .font.Bold = pParent.getHeaderFontBold
        .ForeColor = pParent.getHeaderFontColor
        .caption = pDefinition.getCaption
        .width = getFrame.InsideWidth
        .AutoSize = True
    End With
End Sub




Private Sub clear()
    On Error Resume Next
    With getFrame.controls
        Call .remove(getLabelName(ControlType_Background))
        Call .remove(getLabelName(ControlType_Caption))
    End With
End Sub

Private Sub arrangeControls()
    Call resizeControls
    Call UI.controls.locate(pBackgroundLabel, pLeft, 0)
    Call UI.controls.locate(pCaptionLabel, pLeft + pParent.getPadding, (pBackgroundLabel.height - pCaptionLabel.height) / 2)
    Call updatePaginationControlsLayout
End Sub

Private Sub updatePaginationControlsLayout()
    Dim buttonTop As Single
    Dim left As Single
    Dim ctrl As MSForms.control
    '----------------------------------------------------------------------------------------------------------

    If Not pVisible Then
        Call updatePaginationVisibility(False)
    ElseIf pTotalPages <= 1 Then
        Call updatePaginationVisibility(False)
    Else
        Call updatePaginationVisibility(True)
        Call updatePaginationStatusCaption
        buttonTop = UI.controls.getBottom(pBackgroundLabel) - pPaginationStatus.height + (pPaginationStatus.height - PAGINATION_BUTTON_SIZE) / 2
        
        Call UI.controls.locate(pToLastButton, pBackgroundLabel.width - PAGINATION_BUTTON_SIZE, buttonTop)
        Call UI.controls.locate(pToNextButton, pToLastButton.left - PAGINATION_CONTROLS_SPACE - PAGINATION_BUTTON_SIZE, buttonTop)
        Call UI.controls.locate(pPaginationStatus, pToNextButton.left - PAGINATION_CONTROLS_SPACE - pPaginationStatus.width, _
                                                        UI.controls.getBottom(pBackgroundLabel) - pPaginationStatus.height)
        Call UI.controls.locate(pToPreviousButton, pPaginationStatus.left - PAGINATION_CONTROLS_SPACE - PAGINATION_BUTTON_SIZE, buttonTop)
        Call UI.controls.locate(pToFirstButton, pToPreviousButton.left - PAGINATION_CONTROLS_SPACE - PAGINATION_BUTTON_SIZE, buttonTop)
        
        '[Adjust visibility]
        On Error Resume Next
        If pTotalPages = pCurrentPage Then
            pToNextButton.visible = False
            pToLastButton.visible = False
        ElseIf pCurrentPage = 1 Then
            pToFirstButton.visible = False
            pToPreviousButton.visible = False
        End If
        
    End If
    

End Sub

Private Sub updatePaginationStatusCaption()
    Dim caption As String
    '----------------------------------------------------------------------------------------------------------
    caption = F.Strings.format(MsgService.getText(PAGINATION_STATUS), pCurrentPage, pTotalPages)
    With pPaginationStatus
        .width = pBackgroundLabel.width
        .caption = caption
        .AutoSize = True
        .AutoSize = False
    End With
End Sub

Private Sub resizeControls()
    Call UI.controls.resize(pBackgroundLabel, pWidth, getFrame.InsideHeight)
    Call resizeCaptionLabel
End Sub

Private Sub resizeCaptionLabel()
    With pCaptionLabel
        .width = pWidth
        .AutoSize = True
        .AutoSize = False
    End With
End Sub

Public Sub updateVisibility()
    pBackgroundLabel.visible = pVisible
    pCaptionLabel.visible = pVisible
    Call updatePaginationControlsLayout
End Sub



'[Pagination]
Public Sub updatePaginationVisibility(value As Boolean)
    On Error Resume Next
    pToFirstButton.visible = value
    pToPreviousButton.visible = value
    pPaginationStatus.visible = value
    pToNextButton.visible = value
    pToLastButton.visible = value
End Sub





Private Sub pDefinition_AfterCaptionChanged()
    pCaptionLabel.caption = pDefinition.getCaption
    Call arrangeControls
End Sub




'[Events of pagination buttons]
Private Sub pToFirstButton_Click()
    Call moveToPage(1)
End Sub

Private Sub pToPreviousButton_Click()
    Call moveToPage(pCurrentPage - 1)
End Sub

Private Sub pToNextButton_Click()
    Call moveToPage(pCurrentPage + 1)
End Sub

Private Sub pToLastButton_Click()
    Call moveToPage(pTotalPages)
End Sub

Private Sub moveToPage(pageIndex As Long)
    pCurrentPage = pageIndex
    Call updatePaginationControlsLayout
    Call pParent.changePage(pCurrentPage)
End Sub




'[Implementation of UIControl]
Private Sub IControl_setVisibility(value As Boolean)
    Call setVisibility(value)
End Sub

Private Sub IControl_setEnability(value As Boolean)
    'N/A
End Sub

Private Sub IControl_setWidth(value As Single)
    Call setWidth(value)
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

Private Function IControl_move(Optional x As Variant, Optional Y As Variant) As IControl
    Set IControl_move = Move(x, Y)
End Function

Private Function IControl_offset(Optional x As Single, Optional Y As Single) As IControl
    Set IControl_offset = offset(x, Y)
End Function

Private Sub IControl_triggerEvent(eventId As Long)
    'N/A
End Sub
