VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIButtonsPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IListener

Private Const CLASS_NAME As String = "UIButtonsPanel"
'[UI default values] --------------------------------------------------------------------------------------
Private Const FRAME_NAME As String = "frmButtons"
Private Const OK_BUTTON_NAME As String = "btnOk"
Private Const CANCEL_BUTTON_NAME As String = "btnCancel"
Private Const TOP_BORDER_NAME As String = "lblTopBorder"
Private Const UNSAVED_DATA_ICON_NAME As String = "lblUnsavedDataIcon"
Private Const UNSAVED_DATA_CAPTION_NAME As String = "lblUnsavedDataCaption"
Private Const UNSAVED_DATA_TAG As String = "UnsavedData"
Private Const UNSAVED_DATA_QUESTION As String = "ConfirmCancellingWindowWithUnsavedData"
'[UI default values] --------------------------------------------------------------------------------------
Private Const DEFAULT_IS_VISIBLE As Boolean = True
Private Const DEFAULT_HEIGHT As Single = 40
Private Const DEFAULT_PADDING As Single = 12
Private Const DEFAULT_OK_ACTIVE As Boolean = True
Private Const DEFAULT_BACKGROUND_COLOR As Long = vbWhite
Private Const DEFAULT_BUTTON_HEIGHT As Single = 24
Private Const DEFAULT_BUTTON_WIDTH As Single = 60
Private Const DEFAULT_FONT_FAMILY As String = APP_FONT_FAMILY
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_FONT_COLOR As Long = vbBlack
Private Const DEFAULT_FONT_BOLD As Boolean = True
Private Const DEFAULT_BORDER_COLOR As Long = 14803425       '14079702
Private Const DEFAULT_BORDER_SPACE As Single = 6
Private Const DEFAULT_BUTTONS_SPACE As Single = 6
'...
Private Const UNSAVED_ICON_SIZE As Single = 18
Private Const UNSAVED_CAPTION_FONT_SIZE As Single = 10
Private Const UNSAVED_CAPTION_FONT_NAME As String = APP_FONT_FAMILY
Private Const UNSAVED_CAPTION_FONT_BOLD As Boolean = True
Private Const UNSAVED_CAPTION_FONT_COLOR As Long = 11991
'[UI components] ------------------------------------------------------------------------------------------
Private pWindow As WTemplate
Private pFrame As MSForms.frame
Private pContainer As MSForms.frame
Attribute pContainer.VB_VarHelpID = -1
Private pTopBorder As MSForms.label
Private WithEvents pOkButton As UIButton
Attribute pOkButton.VB_VarHelpID = -1
Private WithEvents pCancelButton As UIButton
Attribute pCancelButton.VB_VarHelpID = -1
Private pUnsavedDataIcon As MSForms.label
Private pUnsavedDataCaption As MSForms.label
Private pExtraButtons As VBA.Collection
'[UI properties] ------------------------------------------------------------------------------------------
Private pHeight As Single
Private pWidth As Single
Private pTop As Single
Private pLeft As Single
Private pPadding As Single
Private pBorderColor As Long
Private pBackgroundColor As Long
Private pButtonHeight As Single
Private pButtonWidth As Single
Private pFontFamily As String
Private pFontSize As Single
Private pFontColor As Long
Private pFontBold As Boolean
Private pButtonsSpace As Single
'[Captions]
Private pOkCaption As String
Private pCancelCaption As String
Private pHasOkButton As Boolean
Private pHasCancelButton As Boolean
'[State] --------------------------------------------------------------------------------------------------
Private pIsRendered As Boolean
Private pIsVisible As Boolean
Private pStatus As Boolean
Private pHasUnsavedData As Boolean
'----------------------------------------------------------------------------------------------------------
Event Confirm()
Event Cancel()
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Initialize()
    Call setDefaultValues
End Sub

Private Sub setDefaultValues()
    pIsVisible = DEFAULT_IS_VISIBLE
    pHeight = DEFAULT_HEIGHT
    pPadding = DEFAULT_PADDING
    pStatus = DEFAULT_OK_ACTIVE
    pBorderColor = DEFAULT_BORDER_COLOR
    pBackgroundColor = DEFAULT_BACKGROUND_COLOR
    pButtonHeight = DEFAULT_BUTTON_HEIGHT
    pButtonWidth = DEFAULT_BUTTON_WIDTH
    pFontFamily = DEFAULT_FONT_FAMILY
    pFontSize = DEFAULT_FONT_SIZE
    pFontBold = DEFAULT_FONT_BOLD
    pFontColor = DEFAULT_FONT_COLOR
    pButtonsSpace = DEFAULT_BUTTONS_SPACE
    Set pExtraButtons = New VBA.Collection
End Sub



'[SETTERS]

Public Function setWindow(value As WTemplate) As UIButtonsPanel
    Set setWindow = Me
    Set pWindow = value
    Set pFrame = pWindow.getBody
End Function

Public Function setHeight(value As Single) As UIButtonsPanel
    Set setHeight = Me
    pHeight = value
    Call updateSize
End Function

Public Function setWidth(value As Single) As UIButtonsPanel
    Set setWidth = Me
    pWidth = value
    Call updateSize
End Function

Public Function setTop(value As Single) As UIButtonsPanel
    Set setTop = Me
    pTop = value
    Call updatePosition
End Function

Public Function setLeft(value As Single) As UIButtonsPanel
    Set setLeft = Me
    pLeft = value
    Call updatePosition
End Function

Public Function setIsVisible(value As Boolean) As UIButtonsPanel
    Set setIsVisible = Me
    pIsVisible = value
    Call updateVisibility
End Function

Public Function setBackColor(value As Long) As UIButtonsPanel
    Set setBackColor = Me
    pBackgroundColor = value
    Call updateBackColor
End Function

Public Function setOkCaption(value As String) As UIButtonsPanel
    Set setOkCaption = Me
    pOkCaption = value
    Call updateButtonsCaptions
End Function

Public Function setCancelCaption(value As String) As UIButtonsPanel
    Set setCancelCaption = Me
    pCancelCaption = value
    Call updateButtonsCaptions
End Function

Public Function setStatus(value As Boolean) As UIButtonsPanel
    Set setStatus = Me
    pStatus = value
    Call updateButtonsVisibility
End Function

Public Function setHasUnsavedData(value As Boolean) As UIButtonsPanel
    Set setHasUnsavedData = Me
    pHasUnsavedData = value
    Call updateUnsavedControlsVisibility
End Function

Public Function setHasCancelButton(value As Boolean) As UIButtonsPanel
    Set setHasCancelButton = Me
    pHasCancelButton = value
    If pIsRendered Then Call updateButtonsVisibility
End Function

Public Function setHasOkButton(value As Boolean) As UIButtonsPanel
    Set setHasOkButton = Me
    pHasOkButton = value
    If pIsRendered Then Call updateButtonsVisibility
End Function

Public Sub addExtraButton(id As String, caption As String, Optional backgroundColor As Long = 13273922, _
                                                            Optional fontColor As Long = vbWhite)
    Dim Button As UIButton
    '------------------------------------------------------------------------------------------------------
    
    Set Button = UI.Create.UIButton(pContainer)
    With Button
        Call .setTag(id)
        Call .setHeight(pButtonHeight)
        Call .setWidth(pButtonWidth)
        Call .setFontSize(pFontSize)
        Call .setFontName(pFontFamily)
        Call .setFontBold(pFontBold)
        Call .setFontColor(fontColor)
        Call .setCaption(caption)
        Call .setBackgroundColor(backgroundColor)
        Call .setBorderColor(backgroundColor)
        Call .setVisibility(True)
        Call .setListener(Me)
        Call .render
    End With
    
    If pExtraButtons Is Nothing Then Set pExtraButtons = New VBA.Collection
    Call pExtraButtons.add(Button)
    Call arrangeButtons
    
End Sub




'[GETTERS]
Public Function getBottom() As Single
    If UI.controls.isValid(pContainer) Then
        With pContainer
            getBottom = .top + .height
        End With
    End If
End Function



'[RENDERING]
Public Function render() As UIButtonsPanel
    Set render = Me
    If Not pIsRendered Then
        Call insertContainer
        Call insertTopBorder
        Call insertOkButton
        Call insertCancelButton
        Call insertUnsavedDataWarnings
        Call arrangeButtons
        pIsRendered = True
    End If
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
    End With
End Sub

Private Sub insertTopBorder()
    If pTopBorder Is Nothing Then
        Set pTopBorder = pContainer.controls.add(CREATE_LABEL_ID, TOP_BORDER_NAME, True)
        With pTopBorder
            .BackStyle = fmBackStyleOpaque
            .backColor = pBorderColor
            .borderStyle = fmBorderStyleNone
            .height = PIXEL_SIZE
            .width = PIXEL_SIZE
        End With
    End If
End Sub

Private Sub insertOkButton()
    If Not pHasOkButton Then
        '[OK] button not needed in this window
    ElseIf pOkButton Is Nothing Then
        Set pOkButton = UI.Create.UIButton(pContainer)
        With pOkButton
            Call .setHeight(pButtonHeight)
            Call .setWidth(pButtonWidth)
            Call .setFontSize(pFontSize)
            Call .setFontName(pFontFamily)
            Call .setFontBold(pFontBold)
            Call .setFontColor(CONFIRM_FONT_COLOR)
            Call .setCaption(MsgService.getText(pOkCaption))
            Call .setBackgroundColor(CONFIRM_BACK_COLOR)
            Call .setBorderColor(CONFIRM_BORDER_COLOR)
            Call .setVisibility(pHasOkButton And pStatus)
            Call .render
        End With
    End If
End Sub

Private Sub insertCancelButton()
    If Not pHasOkButton Then
        '[Cancel] button not needed in this window
    ElseIf pCancelButton Is Nothing Then
        Set pCancelButton = UI.Create.UIButton(pContainer)
        With pCancelButton
            Call .setHeight(pButtonHeight)
            Call .setWidth(pButtonWidth)
            Call .setFontSize(pFontSize)
            Call .setFontName(pFontFamily)
            Call .setFontBold(pFontBold)
            Call .setFontColor(CANCEL_FONT_COLOR)
            Call .setCaption(MsgService.getText(pCancelCaption))
            Call .setBackgroundColor(CANCEL_BACK_COLOR)
            Call .setBorderColor(CANCEL_BORDER_COLOR)
            Call .setVisibility(pHasCancelButton)
            Call .render
        End With
    End If
End Sub

Private Sub insertUnsavedDataWarnings()
    Set pUnsavedDataIcon = pContainer.controls.add(CREATE_LABEL_ID, UNSAVED_DATA_ICON_NAME, False)
    With pUnsavedDataIcon
        .width = UNSAVED_ICON_SIZE
        .height = UNSAVED_ICON_SIZE
        Set .picture = Icons.icoWarning24.picture
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
    End With
    
    Set pUnsavedDataCaption = pContainer.controls.add(CREATE_LABEL_ID, UNSAVED_DATA_CAPTION_NAME, False)
    With pUnsavedDataCaption
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .width = pContainer.InsideWidth
        .caption = MsgService.getText(UNSAVED_DATA_TAG)
        .font.name = UNSAVED_CAPTION_FONT_NAME
        .font.size = UNSAVED_CAPTION_FONT_SIZE
        .font.Bold = UNSAVED_CAPTION_FONT_BOLD
        .ForeColor = UNSAVED_CAPTION_FONT_COLOR
        .AutoSize = True
        .AutoSize = False
    End With
    
End Sub



'[SIZE & POSITION]
Private Sub updatePosition()
    If UI.controls.isValid(pContainer) Then
        With pContainer
            .left = pLeft
            .top = F.Numerics.roundTo(VBA.CDbl(pTop), PIXEL_SIZE)
        End With
    End If
End Sub

Private Sub updateSize()
    If UI.controls.isValid(pContainer) Then
        pContainer.width = pWidth
        Call arrangeButtons
    End If
End Sub

Private Sub updateButtonsVisibility()
    If UI.controls.isValid(pContainer) Then
        If Not pOkButton Is Nothing Then Call pOkButton.setVisibility(pHasOkButton And pStatus)
        If Not pCancelButton Is Nothing Then Call pCancelButton.setVisibility(pHasCancelButton)
        Call arrangeButtons
    End If
End Sub



Private Sub arrangeButtons()
    Dim x As Single
    Dim Button As UIButton
    Dim buttons As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    With pTopBorder
        .top = 0
        .left = 0
        .width = pContainer.InsideWidth
    End With
    
    With pUnsavedDataIcon
        .left = pPadding
        .top = (pContainer.height - .height) / 2
    End With
    
    With pUnsavedDataCaption
        .left = pPadding + UNSAVED_ICON_SIZE + DEFAULT_BUTTONS_SPACE
        .top = (pContainer.height - .height) / 2
    End With
    
    
    x = pContainer.InsideWidth - pPadding
    Set buttons = F.Collections.Create(pCancelButton, pOkButton)
    For Each Button In buttons
        If Button Is Nothing Then
            '
        ElseIf Not Button.isVisible Then
            '
        Else
            With Button
                Call .Move(x - .getWidth, (pContainer.height - .getHeight) / 2)
                x = .getLeft - pButtonsSpace
            End With
        End If
    Next Button
    
    '[Extra buttons]
    x = pPadding
    If F.Collections.isNonEmptyCollection(pExtraButtons) Then
        For Each Button In pExtraButtons
            With Button
                Call .Move(x, (pContainer.height - .getHeight) / 2)
                x = .getRight + pButtonsSpace
            End With
        Next Button
    End If
    
End Sub



Private Sub updateBackColor()
    If UI.controls.isValid(pContainer) Then
        pContainer.backColor = pBackgroundColor
    End If
End Sub

Private Sub updateButtonsCaptions()
    If Not pOkButton Is Nothing Then
        Call pOkButton.setCaption(MsgService.getText(pOkCaption))
    End If
    
    If Not pCancelButton Is Nothing Then
        Call pCancelButton.setCaption(MsgService.getText(pCancelCaption))
    End If
    
End Sub

Private Sub updateVisibility()
    pContainer.visible = pIsVisible
End Sub

Private Sub updateUnsavedControlsVisibility()
    If UI.controls.isValid(pUnsavedDataIcon) Then
        pUnsavedDataIcon.visible = pHasUnsavedData
    End If
    If UI.controls.isValid(pUnsavedDataCaption) Then
        pUnsavedDataCaption.visible = pHasUnsavedData
    End If
End Sub

Private Function confirmSkippingNotSavedData() As Boolean
    If Not pWindow.displayWarningWhenClosingWithUnsavedData Then
        confirmSkippingNotSavedData = True
    ElseIf pHasUnsavedData Then
        confirmSkippingNotSavedData = F.Dialogs.getUserConfirmation(MsgService.getText(UNSAVED_DATA_QUESTION))
    Else
        confirmSkippingNotSavedData = True
    End If
End Function





'[EVENTS]
Private Sub pCancelButton_Click()
    Call ErrorManager.clear
    
    If confirmSkippingNotSavedData Then
        RaiseEvent Cancel
    End If
    
End Sub

Private Sub pOkButton_Click()
    Call ErrorManager.clear
    RaiseEvent Confirm
End Sub




'[Implementation of IListener interface]
Private Sub IListener_catchEvent(src As Variant, ev As Variant, Optional properties As Scripting.IDictionary)
    Dim btn As UIButton
    '------------------------------------------------------------------------------------------------------
    Set btn = src
    Call pWindow.triggerCustomAction(btn.getTag)
End Sub

