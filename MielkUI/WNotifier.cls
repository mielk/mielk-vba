VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum ControlTypeEnum
    ControlType_CaptionLabel = 1
    ControlType_ProgressBarFrame = 2
    ControlType_ProgressBarProgress = 3
    ControlType_ProgressBarCaption = 4
    ControlType_HiddenLabel = 5
End Enum

Private Const CLASS_NAME As String = "WNotifier"
'[Default UI properties] ----------------------------------------------------------------------------------
Private Const DEFAULT_MODE As Long = VBA.vbModeless
Private Const DEFAULT_WIDTH As Single = 528
Private Const DEFAULT_BACKCOLOR As Long = &H8000000F
'[Font]
Private Const DEFAULT_FONT_FAMILY As String = "Century Gothic"
Private Const DEFAULT_FONT_SIZE As Single = 11
Private Const DEFAULT_FONT_COLOR As Long = vbBlack
Private Const DEFAULT_TEXT_ALIGNMENT As Long = MSForms.fmTextAlignCenter
'[Layout]
Private Const DEFAULT_HORIZONTAL_PADDING As Single = 12
'[UI components] ------------------------------------------------------------------------------------------
Private WithEvents pForm As WTemplate
Attribute pForm.VB_VarHelpID = -1
Private pFrame As MSForms.frame
Private pCaptionLabel As MSForms.label
'Private pHiddenLabel As MSForms.label
'[Buttons] ------------------------------------------------------------------------------------------------
Private pHasOkButton As Boolean
Private pOkButtonCaption As String
Private pHasCancelButton As Boolean
Private pCancelButtonCaption As String
'[UI properties] ------------------------------------------------------------------------------------------
Private pMode As VBA.FormShowConstants
Private pTitle As String
Private pCaption As String
Private pWidth As Single
Private pAutoWidth As Single
Private pBackColor As Long
Private pFontColor As Long
Private pFontFamily As String
Private pFontSize As Single
Private pTextAlignment As MSForms.fmTextAlign
Private pPadding As Single
Private pPercentFormat As String
'[State] --------------------------------------------------------------------------------------------------
Private pUuid As String
Private pIsRendered As Boolean
Private pIsCancelledByUser As Boolean
Private pIsConfirmedByUser As Boolean
'[Events] -------------------------------------------------------------------------------------------------
Event Confirm()
Event Cancel()
'----------------------------------------------------------------------------------------------------------




'[INITIALIZE]
Private Sub Class_Initialize()
    Call initializeDefaultValues
End Sub

Private Sub initializeDefaultValues()
    pUuid = F.Crypto.createUUID
    pMode = DEFAULT_MODE
    pWidth = DEFAULT_WIDTH
    pBackColor = DEFAULT_BACKCOLOR
    pFontFamily = DEFAULT_FONT_FAMILY
    pFontSize = DEFAULT_FONT_SIZE
    pFontColor = DEFAULT_FONT_COLOR
    pTextAlignment = DEFAULT_TEXT_ALIGNMENT
    pPadding = DEFAULT_HORIZONTAL_PADDING
End Sub




'[SETTERS]
Public Function setCaption(value As String) As WNotifier
    Set setCaption = Me
    pCaption = value
    If pIsRendered Then Call updateLayout
End Function

Public Function setTitle(value As String) As WNotifier
    Set setTitle = Me
    pTitle = value
End Function

Public Function setAutoWidth(value As Boolean) As WNotifier
    Set setAutoWidth = Me
    pAutoWidth = value
End Function


'[Buttons]
Public Function setHasOkButton(value As Boolean) As WNotifier
    Set setHasOkButton = Me
    pHasOkButton = value
    If pIsRendered Then Call updateLayout
End Function

Public Function setOkButtonCaption(value As String) As WNotifier
    Set setOkButtonCaption = Me
    pHasOkButton = True
    pOkButtonCaption = value
    If pIsRendered Then Call updateLayout
End Function

Public Function setHasCancelButton(value As Boolean) As WNotifier
    Set setHasCancelButton = Me
    pHasCancelButton = value
    If pIsRendered Then Call updateLayout
End Function

Public Function setCancelButtonCaption(value As String) As WNotifier
    Set setCancelButtonCaption = Me
    pHasCancelButton = True
    pCancelButtonCaption = value
    If pIsRendered Then Call updateLayout
End Function



'[Size]
Public Function setWidth(value As Single) As WNotifier
    Set setWidth = Me
    pWidth = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setPadding(value As Single) As WNotifier
    Set setPadding = Me
    pPadding = value
    If pIsRendered Then Call arrangeControls
End Function




'[Colors]
Public Function setBackColor(value As Long) As WNotifier
    Set setBackColor = Me
    pBackColor = value
    If pIsRendered Then Call updateLayout
End Function

'[Font]
Public Function setFontFamily(value As String) As WNotifier
    Set setFontFamily = Me
    pFontFamily = value
    If pIsRendered Then Call updateLayout
End Function

Public Function setFontSize(value As Single) As WNotifier
    Set setFontSize = Me
    pFontSize = value
    If pIsRendered Then Call updateLayout
End Function

Public Function setFontColor(value As Long) As WNotifier
    Set setFontColor = Me
    pFontColor = value
    If pIsRendered Then Call updateLayout
End Function

Public Function setTextAlignment(value As MSForms.fmTextAlign) As WNotifier
    Set setTextAlignment = Me
    pTextAlignment = value
    If pIsRendered Then Call updateLayout
End Function

Public Function setHorizontalPadding(value As Single) As WNotifier
    Set setHorizontalPadding = Me
    pPadding = value
    If pIsRendered Then Call arrangeControls
End Function



'[GETTERS]
Public Function isCancelledByUser() As Boolean
    isCancelledByUser = pIsCancelledByUser
End Function

Public Function isConfirmedByUser() As Boolean
    isConfirmedByUser = pIsConfirmedByUser
End Function





'[DISPLAYING]
Public Sub show(Optional caption As String, Optional mode As VBA.FormShowConstants = vbModeless)
    If VBA.Len(caption) Then pCaption = caption
    
    If Not pIsRendered Then
        Call createForm
        Call render
        Stop        '@@@tm24
        'Call pForm.setShowOnTaskBar(True)
        Call pForm.arrangeComponents
        Call pForm.locateRelativeTo(UI.Create.LocationPropsForScreenCenter)
    End If
    
    If pMode <> mode Then Call hide
    pMode = mode
    Call pForm.display(pMode, False)
    
    DoEvents
    
End Sub

Public Function hide()
    If Not pForm Is Nothing Then
        If pForm.isVisible Then
            Call pForm.hide
        End If
    End If
End Function

Private Sub createForm()
    If pForm Is Nothing Then
        Set pForm = UI.Create.WindowTemplate
        With pForm
            Call .setHasTitleBar(False)
            Call .setHasButtonsPanel(pHasOkButton Or pHasCancelButton)
            Call .setHasOkButton(pHasOkButton)
            Call .setOkCaption(pOkButtonCaption)
            Call .setHasCancelButton(pHasCancelButton)
            Call .setCancelCaption(pCancelButtonCaption)
            Call .setContentFrameHorizontalMargin(pPadding)
            Call .setContentFrameVerticalMargin(pPadding)
            Call .setCaption(pTitle)
            Call .render
        End With
    End If
End Sub


Private Sub render()
    Set pFrame = pForm.getContentFrame
    Call insertCaptionLabel
    Call arrangeControls
    pIsRendered = True
End Sub


Private Sub insertCaptionLabel()
    If pCaptionLabel Is Nothing Then
        Set pCaptionLabel = pFrame.controls.add(CREATE_LABEL_ID)
    End If
    With pCaptionLabel
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        With .font
            .size = pFontSize
            .name = APP_FONT_FAMILY
            .Bold = False
        End With
        .caption = pCaption
    End With
End Sub

'Private Sub insertCaptionLabel()
    '@@@tm24
'    Set pHiddenLabel = pForm.addControl(UIControlType_Label, getLabelName(ControlType_HiddenLabel), True)
'    With pHiddenLabel
'        .BackStyle = fmBackStyleTransparent
'        .BorderStyle = fmBorderStyleNone
'        .ForeColor = pForm.getBackColor
'    End With
'    Set pCaptionLabel = pForm.addControl(UIControlType_Label, getLabelName(ControlType_CaptionLabel), True)
'    With pCaptionLabel
'        .BackStyle = fmBackStyleTransparent
'        .BorderStyle = fmBorderStyleNone
'    End With
'End Sub




'[Size & location]

Private Sub arrangeControls()
    With pCaptionLabel
        Stop
'        .width = pBarWidth
'        .AutoSize = True
'        .AutoSize = False
'        .width = pBarWidth
'        .left = pPadding
'        .top = pPadding
    End With
    
    Call pForm.arrangeComponents
    
End Sub

'
'Public Sub arrangeControls()
'    Dim Y As Single
'    Dim width As Single
'    '----------------------------------------------------------------------------------------------------------
'
'    Call resizeCaptionLabel
'    Call UI.Controls.locate(pHiddenLabel, pPadding, pPadding)
'    Call UI.Controls.locate(pCaptionLabel, pPadding, pPadding)
'    Y = UI.Controls.getBottom(pCaptionLabel) + pPadding
'    width = pCaptionLabel.width + 2 * pPadding
'
'    '@@@tm24
''    If isButtonsPanelVisible Then
''        Call pButtonsPanel.resize(width:=width + 2)
''        Call pButtonsPanel.Move(left:=-1, top:=Y)
''    End If
'
'    With pForm
'    '@@@tm24
''        Call .setCaption(pTitle)
''        Call .setWidth(width)
''        Call .setHeight(calculateUserFormHeight)
'    End With
'
'End Sub

'
'Private Function calculateUserFormHeight() As Single
'    '@@@tm24
''    If pButtonsPanel.isVisible Then
''        calculateUserFormHeight = pButtonsPanel.getBottom - PIXEL_SIZE
''    ElseIf pProgressBarVisible Then
''        calculateUserFormHeight = UI.Controls.getBottom(pProgressBarFrameLabel) + pPadding
''    Else
''        calculateUserFormHeight = UI.Controls.getBottom(pCaptionLabel) + pPadding
''    End If
'End Function
'
'Private Sub resizeCaptionLabel()
'    With pHiddenLabel
'        .width = pWidth - 2 * pPadding
'        .AutoSize = True
'        DoEvents
'        If pAutoWidth Then
'            .width = .width + 2 * pPadding
'        Else
'            .width = pWidth - 2 * pPadding
'        End If
'        .AutoSize = False
'    End With
'
'    With pCaptionLabel
'        .width = pHiddenLabel.width
'        .height = pHiddenLabel.height
'    End With
'
'End Sub


Private Sub updateLayout()
'    With pHiddenLabel
'        .caption = pCaption
'        .font.name = pFontFamily
'        .font.size = pFontSize
'        DoEvents
'        .TextAlign = pTextAlignment
'    End With
    With pCaptionLabel
        .caption = pCaption
        .font.name = pFontFamily
        .font.size = pFontSize
        DoEvents
        .ForeColor = pFontColor
        .TextAlign = pTextAlignment
    End With
    Call arrangeControls
End Sub


'[User form events]

Private Sub pForm_Cancel()
    pIsCancelledByUser = True
    pIsConfirmedByUser = False
    Stop
    'Call hide
    RaiseEvent Cancel
End Sub

Private Sub pForm_Confirm()
    pIsCancelledByUser = False
    pIsConfirmedByUser = True
    Stop
    'Call hide
    RaiseEvent Confirm
End Sub
