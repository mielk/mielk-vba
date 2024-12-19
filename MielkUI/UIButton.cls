VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IControl

Private Const CLASS_NAME As String = "UIButton"
'----------------------------------------------------------------------------------------------------------
Private Const CAPTION_MARGIN As Single = 3
'[Default UI properties] ----------------------------------------------------------------------------------
Private Const DEFAULT_WIDTH As Single = 100
Private Const DEFAULT_HEIGHT As Single = 32
Private Const DEFAULT_VISIBLE As Boolean = True
Private Const DEFAULT_BACK_COLOR As Long = vbWhite
Private Const DEFAULT_BORDER_VISIBLE As Boolean = True
Private Const DEFAULT_BORDER_COLOR As Long = vbBlack
Private Const DEFAULT_FONT_NAME As String = APP_FONT_FAMILY
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_FONT_BOLD As Boolean = True
Private Const DEFAULT_FONT_COLOR As Long = vbBlack
Private Const DEFAULT_DISABLED_FONT_COLOR As Long = 11842740
Private Const DEFAULT_PICTURE_MARGIN As Single = 3
Private Const DEFAULT_PICTURE_SIZE As Single = 18
Private Const DEFAULT_PADDING As Single = 9
'----------------------------------------------------------------------------------------------------------
Private pFrame As MSForms.frame
Private pListener As IListener
Private pBackgroundLabel As MSForms.label
Private pPictureLabel As MSForms.label
Private pCaptionLabel As MSForms.label
Private WithEvents pEventsLayerLabel As MSForms.label
Attribute pEventsLayerLabel.VB_VarHelpID = -1
'----------------------------------------------------------------------------------------------------------
Private pUuid As String
Private pTag As String
Private pIsRendered As Boolean
Private pIsDisabled As Boolean
'----------------------------------------------------------------------------------------------------------
Private pWidth As Single
Private pHeight As Single
Private pLeft As Single
Private pTop As Single
Private pVisible As Boolean
Private pEventName As String
Private pTooltip As String
Private pCaption As String
Private pTextAlignment As fmTextAlign
'[Background and borders]
Private pBackColor As Long
Private pBorderColor As Long
Private pBorderVisible As Boolean
'[Font]
Private pFontName As String
Private pFontSize As Single
Private pFontBold As Boolean
Private pFontColor As Long
Private pDisabledFontColor As Long
'[Picture]
Private pPicture As Variant
Private pPictureMargin As Single
Private pPictureSize As Single

'[Events] -------------------------------------------------------------------------------------------------
Event AfterClick(props As Scripting.Dictionary, ByRef Cancel As Boolean)
Event Click()
Event AfterHover()
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Initialize()
    pUuid = F.Crypto.createUUID
    Call setDefaultProperties
End Sub

Private Sub setDefaultProperties()
    pWidth = DEFAULT_WIDTH
    pHeight = DEFAULT_HEIGHT
    pVisible = DEFAULT_VISIBLE
    pBackColor = DEFAULT_BACK_COLOR
    pBorderVisible = DEFAULT_BORDER_VISIBLE
    pBorderColor = DEFAULT_BORDER_COLOR
    pFontName = DEFAULT_FONT_NAME
    pFontSize = DEFAULT_FONT_SIZE
    pFontBold = DEFAULT_FONT_BOLD
    pFontColor = DEFAULT_FONT_COLOR
    pDisabledFontColor = DEFAULT_DISABLED_FONT_COLOR
    pPictureMargin = DEFAULT_PICTURE_MARGIN
    pPictureSize = DEFAULT_PICTURE_SIZE
End Sub





'[SETTERS]
Public Function setFrame(value As MSForms.frame) As UIButton
    Set setFrame = Me
    Set pFrame = value
End Function

Public Function setListener(value As IListener) As UIButton
    Set setListener = Me
    Set pListener = value
End Function

Public Function setTag(value As String) As UIButton
    Set setTag = Me
    pTag = value
End Function

Public Function setCaption(value As String) As UIButton
    Set setCaption = Me
    pCaption = value
    Call updateCaptionSizeAndContent
End Function

Public Function setHeight(value As Single) As UIButton
    Set setHeight = Me
    pHeight = value
    Call locateMe
    Call arrangeLayout
End Function

Public Function setWidth(value As Single) As UIButton
    Set setWidth = Me
    pWidth = value
    Call locateMe
    Call arrangeLayout
End Function

Public Function setLeft(value As Single) As UIButton
    Set setLeft = Me
    pLeft = value
    Call locateMe
End Function

Public Function setRight(value As Single) As UIButton
    Set setRight = Me
    pLeft = pFrame.InsideWidth - value - getWidth
    Call locateMe
    Call arrangeLayout
End Function

Public Function setTop(value As Single) As UIButton
    Set setTop = Me
    pTop = value
    Call locateMe
End Function

Public Function setTooltip(value As String) As UIButton
    Set setTooltip = Me
    pTooltip = value
    Call updateToolTip
End Function

Public Function setVisibility(value As Boolean) As UIButton
    Set setVisibility = Me
    pVisible = value
    Call updateVisibility
End Function

Public Function setFontName(value As String) As UIButton
    Set setFontName = Me
    pFontName = value
    Call updateCaptionSizeAndContent
End Function

Public Function setFontSize(value As Single) As UIButton
    Set setFontSize = Me
    pFontSize = value
    Call updateCaptionSizeAndContent
End Function

Public Function setFontBold(value As Boolean) As UIButton
    Set setFontBold = Me
    pFontBold = value
    Call updateCaptionSizeAndContent
End Function

Public Function setFontColor(value As Long) As UIButton
    Set setFontColor = Me
    pFontColor = value
    Call updateCaptionView
End Function

Public Function setDisabledFontColor(value As Long) As UIButton
    Set setDisabledFontColor = Me
    pDisabledFontColor = value
    Call updateCaptionView
End Function

Public Function setTextAlignment(value As fmTextAlign) As UIButton
    Set setTextAlignment = Me
    pTextAlignment = value
    Call updateCaptionSizeAndContent
End Function

Public Function setPicture(value As Variant) As UIButton
    Set setPicture = Me
    Set pPicture = value
    Call updatePictureContent
End Function

Public Function setBackgroundColor(value As Long) As UIButton
    Set setBackgroundColor = Me
    pBackColor = value
    Call updateBackground
End Function

Public Function setBorderColor(value As Long) As UIButton
    Set setBorderColor = Me
    pBorderColor = value
    Call updateBorder
End Function

Public Function setBorderVisibility(value As Boolean) As UIButton
    Set setBorderVisibility = Me
    pBorderVisible = value
    Call updateBorder
End Function

Public Function setIsEnabled(value As Boolean) As UIButton
    Set setIsEnabled = Me
    pIsDisabled = Not value
    Call updateCaptionView
    Call updateMousePointer
End Function



'[GETTERS]
Public Function getTag() As String
    getTag = pTag
End Function

Public Function getCaption() As String
    getCaption = pCaption
End Function

Public Function getRight() As Single
    If Not pBackgroundLabel Is Nothing Then
        getRight = pBackgroundLabel.left + pBackgroundLabel.width
    End If
End Function

Public Function getBottom() As Single
    If Not pBackgroundLabel Is Nothing Then
        getBottom = pBackgroundLabel.top + pBackgroundLabel.height
    End If
End Function

Public Function getLeft() As Single
    If Not pBackgroundLabel Is Nothing Then
        getLeft = pBackgroundLabel.left
    End If
End Function

Public Function getTop() As Single
    If Not pBackgroundLabel Is Nothing Then
        getTop = pBackgroundLabel.top
    End If
End Function

Public Function getHeight() As Single
    If Not pBackgroundLabel Is Nothing Then
        getHeight = pBackgroundLabel.height
    End If
End Function

Public Function getWidth() As Single
    If Not pBackgroundLabel Is Nothing Then
        getWidth = pBackgroundLabel.width
    End If
End Function

Public Function getTextAlignment() As fmTextAlign
    getTextAlignment = pTextAlignment
End Function

Public Function isVisible() As Boolean
    isVisible = pVisible
End Function













'[RENDERING]
Public Function render() As UIButton
    Set render = Me
    If Not pIsRendered Then
        Call insertBackgroundLabel
        Call insertCaptionLabel
        Call insertPictureLabel
        Call insertEventsLayer
        pIsRendered = True
        Call locateMe
    End If
    
    Call updateInsideProperties
    
End Function

Private Sub insertBackgroundLabel()
    If pBackgroundLabel Is Nothing Then
        Set pBackgroundLabel = pFrame.controls.add(CREATE_LABEL_ID)
    End If
End Sub

Private Sub insertCaptionLabel()
    If pCaptionLabel Is Nothing Then
        Set pCaptionLabel = pFrame.controls.add(CREATE_LABEL_ID)
    End If
    With pCaptionLabel
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .WordWrap = False
    End With
End Sub

Private Sub insertPictureLabel()
    If pPictureLabel Is Nothing Then
        Set pPictureLabel = pFrame.controls.add(CREATE_LABEL_ID)
        With pPictureLabel
            .BackStyle = fmBackStyleTransparent
            .borderStyle = fmBorderStyleNone
        End With
    End If
End Sub

Private Sub insertEventsLayer()
    If pEventsLayerLabel Is Nothing Then
        Set pEventsLayerLabel = pFrame.controls.add(CREATE_LABEL_ID)
    End If
    With pEventsLayerLabel
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .MousePointer = fmMousePointerCustom
        .MouseIcon = Icons.getMousePointer
    End With
End Sub

Public Function Move(Optional x As Variant, Optional y As Variant) As UIButton
    Dim offsetX As Single
    Dim offsetY As Single
    '------------------------------------------------------------------------------------------------------
    
    Set Move = Me
    
    With pBackgroundLabel
        If VBA.IsNumeric(x) Then
            offsetX = (x - .left)
        End If
        If VBA.IsNumeric(y) Then
            offsetY = (y - .top)
        End If
    End With
    
    Call offset(offsetX, offsetY)
    
End Function

Public Function offset(Optional x As Single, Optional y As Single) As UIButton
    Set offset = Me
    Call UI.Forms.offsetControls(x, y, pBackgroundLabel, pCaptionLabel, pPictureLabel, pEventsLayerLabel)
    With pBackgroundLabel
        pTop = .top
        pLeft = .left
    End With
    Call UI.controls.copySizeAndPosition(pEventsLayerLabel, pBackgroundLabel)
End Function

Public Function alignVertically(ctrl As MSForms.control) As UIButton
    Dim yOffset As Single
    '------------------------------------------------------------------------------------------------------
    
    Set alignVertically = Me
    
    If UI.controls.isValid(ctrl) Then
        yOffset = (ctrl.top + (ctrl.height - pHeight) / 2) - pTop
        If yOffset Then Call offset(y:=yOffset)
    End If
    
End Function


Private Sub locateMe()
    If pIsRendered Then
        With pBackgroundLabel
            .width = UI.Normalize.size(F.Math.getMax(pWidth, pCaptionLabel.width + 2 * DEFAULT_PADDING))
            .height = UI.Normalize.size(pHeight)
            .top = UI.Normalize.position(pTop)
            .left = UI.Normalize.position(pLeft)
        End With
    End If
End Sub

Private Sub arrangeLayout()
    Dim totalWidth As Single
    Dim captionTop As Single
    Dim captionLeft As Single
    Dim pictureVisible As Boolean
    Dim pictureLeft As Single
    Dim pictureTop As Single
    Dim pictureMargin As Single
    '------------------------------------------------------------------------------------------------------
    
    If pIsRendered Then
        If pPictureLabel.visible Then
            pictureVisible = True
            pictureMargin = pPictureMargin
            totalWidth = pCaptionLabel.width + pPictureLabel.width + pictureMargin
        Else
            totalWidth = pCaptionLabel.width
        End If
        
        With pBackgroundLabel
            If pictureVisible Then
                pictureTop = UI.Normalize.position(.top + (.height - pCaptionLabel.height) / 2)
                pictureLeft = UI.Normalize.position(.left + (.width - totalWidth) / 2)
                captionLeft = UI.Normalize.position(pictureLeft + pPictureLabel.width + pictureMargin)
            Else
                captionLeft = UI.Normalize.position(.left + (.width - totalWidth) / 2)
            End If
            captionTop = UI.Normalize.position(.top + (.height - pCaptionLabel.height) / 2)
        End With
        
        If pictureVisible Then
            With pPictureLabel
                .top = pictureTop
                .left = pictureLeft
            End With
        End If
        
        With pCaptionLabel
            .top = captionTop
            .left = captionLeft
        End With
        
        '[Events layer]
        Call UI.controls.copySizeAndPosition(pEventsLayerLabel, pBackgroundLabel)
        
    End If
    
End Sub


Private Sub updateInsideProperties()
    Call updateBackground
    Call updateBorder
    Call updateMousePointer
    Call updateCaptionView
    Call updateCaptionSizeAndContent(False)
    Call updatePictureSizeAndLayout(False)
    Call updatePictureContent(False)
    Call arrangeLayout
    Call updatePosition
End Sub


Private Sub updateBackground()
    If pIsRendered Then
        With pBackgroundLabel
            .BackStyle = fmBackStyleOpaque
            .backColor = pBackColor
        End With
    End If
End Sub

Private Sub updateBorder()
    Dim borderStyle As fmBorderStyle
    Dim borderColor As Long
    '------------------------------------------------------------------------------------------------------
    
    If pIsRendered Then
        If pBorderVisible Then
            borderStyle = fmBorderStyleSingle
            borderColor = pBorderColor
        Else
            borderStyle = fmBorderStyleNone
        End If
        
        With pBackgroundLabel
            .borderStyle = borderStyle
            .borderColor = borderColor
        End With
        
    End If
    
End Sub

Private Sub updatePictureContent(Optional updateLayout As Boolean = True)
    With pPictureLabel
        If Not F.Variables.isUndefined(pPicture) Then
            .visible = True
            Set .picture = pPicture
        Else
            .visible = False
        End If
    End With
    
    If updateLayout Then
        Call arrangeLayout
    End If
End Sub

Private Sub updatePictureSizeAndLayout(Optional updateLayout As Boolean = True)
    Dim hasPicture As Boolean
    '------------------------------------------------------------------------------------------------------
    
    hasPicture = VBA.IsObject(pPicture)
    
    If hasPicture Then
        With pPicture.label
            .visible = hasPicture
            .width = pPictureSize
            .height = pPictureSize
        End With
        
        If updateLayout Then
            Call arrangeLayout
        End If
    End If
End Sub

Private Sub updatePosition()
    Dim left As Single
    Dim top As Single
    Dim offsetX As Single
    Dim offsetY As Single
    '------------------------------------------------------------------------------------------------------
    
    If pIsRendered Then
        With pBackgroundLabel
            left = .left
            top = .top
        End With
        
        offsetX = pLeft - left
        offsetY = pTop - top
        
        If offsetX <> 0 Or offsetY <> 0 Then
            Call UI.Forms.offsetControls(offsetX, offsetY, pBackgroundLabel, pCaptionLabel, pEventsLayerLabel)
        End If
        
    End If
    
End Sub

Private Sub updateMousePointer()
    With pEventsLayerLabel
        If pIsDisabled Then
            .MousePointer = fmMousePointerDefault
        Else
            .MousePointer = fmMousePointerCustom
            Set .MouseIcon = Icons.getMousePointer
        End If
    End With
End Sub

Private Sub updateCaptionView()
    If pIsRendered Then
        With pCaptionLabel
            .ForeColor = VBA.IIf(pIsDisabled, pDisabledFontColor, pFontColor)
        End With
    End If
End Sub

Private Sub updateCaptionSizeAndContent(Optional updateLayout As Boolean = True)
    If pIsRendered Then
        With pCaptionLabel
            .font.name = pFontName
            .font.Bold = VBA.IIf(pFontBold, 1, False)
            .font.size = pFontSize
            .TextAlign = fmTextAlignCenter
            .caption = pCaption
            .WordWrap = False
            .AutoSize = False
            .AutoSize = True
        End With
        
        If updateLayout Then
            Call arrangeLayout
        End If
        
    End If
End Sub


Private Sub updateToolTip()
    pEventsLayerLabel.ControlTipText = pTooltip
End Sub

Public Sub hideOverflow()
    Dim arrContent() As String
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    With pCaptionLabel
        .AutoSize = True
        If .width > pBackgroundLabel.width Then
            arrContent = VBA.Split(pCaption, VBA.vbCrLf)
            For i = LBound(arrContent) To UBound(arrContent)
                arrContent(i) = getWrappedCaptionForLine(arrContent(i))
            Next i
            .caption = VBA.Join(arrContent, VBA.vbCrLf)
        End If
    End With
    
End Sub

Private Function getWrappedCaptionForLine(caption As String) As String
    Const ADD_TEXT As String = "..."
    Const EXTRA_MARGIN As Single = 1.1
    '------------------------------------------------------------------------------------------------------
    Dim text As String
    Dim ratio As Double
    Dim length As Long
    '------------------------------------------------------------------------------------------------------
    
    With pCaptionLabel
        .caption = caption
        .AutoSize = True
        
        If .width > pBackgroundLabel.width Then
            ratio = pBackgroundLabel.width / .width
            text = VBA.left(caption, VBA.Len(caption) * ratio * EXTRA_MARGIN)
            length = VBA.Len(text)
            
            Do
                .AutoSize = True
                If .width >= pBackgroundLabel.width Then
                    .caption = VBA.left$(text, length) & ADD_TEXT
                    length = length - 1
                ElseIf .width < pBackgroundLabel.width Then
                    Exit Do
                ElseIf VBA.Len(.caption) = 0 Then
                    Exit Do
                End If
            Loop
            
        End If
        
        getWrappedCaptionForLine = .caption
        
    End With
    
End Function




Public Sub delete()
    On Error Resume Next
    With pFrame.controls
        Call .remove(pBackgroundLabel.name)
        Call .remove(pCaptionLabel.name)
        Call .remove(pPictureLabel.name)
        Call .remove(pEventsLayerLabel.name)
    End With
End Sub

Public Sub updateVisibility()
    If pIsRendered Then
        On Error Resume Next
        pBackgroundLabel.visible = pVisible
        pCaptionLabel.visible = pVisible
        pEventsLayerLabel.visible = pVisible
        pPictureLabel.visible = pVisible
    End If
End Sub




'[Events]
Private Sub pEventsLayerLabel_Click()
    Dim props As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Call ErrorManager.clear
    
    If Not pIsDisabled Then
        RaiseEvent Click
        If Not pListener Is Nothing Then
            Call pListener.catchEvent(Me, UI.Events.Click.getName)
        End If
    End If
    
End Sub



'[Implementation of IControl interface]
Private Sub IControl_setVisibility(value As Boolean)
    Call setVisibility(value)
End Sub

Private Sub IControl_setEnability(value As Boolean)
    Stop
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

Private Function IControl_getRectangle() As MielkCore.RECT
    Stop
End Function

Private Sub IControl_alignVertically(baseCtrl As MSForms.IControl)
    Call alignVertically(baseCtrl)
End Sub

Private Function IControl_move(Optional x As Variant, Optional y As Variant) As IControl
    Set IControl_move = Move(x, y)
End Function

Private Function IControl_offset(Optional x As Single, Optional y As Single) As IControl
    Set IControl_offset = offset(x, y)
End Function

Private Sub IControl_resize(Optional width As Variant, Optional height As Variant)
    Stop
End Sub

Private Sub IControl_triggerEvent(eventId As Long)
    Stop
End Sub
