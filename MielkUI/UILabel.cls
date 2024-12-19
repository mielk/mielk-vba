VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UILabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IControl

Private Const CLASS_NAME As String = "UILabel"
'----------------------------------------------------------------------------------------------------------
Private Const CAPTION_MARGIN As Single = 3
'----------------------------------------------------------------------------------------------------------
Private pFrame As MSForms.frame
Private pListener As IListener
Private WithEvents pBackgroundLabel As MSForms.label
Attribute pBackgroundLabel.VB_VarHelpID = -1
Private WithEvents pCaptionLabel As MSForms.label
Attribute pCaptionLabel.VB_VarHelpID = -1
'----------------------------------------------------------------------------------------------------------
Private pUuid As String
Private pTag As String
Private pIsRendered As Boolean
'----------------------------------------------------------------------------------------------------------
Private pWidth As Single
Private pHeight As Single
Private pLeft As Single
Private pTop As Single
Private pVisible As Boolean
Private pTooltip As String
Private pCaption As String
Private pNumberFormat As String
Private pTextAlignment As fmTextAlign
Private pIsUnderlined As Boolean
Private pMousePointer As Object
'[Font]
Private pFontName As String
Private pFontSize As Single
Private pFontBold As Boolean
Private pFontColor As Long
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Initialize()
    pUuid = F.Crypto.createUUID
    pTextAlignment = fmTextAlignLeft
    pNumberFormat = "@"
    pVisible = True
End Sub







'[SETTERS]
Public Function setFrame(value As MSForms.frame) As UILabel
    Set setFrame = Me
    Set pFrame = value
End Function

Public Function setListener(value As IListener) As UILabel
    Set setListener = Me
    Set pListener = value
End Function

Public Function setTag(value As String) As UILabel
    Set setTag = Me
    pTag = value
End Function

Public Function setCaption(value As String) As UILabel
    Set setCaption = Me
    pCaption = value
    Call updateCaption
End Function

Public Function setHeight(value As Single) As UILabel
    Set setHeight = Me
    pHeight = value
    Call updateSize
End Function

Public Function setWidth(value As Single) As UILabel
    Set setWidth = Me
    pWidth = value
    Call updateSize
End Function

Public Function setTooltip(value As String) As UILabel
    Set setTooltip = Me
    pTooltip = value
    Call updateToolTips
End Function

Public Function setVisibility(value As Boolean) As UILabel
    Set setVisibility = Me
    pVisible = value
    Call updateVisibility
End Function

Public Function setFontName(value As String) As UILabel
    Set setFontName = Me
    pFontName = value
    Call updateCaption
End Function

Public Function setFontSize(value As Single) As UILabel
    Set setFontSize = Me
    pFontSize = value
    Call updateCaption
End Function

Public Function setFontBold(value As Boolean) As UILabel
    Set setFontBold = Me
    pFontBold = value
    Call updateCaption
End Function

Public Function setFontColor(value As Long) As UILabel
    Set setFontColor = Me
    pFontColor = value
    Call updateCaption
End Function

Public Function setTextAlignment(value As fmTextAlign) As UILabel
    Set setTextAlignment = Me
    pTextAlignment = value
    Call updateTextAlignment
End Function

Public Function setNumberFormat(value As String) As UILabel
    Set setNumberFormat = Me
    pNumberFormat = value
    Call updateCaption
End Function

Public Function setIsUnderlined(value As Boolean) As UILabel
    Set setIsUnderlined = Me
    pIsUnderlined = value
End Function

Public Function setMousePointer(value As Object) As UILabel
    Set setMousePointer = Me
    Set pMousePointer = value
End Function




'[GETTERS]

Public Function getTag() As String
    getTag = pTag
End Function

Public Function getCaption() As String
    getCaption = pCaption
End Function

Public Function getLeft() As Single
    If UI.controls.isValid(pBackgroundLabel) Then
        getLeft = pBackgroundLabel.left
    End If
End Function

Public Function getRight() As Single
    If UI.controls.isValid(pBackgroundLabel) Then
        getRight = pBackgroundLabel.left + pBackgroundLabel.width
    End If
End Function

Public Function getTop() As Single
    If UI.controls.isValid(pBackgroundLabel) Then
        getTop = pBackgroundLabel.top
    End If
End Function

Public Function getBottom() As Single
    If UI.controls.isValid(pBackgroundLabel) Then
        getBottom = pBackgroundLabel.top + pBackgroundLabel.height
    End If
End Function

Public Function getHeight() As Single
    If UI.controls.isValid(pBackgroundLabel) Then
        getHeight = pBackgroundLabel.height
    End If
End Function

Public Function getWidth() As Single
    If UI.controls.isValid(pBackgroundLabel) Then
        getWidth = pBackgroundLabel.width
    End If
End Function

Public Function getTextAlignment() As fmTextAlign
    getTextAlignment = pTextAlignment
End Function

Public Function getNumberFormat() As String
    getNumberFormat = pNumberFormat
End Function








'[RENDERING]
Public Function render() As UILabel
    Set render = Me
    If Not pIsRendered Then
        Call insertBackgroundLabel
        Call insertCaptionLabel
        Call updateCaption
        pIsRendered = True
    End If
End Function

Private Sub insertBackgroundLabel()
    Set pBackgroundLabel = pFrame.controls.add(CREATE_LABEL_ID)
    With pBackgroundLabel
        .height = pHeight
        .width = pWidth
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .visible = pVisible
        .ControlTipText = pTooltip
    End With
End Sub

Private Sub insertCaptionLabel()
    Set pCaptionLabel = pFrame.controls.add(CREATE_LABEL_ID)
    With pCaptionLabel
        .height = pHeight
        .width = pWidth
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .visible = pVisible
        .ControlTipText = pTooltip
        .WordWrap = False
    End With
    Call updateCaption
End Sub



Public Function Move(Optional x As Variant, Optional y As Variant) As UILabel
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

Public Function offset(Optional x As Single, Optional y As Single) As UILabel
    Set offset = Me
    Call UI.Forms.offsetControls(x, y, pBackgroundLabel, pCaptionLabel)
    With pBackgroundLabel
        pTop = .top
        pLeft = .left
    End With
End Function

Public Function alignVertically(ctrl As MSForms.control) As UILabel
    Dim offset As Single
    '------------------------------------------------------------------------------------------------------
    
    Set alignVertically = Me
    
    If UI.controls.isValid(ctrl) Then
        offset = (ctrl.top + (ctrl.height - pHeight) / 2) - pTop
    End If
    
    If offset Then
        If UI.controls.isValid(pBackgroundLabel) Then
            pBackgroundLabel.top = UI.Normalize.position(pBackgroundLabel.top + offset)
        End If
        If UI.controls.isValid(pCaptionLabel) Then
            pCaptionLabel.top = UI.Normalize.position(pCaptionLabel.top + offset)
        End If
    End If
    
End Function

Public Function alignHorizontally(column As DColumn) As UILabel
    Set alignHorizontally = Me
    With column
        If .getTextAlignment = fmTextAlignLeft Then
            pCaptionLabel.left = .getLeft + .getIndent
        ElseIf .getTextAlignment = fmTextAlignCenter Then
            pCaptionLabel.left = .getLeft + (.getWidth - pCaptionLabel.width) / 2
            pCaptionLabel.TextAlign = fmTextAlignCenter
        ElseIf .getTextAlignment = fmTextAlignRight Then
            pCaptionLabel.left = .getLeft + (.getWidth - pCaptionLabel.width) - .getIndent
            pCaptionLabel.TextAlign = fmTextAlignRight
        End If
    End With
End Function

Public Sub delete()
    On Error Resume Next
    With pFrame.controls
        Call .remove(pBackgroundLabel.name)
        Call .remove(pCaptionLabel.name)
    End With
End Sub


Private Sub updateSize()
    If pIsRendered Then
        pBackgroundLabel.width = pWidth
        Call updateCaption
    End If
End Sub

Private Sub updateCaption()
    If UI.controls.isValid(pCaptionLabel) Then
        With pCaptionLabel
            .caption = VBA.format(pCaption, pNumberFormat)
            '.WordWrap = (VBA.Instr(1, .caption, vba.vbcrlf) > 0)
            With .font
                .name = pFontName
                .Bold = VBA.IIf(pFontBold, 1, False)
                .size = pFontSize
                .Underline = pIsUnderlined
            End With
            .ForeColor = pFontColor
            .width = pBackgroundLabel.width
            .AutoSize = True
            .AutoSize = False
            .width = pBackgroundLabel.width
            .top = calculateCaptionLabelTop
            .TextAlign = pTextAlignment
            If Not pMousePointer Is Nothing And VBA.Len(pCaption) > 0 Then
                .MousePointer = fmMousePointerCustom
                Set .MouseIcon = Icons.getMousePointer
            Else
                .MousePointer = fmMousePointerDefault
            End If
            
            If .height > pHeight Then
                pBackgroundLabel.height = .height + 2 * CAPTION_MARGIN
            ElseIf pBackgroundLabel.height > pHeight Then
                pBackgroundLabel.height = pHeight
            End If
            
        End With
    End If
End Sub

Private Function calculateCaptionLabelTop() As Single
    If pCaptionLabel.height > pHeight Then
        calculateCaptionLabelTop = pBackgroundLabel.top + CAPTION_MARGIN
    Else
        calculateCaptionLabelTop = pBackgroundLabel.top + (pHeight - pCaptionLabel.height) / 2
    End If
End Function

Private Sub updateVisibility()
    If pIsRendered Then
        pBackgroundLabel.visible = pVisible
        pCaptionLabel.visible = pVisible
    End If
End Sub

Private Sub updateToolTips()
    If pIsRendered Then
        pBackgroundLabel.ControlTipText = pTooltip
        pCaptionLabel.ControlTipText = pTooltip
    End If
End Sub

Private Sub updateTextAlignment()
    If pIsRendered Then
        pCaptionLabel.TextAlign = pTextAlignment
    End If
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













'[Controls' events]

Private Sub pCaptionLabel_Click()
    Call ErrorManager.clear
    If Not pListener Is Nothing Then
        Call pListener.catchEvent(Me, Events.Click.getName)
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

Private Function IControl_getRectangle() As RECT
    Stop
End Function

Private Sub IControl_resize(Optional width As Variant, Optional height As Variant)
    Stop
End Sub

Private Sub IControl_alignVertically(baseCtrl As MSForms.IControl)
    Call alignVertically(baseCtrl)
End Sub

Private Function IControl_move(Optional x As Variant, Optional y As Variant) As IControl
    Set IControl_move = Move(x, y)
End Function

Private Function IControl_offset(Optional x As Single, Optional y As Single) As IControl
    Set IControl_offset = offset(x, y)
End Function

Private Sub IControl_triggerEvent(eventId As Long)
    Stop
End Sub

