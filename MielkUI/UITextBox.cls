VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UITextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IControl

Private Const CLASS_NAME As String = "UITextBox"
'----------------------------------------------------------------------------------------------------------
Private Const CAPTION_MARGIN As Single = 3
'----------------------------------------------------------------------------------------------------------
Private pFrame As MSForms.frame
Private pListener As IListener
Private WithEvents pBackgroundLabel As MSForms.label
Attribute pBackgroundLabel.VB_VarHelpID = -1
Private WithEvents pTextBox As MSForms.textBox
Attribute pTextBox.VB_VarHelpID = -1
Private pEventsCatcher As EventCatcher
'----------------------------------------------------------------------------------------------------------
Private pUuid As String
Private pTag As String
Private pValue As String
Private pIsRendered As Boolean
Private pEventsDisabled As Boolean
'----------------------------------------------------------------------------------------------------------
Private pWidth As Single
Private pHeight As Single
Private pLeft As Single
Private pTop As Single
Private pVisible As Boolean
Private pTooltip As String
Private pTextAlignment As fmTextAlign
'[Font]
Private pFontName As String
Private pFontSize As Single
Private pFontBold As Boolean
Private pFontColor As Long
'[Borders & background]
Private pBorderColor As Long
Private pHoverBackColor As Long
'[Events] -------------------------------------------------------------------------------------------------
Event AfterChange()
Event KeyPressed(ByVal KeyCode As MSForms.ReturnInteger)
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Initialize()
    pUuid = F.Crypto.createUuid
    pVisible = True
End Sub







'[SETTERS]
Public Function setFrame(value As MSForms.frame) As UITextBox
    Set setFrame = Me
    Set pFrame = value
End Function

Public Function setListener(value As IListener) As UITextBox
    Set setListener = Me
    Set pListener = value
End Function

Public Function setTag(value As String) As UITextBox
    Set setTag = Me
    pTag = value
End Function

Public Function setValue(value As String) As UITextBox
    Set setValue = Me
    pValue = value
    Call updateValue
End Function


Public Function setHeight(value As Single) As UITextBox
    Set setHeight = Me
    pHeight = value
    If pIsRendered Then Call updateSize
End Function

Public Function setWidth(value As Single) As UITextBox
    Set setWidth = Me
    pWidth = value
    If pIsRendered Then Call updateSize
End Function

Public Function setTooltip(value As String) As UITextBox
    Set setTooltip = Me
    pTooltip = value
    Call updateToolTips
End Function

Public Function setVisibility(value As Boolean) As UITextBox
    Set setVisibility = Me
    pVisible = value
    Call updateVisibility
End Function

Public Function setFontName(value As String) As UITextBox
    Set setFontName = Me
    pFontName = value
    If pIsRendered Then
        Call updateValue
        Call updateSize
    End If
End Function

Public Function setFontSize(value As Single) As UITextBox
    Set setFontSize = Me
    pFontSize = value
    If pIsRendered Then
        Call updateValue
        Call updateSize
    End If
End Function

Public Function setFontBold(value As Boolean) As UITextBox
    Set setFontBold = Me
    pFontBold = value
    If pIsRendered Then
        Call updateValue
        Call updateSize
    End If
End Function

Public Function setFontColor(value As Long) As UITextBox
    Set setFontColor = Me
    pFontColor = value
    If pIsRendered Then
        Call updateValue
    End If
End Function

Public Function setBorderColor(value As Long) As UITextBox
    Set setBorderColor = Me
    pBorderColor = value
    If pIsRendered Then
        Call updateBorderColor
    End If
End Function

Public Function setHoverBackColor(value As Long) As UITextBox
    Set setHoverBackColor = Me
    pHoverBackColor = value
End Function

Public Function setTextAlignment(value As fmTextAlign) As UITextBox
    Set setTextAlignment = Me
    pTextAlignment = value
    Call updateTextAlignment
End Function





'[GETTERS]

Public Function getTag() As String
    getTag = pTag
End Function

Public Function getValue() As String
    getValue = pValue
End Function

Public Function getLeft() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getLeft = pBackgroundLabel.left
    End If
End Function

Public Function getRight() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getRight = pBackgroundLabel.left + pBackgroundLabel.width
    End If
End Function

Public Function getTop() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getTop = pBackgroundLabel.top
    End If
End Function

Public Function getBottom() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getBottom = pBackgroundLabel.top + pBackgroundLabel.height
    End If
End Function

Public Function getHeight() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getHeight = pBackgroundLabel.height
    End If
End Function

Public Function getWidth() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getWidth = pBackgroundLabel.width
    End If
End Function

Public Function getTextAlignment() As fmTextAlign
    getTextAlignment = pTextAlignment
End Function








'[RENDERING]
Public Function render() As UITextBox
    Set render = Me
    If Not pIsRendered Then
        Call insertBackgroundLabel
        Call insertTextBox
        Call updateValue
        Call updateSize
        pIsRendered = True
    End If
End Function

Private Sub insertBackgroundLabel()
    Set pBackgroundLabel = pFrame.controls.Add(CREATE_LABEL_ID)
    With pBackgroundLabel
        .height = pHeight
        .width = pWidth
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .visible = pVisible
        .ControlTipText = pTooltip
    End With
End Sub

Private Sub insertTextBox()
    Set pTextBox = pFrame.controls.Add(CREATE_TEXTBOX_ID)
    With pTextBox
        .SpecialEffect = fmSpecialEffectFlat
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .visible = pVisible
        .ControlTipText = pTooltip
    End With
    
    '[Create eventsCatcher to be bound to this TextBox]
    Set pEventsCatcher = New EventCatcher
    Call pEventsCatcher.setParentAndControl(Me, pTextBox)
    
End Sub



Public Function Move(Optional x As Variant, Optional y As Variant) As UITextBox
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

Public Function offset(Optional x As Single, Optional y As Single)
    Set offset = Me
    Call UI.Forms.offsetControls(x, y, pBackgroundLabel, pTextBox)
    With pBackgroundLabel
        pTop = .top
        pLeft = .left
    End With
End Function

Public Function alignVertically(ctrl As MSForms.control) As UITextBox
    Dim offset As Single
    '------------------------------------------------------------------------------------------------------
    
    Set alignVertically = Me
    
    If UI.Forms.isControlValid(ctrl) Then
        offset = (ctrl.top + (ctrl.height - pHeight) / 2) - pTop
    End If
    
    If offset Then
        If UI.Forms.isControlValid(pBackgroundLabel) Then
            pBackgroundLabel.top = UI.Normalize.position(pBackgroundLabel.top + offset)
        End If
        If UI.Forms.isControlValid(pTextBox) Then
            pTextBox.top = UI.Normalize.position(pTextBox.top + offset)
        End If
    End If
    
End Function

Public Sub delete()
    On Error Resume Next
    With pFrame.controls
        Call .remove(pBackgroundLabel.name)
        Call .remove(pTextBox.name)
    End With
End Sub


Private Sub updateSize()
    pEventsDisabled = True
    pBackgroundLabel.width = pWidth
    With pTextBox
        .value = "kj"
        .AutoSize = True
        .AutoSize = False
        .value = pValue
        .top = calculateTextBoxTop
        .width = pWidth - 2 * CAPTION_MARGIN
        .left = pBackgroundLabel.left + CAPTION_MARGIN
    End With
End Sub

Private Sub updateValue()
    If Not pEventsDisabled Then
        pEventsDisabled = True
        
        If UI.Forms.isControlValid(pTextBox) Then
            With pTextBox
                .value = pValue
                With .font
                    .name = pFontName
                    .Bold = pFontBold
                    .size = pFontSize
                End With
                .TextAlign = pTextAlignment
                .ForeColor = pFontColor
                RaiseEvent AfterChange
            End With
        End If
    End If
End Sub

Private Function calculateTextBoxTop() As Single
    If pTextBox.height > pHeight Then
        calculateTextBoxTop = pBackgroundLabel.top + CAPTION_MARGIN
    Else
        calculateTextBoxTop = pBackgroundLabel.top + (pHeight - pTextBox.height) / 2
    End If
End Function



Private Sub updateVisibility()
    If pIsRendered Then
        pBackgroundLabel.visible = pVisible
        pTextBox.visible = pVisible
    End If
End Sub

Private Sub updateToolTips()
    If pIsRendered Then
        pBackgroundLabel.ControlTipText = pTooltip
        pTextBox.ControlTipText = pTooltip
    End If
End Sub

Private Sub updateBorderColor()
    If pIsRendered Then
        pBackgroundLabel.borderColor = pBorderColor
    End If
End Sub

Private Sub updateTextAlignment()
    If pIsRendered Then
        pTextBox.TextAlign = pTextAlignment
    End If
End Sub



Public Sub setFocus()
    Call ErrorManager.clear
    Call pTextBox.setFocus
End Sub








'[Controls' events]

Private Sub pTextBox_Change()
    Call ErrorManager.clear
    
    If Not pEventsDisabled Then
        pValue = pTextBox.value
        RaiseEvent AfterChange
        If Not pListener Is Nothing Then
            Call pListener.catchEvent(Me, Events.Click.getName)
        End If
    End If
End Sub

Private Sub pTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call ErrorManager.clear
    
    If Not pEventsDisabled Then
        RaiseEvent KeyPressed(KeyCode)
        If Not pListener Is Nothing Then
            Call pListener.catchEvent(Me, EVENT_KEY_PRESSED, _
                                            F.Dictionaries.createWithItems(False, KeyValue(TAG_KEY, KeyCode)))
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

Private Sub IControl_resize(Optional width As Variant, Optional height As Variant)
    Stop
End Sub

Private Sub IControl_alignVertically(baseCtrl As MSForms.IControl)
    Call alignVertically(baseCtrl)
End Sub

Private Function IControl_move(Optional x As Single, Optional y As Single) As IControl
    Set IControl_move = Move(x, y)
End Function

Private Function IControl_offset(Optional x As Single, Optional y As Single) As IControl
    Set IControl_offset = offset(x, y)
End Function

Private Sub IControl_triggerEvent(eventId As Long)
    Call handleEvent(eventId)
End Sub




Private Sub handleEvent(eventId As Long)
    If eventId = Events.Enter.getId Then
        Call triggerEnterEvent
    ElseIf eventId = Events.Leave.getId Then
        Call triggerLeaveEvent
    End If
End Sub

Private Sub triggerLeaveEvent()
    If Not pListener Is Nothing Then
        Call pListener.catchEvent(Me, Events.Leave.getName)
    End If
End Sub

Private Sub triggerEnterEvent()
    If Not pListener Is Nothing Then
        Call pListener.catchEvent(Me, Events.Enter.getName)
    End If
End Sub
