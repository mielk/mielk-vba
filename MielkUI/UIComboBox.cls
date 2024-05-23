VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IValueControl
Implements IControl

'Private Enum ControlTypeEnum
'    ControlType_Background = 1
'    ControlType_ComboBox = 2
'End Enum

Private Const CLASS_NAME As String = "UIComboBox"
'[UI constants] -------------------------------------------------------------------------------------------
Private Const CAPTION_MARGIN As Single = 3
'[Error messages] -----------------------------------------------------------------------------------------
Private Const ERR_EMPTY_VALUE As String = "UiComboBox.Errors.EmptyValue"
Private Const ERR_INVALID_VALUE As String = "UiComboBox.Errors.InvalidValue"
'[Default values] -----------------------------------------------------------------------------------------
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_FONT_FAMILY As String = "Century Gothic"
Private Const DEFAULT_AUTO_SIZE As Boolean = False
Private Const DEFAULT_HORIZONTAL_PADDING As Single = 1.5
Private Const DEFAULT_VERTICAL_PADDING As Single = 0.75
Private Const DEFAULT_BACK_STYLE As Variant = MSForms.fmBackStyleOpaque
Private Const DEFAULT_BACK_COLOR As Long = VBA.vbWhite
Private Const DEFAULT_INVALID_BACK_COLOR As Long = &HC0C0FF
Private Const DEFAULT_BORDER_STYLE As Variant = MSForms.fmBorderStyleSingle
Private Const DEFAULT_BORDER_COLOR As Long = &H80000006
Private Const DEFAULT_MATCH_REQUIRED As Boolean = False
Private Const DEFAULT_MATCH_ENTRY As Long = MSForms.fmMatchEntryNone
Private Const DEFAULT_TEXT_ALIGN As Variant = MSForms.fmTextAlignCenter
Private Const DEFAULT_IS_REQUIRED As Boolean = True
'----------------------------------------------------------------------------------------------------------
Private pFrame As MSForms.frame
Private pListener As IListener
Private WithEvents pBackgroundLabel As MSForms.label
Attribute pBackgroundLabel.VB_VarHelpID = -1
'Private WithEvents pTextBox As MSForms.textBox
Private WithEvents pComboBox As MSForms.comboBox
Attribute pComboBox.VB_VarHelpID = -1
Private pEventsCatcher As EventCatcher
'[State] --------------------------------------------------------------------------------------------------
Private pUuid As String
Private pTag As String
Private pIsRendered As Boolean
Private pEventsDisabled As Boolean
Private pText As String
Private pItems As Scripting.Dictionary
Private pStatus As Boolean
Private pIsExpanded As Boolean
Private pShiftPressed As Integer
'[UI components] ------------------------------------------------------------------------------------------
'Private pListener As IControlListener
'Private pParent As Variant
'[UI properties] ------------------------------------------------------------------------------------------
Private pWidth As Single
Private pHeight As Single
Private pLeft As Single
Private pTop As Single
Private pVisible As Boolean
Private pTooltip As String
Private pTextAlignment As fmTextAlign
Private pHorizontalPadding As Single
Private pVerticalPadding As Single
'[Font]
Private pFontName As String
Private pFontSize As Single
Private pFontBold As Boolean
Private pFontColor As Long
'[Background]
Private pBackStyle As MSForms.fmBackStyle
Private pBackColor As Long
Private pInvalidBackColor As Long
Private pHoverBackColor As Long
'[Borders]
Private pBorderStyle As MSForms.fmBorderStyle
Private pBorderColor As Long
'[Control-specific properties]
Private pMatchRequired As Boolean
Private pMatchEntry As MSForms.fmMatchEntry
Private pIsRequired As Boolean
Private pDisplayValuesInCapitalLetters As Boolean
'[State] --------------------------------------------------------------------------------------------------
'Private pUuid As String
'Private pValue As Variant
'Private pIsRendered As Boolean
'Private pEventsDisabled As Boolean
'[Events] -------------------------------------------------------------------------------------------------
Event AfterChange()
Event KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
'----------------------------------------------------------------------------------------------------------



'[INITIALIZE]
Private Sub Class_Initialize()
    Call initializeContainers
    Call initializeDefaultValues
End Sub

Private Sub initializeContainers()
    Set pItems = F.dictionaries.Create(False)
End Sub

Private Sub initializeDefaultValues()
    pStatus = True
    pUuid = F.Crypto.createUUID
    pFontSize = DEFAULT_FONT_SIZE
    pFontName = DEFAULT_FONT_FAMILY
    pHorizontalPadding = DEFAULT_HORIZONTAL_PADDING
    pVerticalPadding = DEFAULT_VERTICAL_PADDING
    pBackStyle = DEFAULT_BACK_STYLE
    pBackColor = DEFAULT_BACK_COLOR
    pInvalidBackColor = DEFAULT_INVALID_BACK_COLOR
    pBorderStyle = DEFAULT_BORDER_STYLE
    pBorderColor = DEFAULT_BORDER_COLOR
    pMatchEntry = DEFAULT_MATCH_ENTRY
    pMatchRequired = DEFAULT_MATCH_REQUIRED
    pIsRequired = DEFAULT_IS_REQUIRED
    pTextAlignment = DEFAULT_TEXT_ALIGN
End Sub



'[SETTERS]
Public Function setListener(value As IListener) As UIComboBox
    Set setListener = Me
    Set pListener = value
End Function

Public Function setFrame(value As MSForms.frame) As UIComboBox
    Set setFrame = Me
    Set pFrame = value
End Function

Public Function setTag(value As String) As UIComboBox
    Set setTag = Me
    pTag = value
End Function

Public Function setVisibility(value As Boolean) As UIComboBox
    Set setVisibility = Me
    pVisible = value
    Call updateVisibility
End Function

Public Function setTooltip(value As String) As UIComboBox
    Set setTooltip = Me
    pTooltip = value
    Call updateToolTips
End Function


'[Position]
Public Function setTop(value As Single) As UIComboBox
    Set setTop = Me
    pTop = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBottom(value As Single) As UIComboBox
    Set setBottom = Me
    pTop = value - pHeight
    If pIsRendered Then Call arrangeControls
End Function

Public Function setLeft(value As Single) As UIComboBox
    Set setLeft = Me
    pLeft = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setRight(value As Single) As UIComboBox
    Set setRight = Me
    pLeft = value - pWidth
    If pIsRendered Then Call arrangeControls
End Function


'[Size]
Public Function setWidth(value As Single) As UIComboBox
    Set setWidth = Me
    pWidth = value
    If pIsRendered Then Call updateSize
End Function

Public Function setHeight(value As Single) As UIComboBox
    Set setHeight = Me
    pHeight = value
    If pIsRendered Then Call updateSize
End Function

Public Function setHorizontalPadding(value As Single) As UIComboBox
    Set setHorizontalPadding = Me
    pHorizontalPadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setVerticalPadding(value As Single) As UIComboBox
    Set setVerticalPadding = Me
    pVerticalPadding = value
    If pIsRendered Then Call arrangeControls
End Function



'[Font]
Public Function setFontSize(value As Single) As UIComboBox
    Set setFontSize = Me
    pFontSize = value
    If pIsRendered Then Call updateFont
End Function

Public Function setFontName(value As String) As UIComboBox
    Set setFontName = Me
    pFontName = value
    If pIsRendered Then Call updateFont
End Function

Public Function setFontBold(value As Boolean) As UIComboBox
    Set setFontBold = Me
    pFontBold = value
    If pIsRendered Then Call updateFont
End Function

Public Function setFontColor(value As Long) As UIComboBox
    Set setFontColor = Me
    pFontColor = value
    If pIsRendered Then Call updateFont
End Function

Public Function setTextAlignment(value As MSForms.fmTextAlign) As UIComboBox
    Set setTextAlignment = Me
    pTextAlignment = value
    Call updateTextAlignment
End Function



'[BACKGROUND & BORDERS]
Public Function setBackStyle(value As MSForms.fmBackStyle) As UIComboBox
    Set setBackStyle = Me
    pBackStyle = value
    If pIsRendered Then Call updateBackColor
End Function

Public Function setBackColor(value As Long) As UIComboBox
    Set setBackColor = Me
    pBackColor = value
    If pIsRendered Then Call updateBackColor
End Function

Public Function setInvalidBackColor(value As Long) As UIComboBox
    Set setInvalidBackColor = Me
    pInvalidBackColor = value
    If pIsRendered Then Call updateBackColor
End Function

Public Function setBorderStyle(value As MSForms.fmBorderStyle) As UIComboBox
    Set setBorderStyle = Me
    pBorderStyle = value
    If pIsRendered Then Call updateBorderColor
End Function

Public Function setBorderColor(value As Long) As UIComboBox
    Set setBorderColor = Me
    pBorderColor = value
    If pIsRendered Then Call updateBorderColor
End Function





'[STATE PROPERTIES]

Public Function setIsRequired(value As Boolean) As UIComboBox
    Set setIsRequired = Me
    pIsRequired = value
End Function

Public Function setMatchRequired(value As Boolean) As UIComboBox
    Set setMatchRequired = Me
    pMatchRequired = value
    If pIsRendered Then Call updateComboBoxState
End Function

Public Function setMatchEntry(value As MSForms.fmMatchEntry) As UIComboBox
    Set setMatchEntry = Me
    pMatchEntry = value
    If pIsRendered Then Call updateComboBoxState
End Function

Public Function setDisplayValuesInCapitalLetters(value As Boolean) As UIComboBox
    Set setDisplayValuesInCapitalLetters = Me
    pDisplayValuesInCapitalLetters = value
    If pIsRendered Then Call updateComboBoxState
End Function



Public Function setValue(value As Variant) As UIComboBox
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------

    Set setValue = Me
    key = getKeyForValue(value)
    
    If Not F.Variables.isUndefined(key) Then
        pText = key
    Else
        pText = VBA.vbNullString
    End If

    If pIsRendered Then Call updateComboBoxState

End Function

Public Function setText(text As String) As UIComboBox
    Set setText = Me
    If pItems.Exists(text) Then
        pComboBox.text = text
        pText = text
    End If
End Function


Public Function setItems(items As Scripting.Dictionary) As UIComboBox
    Set setItems = Me
    Set pItems = items
    If pIsRendered Then Call loadComboBoxItems
End Function







'[GETTERS]
Public Function isValid() As Boolean
    isValid = UI.Forms.isControlValid(pBackgroundLabel)
End Function

Public Function getTag() As String
    getTag = pTag
End Function




'[SIZE & POSITION]
Public Function getLeft() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getLeft = pBackgroundLabel.left
    End If
End Function

Public Function getRight() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getRight = UI.Forms.getRight(pBackgroundLabel)
    End If
End Function

Public Function getTop() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getTop = pBackgroundLabel.top
    End If
End Function

Public Function getBottom() As Single
    If UI.Forms.isControlValid(pBackgroundLabel) Then
        getBottom = UI.Forms.getBottom(pBackgroundLabel)
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




'[SELECTION]

Public Function getText() As String
    getText = pText
End Function

Public Function getValue() As Variant
    Call F.Variables.assign(getValue, getValueForKey(pText))
End Function

Private Function getKeyForValue(value As Variant) As String
    Dim keysWithGivenValue As VBA.Collection
    '----------------------------------------------------------------------------------------------------------
    Set keysWithGivenValue = F.dictionaries.getKeysForValue(pItems, value)
    If F.Collections.isNonEmptyCollection(keysWithGivenValue) Then
        getKeyForValue = F.Collections.getFirstItem(keysWithGivenValue)
    End If
End Function

Private Function getValueForKey(key As String) As Variant
    If pItems.Exists(key) Then
        Call F.Variables.assign(getValueForKey, pItems.item(key))
    End If
End Function

Public Function getStatus() As Boolean
    getStatus = pStatus And getInternalValidation.getStatus
End Function

Public Function getInternalValidation() As MielkCore.ValidationObject
    Dim text As String
    '----------------------------------------------------------------------------------------------------------
    
    If UI.Forms.isControlValid(pComboBox) Then
        text = pComboBox.value
        
        If VBA.Len(text) = 0 Then
            If pIsRequired Then
                Set getInternalValidation = F.Create.ValidationObject(False, MsgService.getText(ERR_EMPTY_VALUE))
            End If
        Else
            If Not pItems.Exists(text) Then
                Set getInternalValidation = F.Create.ValidationObject(False, MsgService.getText(ERR_INVALID_VALUE))
            End If
        End If
    End If
    
    If getInternalValidation Is Nothing Then
        Set getInternalValidation = F.Create.ValidationObject(True)
    End If
    
End Function

Public Function isShiftPressed() As Integer
    isShiftPressed = pShiftPressed
End Function













'[API]
Public Function render() As UIComboBox
    Set render = Me
    If Not pIsRendered Then
        Call insertBackgroundLabel
        Call insertComboBox
        Call arrangeControls
        Call loadComboBoxItems
        Call validateControls
        pIsRendered = True
    End If
End Function

Private Sub insertBackgroundLabel()
    Set pBackgroundLabel = pFrame.controls.Add(CREATE_LABEL_ID)
    With pBackgroundLabel
        .BackStyle = pBackStyle
        .backColor = VBA.IIf(getStatus, pBackColor, pInvalidBackColor)
        .borderStyle = pBorderStyle
        .borderColor = pBorderColor
        .height = pHeight
        .width = pWidth
        .visible = pVisible
        .ControlTipText = pTooltip
    End With
End Sub

Private Sub insertComboBox()
    Set pComboBox = pFrame.controls.Add(CREATE_COMBOBOX_ID)
    With pComboBox
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .visible = pVisible
        .ControlTipText = pTooltip
        Call updateTextAlignment
        Call updateFont
        Call updateTextAlignment
        Call updateComboBoxState
    End With
    
    '[Create eventsCatcher to be bound to this TextBox]
    Set pEventsCatcher = New EventCatcher
    Call pEventsCatcher.setParentAndControl(Me, pComboBox)
    
End Sub





Public Function Move(Optional x As Variant, Optional Y As Variant) As UIComboBox
    Dim offsetX As Single
    Dim offsetY As Single
    '------------------------------------------------------------------------------------------------------
    
    Set Move = Me
    
    With pBackgroundLabel
        If VBA.IsNumeric(x) Then
            offsetX = (x - .left)
        End If
        If VBA.IsNumeric(Y) Then
            offsetY = (Y - .top)
        End If
    End With
    
    Call offset(offsetX, offsetY)
    
End Function

Public Function offset(Optional x As Single, Optional Y As Single) As UIComboBox
    Set offset = Me
    Call UI.Forms.offsetControls(x, Y, pBackgroundLabel, pComboBox)
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
        If UI.Forms.isControlValid(pComboBox) Then
            pComboBox.top = UI.Normalize.position(pComboBox.top + offset)
        End If
    End If
    
End Function

Private Sub delete()
    On Error Resume Next
    With pFrame.controls
        Call .remove(pBackgroundLabel.name)
        Call .remove(pComboBox.name)
    End With
End Sub



Private Sub updateSize()
    pEventsDisabled = True
    pBackgroundLabel.width = pWidth
    With pComboBox
        .value = "kj"
        .AutoSize = True
        .AutoSize = False
        .value = pText
        '.top = calculateComboBoxTop
        .width = pWidth - 2 * CAPTION_MARGIN
        '.left = pBackgroundLabel.left + CAPTION_MARGIN
    End With
    Call arrangeControls
    pEventsDisabled = False
End Sub

Private Function calculateComboBoxTop() As Single
    If pComboBox.height > pHeight Then
        calculateComboBoxTop = pBackgroundLabel.top + CAPTION_MARGIN
    Else
        calculateComboBoxTop = pBackgroundLabel.top + (pHeight - pComboBox.height) / 2
    End If
End Function

Private Sub updateText()
    If Not pEventsDisabled Then
        pEventsDisabled = True
        If UI.Forms.isControlValid(pComboBox) Then
            With pComboBox
                .value = VBA.IIf(pDisplayValuesInCapitalLetters, VBA.UCase$(pText), pText)
                RaiseEvent AfterChange
            End With
        End If
        pEventsDisabled = False
    End If
End Sub

Private Sub updateFont()
    With pComboBox
        .ForeColor = pFontColor
        .font.name = pFontName
        .font.size = pFontSize
        .font.Bold = VBA.IIf(pFontBold, 1, 0)
    End With
End Sub

Private Sub updateVisibility()
    pBackgroundLabel.visible = pVisible
    pComboBox.visible = pVisible
End Sub

Private Sub updateToolTips()
    pBackgroundLabel.ControlTipText = pTooltip
    pComboBox.ControlTipText = pTooltip
End Sub

Private Sub updateBorderColor()
    With pBackgroundLabel
        .borderColor = pBorderColor
        .borderStyle = pBorderStyle
    End With
End Sub

Private Sub updateBackColor()
    With pBackgroundLabel
        .BackStyle = pBackStyle
        .backColor = pBackColor
    End With
End Sub

Private Sub updateTextAlignment()
    pComboBox.TextAlign = pTextAlignment
End Sub

Private Sub updateComboBoxState()
    With pComboBox
        .MatchEntry = pMatchEntry
        .MatchRequired = pMatchRequired
    End With
End Sub


Public Sub setFocus()
    Call ErrorManager.clear
    On Error Resume Next
    Call pComboBox.setFocus
End Sub






















Public Sub changeStatus(value As Boolean)
    pStatus = value
    Call validateControls
End Sub

Public Sub changeSelection(value As String, Optional skipIfNotAvailable As Boolean = False)
    If Not skipIfNotAvailable Then
        pComboBox.text = value
    Else
        If pItems.Exists(value) Then
            pComboBox.text = value
        End If
    End If
End Sub




'[RENDERING]


Public Sub show()
    On Error Resume Next
    pBackgroundLabel.visible = True
    pComboBox.visible = True
End Sub

Public Sub hide()
    On Error Resume Next
    pBackgroundLabel.visible = False
    pComboBox.visible = False
End Sub


Private Sub arrangeControls()
    Call resizeControls
    Call UI.Forms.locateControl(pBackgroundLabel, pLeft, pTop)
    Call UI.Forms.locateControl(pComboBox, pLeft + (pWidth - pComboBox.width) / 2, pTop + (pHeight - pComboBox.height) / 2)
End Sub

Private Sub resizeControls()
    If pBackgroundLabel.height <> pHeight Or pBackgroundLabel.width <> pWidth Then
        Call UI.Forms.resizeControl(pBackgroundLabel, pWidth, pHeight)
        Call resizeComboBox
    End If
End Sub

Private Sub resizeComboBox()
    Dim maxHeight As Single:                    maxHeight = pHeight - 2 * pVerticalPadding
    '------------------------------------------------------------------------------------------------------
    On Error Resume Next
    With pComboBox
        .width = pWidth - 2 * pHorizontalPadding
        .font.size = pFontSize
        .AutoSize = False
        .AutoSize = True
        If .height > maxHeight Then
            .height = maxHeight
            Call changeFontSizeToFitIntoHeight(pComboBox, maxHeight)
        End If
        .AutoSize = False
        .width = pWidth - 2 * Excel.Application.WorksheetFunction.Max(pHorizontalPadding, (pHeight - .height) / 2)
    End With
End Sub

Private Sub changeFontSizeToFitIntoHeight(ctrl As MSForms.control, height As Single)
    Const CHECK_VALUE As String = "jd"
    Const MIN_FONT_SIZE As Single = 8
    '------------------------------------------------------------------------------------------------------
    Dim saveValue As String
    '------------------------------------------------------------------------------------------------------
    
    pEventsDisabled = True
    saveValue = ctrl.value
    ctrl.text = CHECK_VALUE
    
    With ctrl
        Do
            If .font.size = MIN_FONT_SIZE Then Exit Do
            .font.size = .font.size - 1
            .AutoSize = False
            .AutoSize = True
            If .height <= height Then Exit Do
            If .font.size = 0 Then Exit Do
        Loop
    End With
    
    ctrl.text = saveValue
    pEventsDisabled = False
    
End Sub



'[VALUES]
Private Sub loadComboBoxItems()
    Dim currentValue As String
    Dim key As Variant
    Dim text As String
    '------------------------------------------------------------------------------------------------------

    pEventsDisabled = True

    With pComboBox
    
        'Remember current value to restore it later.
        currentValue = .value
        
        Call .clear
        
        For Each key In pItems.keys
            text = F.Strings.stringify(key)
            If pDisplayValuesInCapitalLetters Then text = VBA.UCase$(text)
            Call .addItem(text)
        Next key
        
        If pItems.Exists(currentValue) Then
            pText = VBA.IIf(pDisplayValuesInCapitalLetters, VBA.UCase$(currentValue), currentValue)
        Else
            pText = vbNullString
        End If
        
        If pIsRendered Then Call updateText
        
    End With

    pEventsDisabled = False

End Sub

Private Sub validateControls()
    pBackgroundLabel.backColor = VBA.IIf(getStatus, pBackColor, pInvalidBackColor)
End Sub







'[EVENTS]
Private Sub changeValue()

    Stop
'
'    If pItems.Exists(pComboBox.text) Then
'        If pText <> pComboBox.text Then
'            Call setText(pComboBox.text)
'            Call validateControls
'            RaiseEvent AfterChange
'            If Not pListener Is Nothing Then Call pListener.catchChangeEvent(Me)
'        End If
'        pShiftPressed = 0
'    Else
'        pComboBox.text = pText
'    End If
End Sub

Private Sub dropdown()
    If Not pIsExpanded Then
        Call pComboBox.dropdown
        pIsExpanded = True
    Else
        pIsExpanded = False
    End If
    'If Not pListener Is Nothing Then Call pListener.catchChangeEvent(Me)
End Sub

Private Sub pComboBox_Change()
Stop
    If Not pEventsDisabled Then Call changeValue
End Sub

Private Sub pBackground_Click()
Stop
    If Not pEventsDisabled Then Call dropdown
End Sub

Private Sub pBackground_MouseDown(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal Y As Single)
Stop
    If Not pEventsDisabled Then Call dropdown
End Sub

Private Sub pComboBox_Click()
Stop
    If Not pEventsDisabled Then Call dropdown
End Sub

Private Sub pComboBox_KeyUp(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    Dim properties As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    If Not pListener Is Nothing Then
        Set properties = createEventPropertiesDict(keyCode, shift)
        Call pListener.catchEvent(Me, Events.KeyUp.getName, properties)
    End If
End Sub

Private Sub pComboBox_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    Dim properties As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    If Not pListener Is Nothing Then
        Set properties = createEventPropertiesDict(keyCode, shift)
        Call pListener.catchEvent(Me, Events.KeyUp.getName, properties)
    End If
    RaiseEvent KeyDown(keyCode, shift)
End Sub

Private Sub pComboBox_MouseDown(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal Y As Single)
Stop
    pShiftPressed = shift
    If Not pEventsDisabled Then Call dropdown
End Sub

Private Function createEventPropertiesDict(keyCode As MSForms.ReturnInteger, shift As Integer) As Scripting.Dictionary
    Set createEventPropertiesDict = F.dictionaries.Create(False)
    With createEventPropertiesDict
        Call .Add(props.keyCode.getName, keyCode)
        Call .Add(props.shift.getName, shift)
    End With
End Function




''[Implementation of UIControl]
''[Setters]
'Private Function UIControl_setParent(value As Variant) As UIControl
'    Set UIControl_setParent = setParent(value)
'End Function
'
'Private Function UIControl_setPlaceholder(value As MSForms.IControl) As UIControl
'    Set UIControl_setPlaceholder = setPlaceholder(value)
'End Function
'
'
''[State]
'Private Function UIControl_isValid() As Boolean
'    UIControl_isValid = isValid
'End Function
'
'
''[Location:setters]
'Private Function UIControl_setTop(value As Single) As UIControl
'    Set UIControl_setTop = setTop(value)
'End Function
'
'Private Function UIControl_setBottom(value As Single) As UIControl
'    Set UIControl_setBottom = setBottom(value)
'End Function
'
'Private Function UIControl_setLeft(value As Single) As UIControl
'    Set UIControl_setLeft = setLeft(value)
'End Function
'
'Private Function UIControl_setRight(value As Single) As UIControl
'    Set UIControl_setRight = setRight(value)
'End Function
'
'Private Sub UIControl_move(Optional left As Variant, Optional top As Variant)
'    Call Move(left, top)
'End Sub
'
'
''[Location:getters]
'Private Function UIControl_getLeft() As Single
'    UIControl_getLeft = getLeft
'End Function
'
'Private Function UIControl_getRight() As Single
'    UIControl_getRight = getRight
'End Function
'
'Private Function UIControl_getTop() As Single
'    UIControl_getTop = getTop
'End Function
'
'Private Function UIControl_getBottom() As Single
'    UIControl_getBottom = getBottom
'End Function
'
'
'
'
''[Size:setters]
'Private Function UIControl_setWidth(value As Single) As UIControl
'    Set UIControl_setWidth = setWidth(value)
'End Function
'
'Private Function UIControl_setHeight(value As Single) As UIControl
'    Set UIControl_setHeight = setHeight(value)
'End Function
'
'Private Sub UIControl_resize(Optional width As Variant, Optional height As Variant)
'    Call resize(width, height)
'End Sub
'
'
'
''[Size:getters]
'Private Function UIControl_getWidth() As Single
'    UIControl_getWidth = getWidth
'End Function
'
'Private Function UIControl_getHeight() As Single
'    UIControl_getHeight = getHeight
'End Function
'
'




'Private Sub UIControl_clear()
'    Call clear
'End Sub
'
'Private Sub UIControl_show()
'    Call show
'End Sub
'
'Private Sub UIControl_hide()
'    Call hide
'End Sub
'
'





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

Private Function IControl_move(Optional x As Variant, Optional Y As Variant) As IControl
    Set IControl_move = Move(x, Y)
End Function

Private Function IControl_offset(Optional x As Single, Optional Y As Single) As IControl
    Set IControl_offset = offset(x, Y)
End Function

Private Sub IControl_triggerEvent(eventId As Long)
    Stop
    'Call handleEvent(eventId)
End Sub


'
'
'Private Sub handleEvent(eventId As Long)
'    If eventId = Events.Enter.getId Then
'        Call triggerEnterEvent
'    ElseIf eventId = Events.Leave.getId Then
'        Call triggerLeaveEvent
'    End If
'End Sub

'Private Sub triggerLeaveEvent()
'    If Not pListener Is Nothing Then
'        Call pListener.catchEvent(Me, Events.Leave.getName)
'    End If
'End Sub

'Private Sub triggerEnterEvent()
'    If Not pListener Is Nothing Then
'        Call pListener.catchEvent(Me, Events.Enter.getName)
'    End If
'End Sub








'[Implementation of IValueControl]
Private Function IValueControl_getInternalValidation() As MielkCore.ValidationObject
    Set IValueControl_getInternalValidation = getInternalValidation.Clone
End Function

Private Function IValueControl_getValue() As Variant
    Call F.Variables.assign(IValueControl_getValue, getValue)
End Function

Private Function IValueControl_getDisplayedValue() As String
    IValueControl_getDisplayedValue = getText
End Function

Private Function IValueControl_getKey() As Variant
    IValueControl_getKey = getText
End Function

Private Sub IValueControl_setValue(value As Variant, Optional displayValue As String)
    Call setValue(value)
End Sub
