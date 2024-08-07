VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xxPropertyLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "xxPropertyLine"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_FONT_FAMILY As String = APP_FONT_FAMILY
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_FONT_COLOR As Long = vbBlack
Private Const DEFAULT_FONT_BOLD As Boolean = False
Private Const DEFAULT_ALIGNMENT As Variant = fmTextAlignCenter
'[Textbox]
Private Const VALID_TEXTBOX_BACK_COLOR As Long = 16777215
Private Const VALID_TEXTBOX_BORDER_COLOR As Long = 0
Private Const VALID_TEXTBOX_FONT_COLOR As Long = 0
Private Const WARNING_TEXTBOX_BACK_COLOR As Long = 13035765         'rgb(245, 232, 198)
Private Const WARNING_TEXTBOX_BORDER_COLOR As Long = 0
Private Const WARNING_TEXTBOX_FONT_COLOR As Long = 0
Private Const INVALID_TEXTBOX_BACK_COLOR As Long = 12632319
Private Const INVALID_TEXTBOX_BORDER_COLOR As Long = 166
Private Const INVALID_TEXTBOX_FONT_COLOR As Long = 166
Private Const READ_ONLY_TEXTBOX_BACK_COLOR As Long = 15132390
'[Action button]
Private Const DEFAULT_BUTTON_WIDTH As Single = 10.5
Private Const DEFAULT_BUTTON_HEIGHT As Single = 10.5
Private Const DEFAULT_BUTTON_LEFT_MARGIN As Single = 3
'----------------------------------------------------------------------------------------------------------
Private pParent As XPropertiesPanel
Private pUuid As String
Private pDictDefinition As Scripting.Dictionary
Private pItems As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------
Private pLabel As MSForms.label
Private WithEvents pTextBox As MSForms.textBox
Attribute pTextBox.VB_VarHelpID = -1
Private WithEvents pCheckBox As MSForms.Checkbox
Attribute pCheckBox.VB_VarHelpID = -1
Private WithEvents pComboBox As MSForms.comboBox
Attribute pComboBox.VB_VarHelpID = -1
Private WithEvents pActionButton As MSForms.label
Attribute pActionButton.VB_VarHelpID = -1
Private pValueControl As MSForms.control
Private pValidationLabel As UIValidationLabel
'[UI properties] ------------------------------------------------------------------------------------------
Private pTop As Single
Private pLeft As Single
Private pFontFamily As String
Private pFontSize As Single
Private pFontColor As Long
Private pFontBold As Boolean
'----------------------------------------------------------------------------------------------------------
Private pIsRendered As Boolean
Private pEventsDisabled As Boolean
Private pValidation As ValidationObject
'----------------------------------------------------------------------------------------------------------
Private pValidationParams As Scripting.Dictionary
Private pActionWhenValid As Variant
Private pChangeCallbackMethod As String
Private pUpdateView As Boolean
Private pVisibilityProperty As String
Private pIsLocked As Boolean
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    pUuid = F.Crypto.createUUID
    pFontFamily = DEFAULT_FONT_FAMILY
    pFontSize = DEFAULT_FONT_SIZE
    pFontColor = DEFAULT_FONT_COLOR
    pFontBold = DEFAULT_FONT_BOLD
End Sub



'[SETTERS]
Public Function setParent(value As XPropertiesPanel) As xxPropertyLine
    Set setParent = Me
    Set pParent = value
End Function

Public Function setDefinition(value As Scripting.Dictionary) As xxPropertyLine
    Set setDefinition = Me
    Set pDictDefinition = value
    
    '[Cache some values to improve performance when re-using them later]
    Set pValidationParams = createValidationParamsDictionary( _
                                F.dictionaries.getObject(pDictDefinition, TAG_VALIDATION_PARAMS))
    Call F.Variables.assign(pActionWhenValid, F.dictionaries.getItem(pDictDefinition, TAG_ACTION_WHEN_VALID))
    pChangeCallbackMethod = F.dictionaries.getItem(pDictDefinition, TAG_CHANGE_CALLBACK)
    pUpdateView = F.dictionaries.getItem(pDictDefinition, TAG_UPDATE_VIEW)
    pVisibilityProperty = F.dictionaries.getItem(pDictDefinition, TAG_VISIBILITY_PROPERTY)
    pIsLocked = F.dictionaries.getItem(pDictDefinition, TAG_IS_LOCKED)
End Function

Public Function createValidationParamsDictionary(value As Variant) As Scripting.Dictionary
    If F.dictionaries.isNonEmptyDictionary(value) Then
        Set createValidationParamsDictionary = F.Cast.toDictionary(value)
    Else
        Set createValidationParamsDictionary = F.dictionaries.Create(False)
    End If
    
    Call F.dictionaries.addItem(createValidationParamsDictionary, TAG_PROPERTY, getProperty)
    
End Function

Public Function setTop(value As Single) As xxPropertyLine
    Set setTop = Me
    pTop = value
    Call arrangeControls
End Function

Public Function setLeft(value As Single) As xxPropertyLine
    Set setLeft = Me
    pLeft = value
    Call arrangeControls
End Function

Public Function setItemsList(value As Scripting.Dictionary) As xxPropertyLine
    Dim selectedValue As Variant
    '------------------------------------------------------------------------------------------------------
    Set setItemsList = Me
    Call F.dictionaries.addItem(pDictDefinition, TAG_ITEMS, value)
    
    Call UI.controls.fillListBox(pComboBox, getItems)
    pComboBox.value = getComboBoxDefaultValue
    
End Function

Public Function setIsLocked(value As Boolean) As xxPropertyLine
    Set setIsLocked = Me
    pIsLocked = value
    Call updateValueControlView
End Function




'[GETTERS]
Private Function getFrame() As MSForms.frame
    Set getFrame = pParent.getInnerFrame
End Function

Private Function getValidationFrame() As MSForms.frame
    Set getValidationFrame = pParent.getValidationFrame
End Function

Public Function getRight() As Single
    If UI.controls.isValid(pActionButton) Then
        getRight = UI.controls.getRight(pActionButton)
    ElseIf UI.controls.isValid(pValueControl) Then
        getRight = UI.controls.getRight(pValueControl)
    End If
End Function

Private Function getLabelCaption() As String
    getLabelCaption = F.dictionaries.getItem(pDictDefinition, TAG_CAPTION)
End Function

Public Function getValueKey() As String
    getValueKey = F.dictionaries.getItem(pDictDefinition, TAG_PROPERTY)
End Function

Public Function getValidationWidth() As String
    getValidationWidth = F.dictionaries.getItem(pDictDefinition, TAG_VALIDATION_WIDTH)
End Function

Public Function getNumberFormat() As String
    Const DEFAULT_NUMBER_FORMAT As String = "@"
    '------------------------------------------------------------------------------------------------------
    
    With pDictDefinition
        If .Exists(TAG_NUMBER_FORMAT) Then
            getNumberFormat = .item(TAG_NUMBER_FORMAT)
        Else
            getNumberFormat = DEFAULT_NUMBER_FORMAT
        End If
    End With
    
End Function

Public Function getIsLocked() As Boolean
    With pDictDefinition
        If .Exists(TAG_IS_LOCKED) Then
            getIsLocked = .item(TAG_IS_LOCKED)
        End If
    End With
End Function

Public Function getTextAlign() As fmTextAlign
    With pDictDefinition
        If .Exists(TAG_TEXT_ALIGN) Then
            getTextAlign = convertAlignTextToEnum(.item(TAG_TEXT_ALIGN), True)
        Else
            getTextAlign = DEFAULT_ALIGNMENT
        End If
    End With
End Function

Public Function isReadOnly() As Boolean
    isReadOnly = F.dictionaries.getItem(pDictDefinition, TAG_READ_ONLY)
End Function

Public Function getActionButtonPicture() As Variant
    Set getActionButtonPicture = F.dictionaries.getObject(pDictDefinition, TAG_PICTURE)
End Function

Public Function getControlType() As ControlTypeEnum
    getControlType = F.dictionaries.getItem(pDictDefinition, TAG_CONTROL_TYPE)
End Function

Private Function getItems() As Scripting.Dictionary
    Set getItems = F.dictionaries.sortByKey(F.dictionaries.getObject(pDictDefinition, TAG_ITEMS))
End Function

Private Function getValue() As Variant
    Call F.Variables.assign(getValue, getItem.getValue(getValueKey))
End Function

Public Function getValidation() As ValidationObject
    Set getValidation = pValidation
End Function

Public Function getProperty() As String
    getProperty = F.dictionaries.getItem(pDictDefinition, TAG_PROPERTY)
End Function

Public Function getItem() As IEditItem
    Set getItem = pParent.getEditItem
End Function

Public Function getStatus() As Variant
    If pValidation Is Nothing Then
        getStatus = True
    Else
        getStatus = pValidation.getStatus
    End If
End Function

Private Function getValidationLinks() As VBA.Collection
    Set getValidationLinks = F.dictionaries.getObject(pDictDefinition, TAG_VALIDATION_LINKS)
End Function

Public Function getChangeCallbackMethod() As String
    getChangeCallbackMethod = pChangeCallbackMethod
End Function

Public Function getParent() As XPropertiesPanel
    Set getParent = pParent
End Function

Public Function getValidationLabel() As UIValidationLabel
    Set getValidationLabel = pValidationLabel
End Function

Public Function getTop() As Single
    If UI.controls.isValid(pLabel) Then
        getTop = pLabel.top
    Else
        getTop = pTop
    End If
End Function

Private Function confirmAfterEnterClick() As Boolean
    confirmAfterEnterClick = F.dictionaries.getItem(pDictDefinition, TAG_CONFIRM_WITH_ENTER)
End Function



'[RENDERING]
Public Sub render()
    If Not pIsRendered Then
        pEventsDisabled = True
        Call insertLabel
        Call insertValueControl
        Call insertValidationLabel
        Call validate
        pIsRendered = True
        pEventsDisabled = False
    End If
End Sub
    
Private Sub insertLabel()
    Set pLabel = getFrame.controls.add(CREATE_LABEL_ID)
    With pLabel
        .BackStyle = fmBackStyleTransparent
        .caption = getLabelCaption
        .ForeColor = pFontColor
        With .font
            .name = pFontFamily
            .size = pFontSize
            .Bold = VBA.IIf(pFontBold, 1, False)
        End With
        .width = pParent.getLabelWidth
        .WordWrap = False
        .AutoSize = True
        .width = pParent.getLabelWidth
    End With
End Sub

Private Sub insertValueControl()
    Select Case getControlType
        Case ControlType_Label:                         Call insertTextBox
        Case ControlType_TextBox:                       Call insertTextBox
        Case ControlType_CheckBox:                      Call insertCheckBox
        Case ControlType_ComboBox:                      Call insertComboBox
        Case ControlType_LabelWithActionButton:         Call insertLabelWithActionButton
    End Select
End Sub

Private Sub insertCheckBox()
    Set pCheckBox = getFrame.controls.add(CREATE_CHECKBOX_ID)
    With pCheckBox
        .AutoSize = True
        .width = .height
        .value = getValue
        If .Enabled = pIsLocked Then .Enabled = Not pIsLocked
    End With
    Set pValueControl = pCheckBox
End Sub

Private Sub insertComboBox()
    Set pComboBox = getFrame.controls.add(CREATE_COMBOBOX_ID)
    With pComboBox
        .borderStyle = fmBorderStyleSingle
        .width = pParent.getValueWidth
        .TextAlign = getTextAlign
        With .font
            .name = pFontFamily
            .size = pFontSize
            .Bold = VBA.IIf(pFontBold, 1, False)
        End With
        Call UI.controls.fillListBox(pComboBox, getItems)
        pComboBox.value = getComboBoxDefaultValue
    End With
    
    Set pValueControl = pComboBox
    Call updateValueControlView
    
End Sub

Private Function getComboBoxDefaultValue() As String
    Dim items As Scripting.Dictionary
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set items = getItems
    Call F.Variables.assign(value, getValue)
    
    If Not F.dictionaries.isNonEmptyDictionary(items) Then
        ' Items not loaded.
    ElseIf F.Variables.isUndefined(value) Then
        ' Default value not set.
    Else
        getComboBoxDefaultValue = F.dictionaries.getKeyByItem(items, value)
    End If
End Function

Private Sub insertTextBox()
    Set pTextBox = getFrame.controls.add(CREATE_TEXTBOX_ID)
    With pTextBox
        .borderStyle = fmBorderStyleSingle
        .width = pParent.getValueWidth
        With .font
            .name = pFontFamily
            .size = pFontSize
            .Bold = VBA.IIf(pFontBold, 1, False)
        End With
        .TextAlign = getTextAlign
        .value = VBA.format(getValue, getNumberFormat)
        .Locked = isReadOnly
    End With
    
    Set pValueControl = pTextBox
    Call updateValueControlView
    
End Sub

Private Sub insertLabelWithActionButton()
    Set pTextBox = getFrame.controls.add(CREATE_TEXTBOX_ID)
    With pTextBox
        .borderStyle = fmBorderStyleSingle
        .width = pParent.getValueWidth - DEFAULT_BUTTON_WIDTH - DEFAULT_BUTTON_LEFT_MARGIN
        With .font
            .name = pFontFamily
            .size = pFontSize
            .Bold = VBA.IIf(pFontBold, 1, False)
        End With
        .Locked = pIsLocked
        .TextAlign = getTextAlign
        .value = VBA.format(getValue, getNumberFormat)
    End With
    
    Set pActionButton = getFrame.controls.add(CREATE_LABEL_ID)
    With pActionButton
        .width = DEFAULT_BUTTON_WIDTH
        .height = DEFAULT_BUTTON_HEIGHT
        Set .picture = getActionButtonPicture
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = Icons.getMousePointer
    End With
    
    Set pValueControl = pTextBox
    Call updateValueControlView
    
End Sub


Private Sub updateValueControlView()
    Dim status As Variant:          status = getStatus
    '------------------------------------------------------------------------------------------------------
    
    If UI.controls.isValid(pValueControl) Then
        If TypeOf pValueControl Is MSForms.textBox Or TypeOf pValueControl Is MSForms.comboBox Then
            With pValueControl
                .backColor = getBackColorByStatus(status)
                .borderColor = getBorderColorByStatus(status)
                .ForeColor = getFontColorByStatus(status)
                .Locked = pIsLocked
            End With
        End If
    End If
    
End Sub

Private Function getBackColorByStatus(status As Variant) As Long
    If isReadOnly Then
        getBackColorByStatus = READ_ONLY_TEXTBOX_BACK_COLOR
    ElseIf status = True Then
        getBackColorByStatus = VALID_TEXTBOX_BACK_COLOR
    ElseIf status = False Then
        getBackColorByStatus = INVALID_TEXTBOX_BACK_COLOR
    ElseIf status = WARNING_CODE Then
        getBackColorByStatus = WARNING_TEXTBOX_BACK_COLOR
    End If
End Function

Private Function getBorderColorByStatus(status As Variant) As Long
    If status = True Then
        getBorderColorByStatus = VALID_TEXTBOX_BORDER_COLOR
    ElseIf status = False Then
        getBorderColorByStatus = INVALID_TEXTBOX_BORDER_COLOR
    ElseIf status = WARNING_CODE Then
        getBorderColorByStatus = WARNING_TEXTBOX_BORDER_COLOR
    End If
End Function

Private Function getFontColorByStatus(status As Variant) As Long
    If status = True Then
        getFontColorByStatus = VALID_TEXTBOX_FONT_COLOR
    ElseIf status = False Then
        getFontColorByStatus = INVALID_TEXTBOX_FONT_COLOR
    ElseIf status = WARNING_CODE Then
        getFontColorByStatus = WARNING_TEXTBOX_FONT_COLOR
    End If
End Function



Public Sub updateDisplayValue()
    If UI.controls.isValid(pValueControl) Then
        If pValueControl Is pComboBox Then
            pComboBox.value = getComboBoxDefaultValue
        Else
            pValueControl.value = getValue
        End If
    End If
End Sub

Public Sub updateLockStatus()
    Stop
End Sub

Private Sub insertValidationLabel()
    Dim window As WTemplate
    Dim width As Single
    '------------------------------------------------------------------------------------------------------
    
    If UI.controls.isValid(pValueControl) Then
        Set window = pParent.getWindow
        Set pValidationLabel = window.insertValidationLabel(pUuid)
        
        '[Setting size]
        width = F.Variables.coalesce(VBA.Val(getValidationWidth), pParent.getWindow.getValidationPanelWidth)
        With pValidationLabel
            Call .setSize(width, pValueControl.height)
        End With
    End If
    
End Sub

Public Sub Move(left As Single, top As Single)
    pLeft = left
    pTop = top
    Call arrangeControls
End Sub

Public Sub arrangeControls()
    Dim lineHeight As Single:               lineHeight = pParent.getLineHeight
    '------------------------------------------------------------------------------------------------------
    
    If pIsRendered Then
        Call UI.controls.locate(pLabel, pLeft, pTop + (lineHeight - pLabel.height) / 2)
        Call UI.controls.locate(pValueControl, _
                            UI.controls.getRight(pLabel) + pParent.getLabelValueSpace, _
                            pTop + (lineHeight - pValueControl.height) / 2)
        If UI.controls.isValid(pActionButton) Then
            Call UI.controls.locate(pActionButton, _
                            UI.controls.getRight(pValueControl) + DEFAULT_BUTTON_LEFT_MARGIN, _
                            pTop + (lineHeight - pActionButton.height) / 2)
        End If
        
        If Not pValidationLabel Is Nothing Then
            Call pValidationLabel.setTop(pValueControl.top + PIXEL_SIZE)
        End If
        
    End If
    
End Sub

Public Sub validate(Optional value As Variant, Optional isLinkedValidation As Boolean = False)
    Dim functions As Variant
    Dim varFunction As Variant
    Dim validation As ValidationObject
    Dim status As Variant
    '------------------------------------------------------------------------------------------------------
    
    functions = F.dictionaries.getItem(pDictDefinition, TAG_VALIDATION_FUNCTION)
    If F.Variables.isPrimitive(functions) Then functions = VBA.Array(functions)
    
    If VBA.IsMissing(value) Then
        Call F.dictionaries.addItem(pValidationParams, VALUE_____, getValue, duplicateBehaviour_Override)
    Else
        Call F.dictionaries.addItem(pValidationParams, VALUE_____, value, duplicateBehaviour_Override)
    End If
    
    Set pValidation = F.Create.ValidationObject(True)
    For Each varFunction In functions
        If Not F.Variables.isUndefined(varFunction, True, True) Then
            Set validation = getItem.getValidation(VBA.CStr(varFunction), pValidationParams)
            status = validation.getStatus
            
            If status = True Then
                'don't change
            ElseIf status = WARNING_CODE Then
                Set pValidation = validation
            ElseIf status = False Then
                Set pValidation = validation
                Exit For
            End If
        End If
    Next varFunction
    
    Call updateValueControlView
    Call updateValidationLabel
    
End Sub

Public Sub runActionWhenValid()
    Dim dict As Scripting.Dictionary
    Dim methodName As String
    Dim methodInvoker As Object
    Dim methodParams As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    '[Trigger action-when-valid]
    If Not pIsRendered Then
        '
    ElseIf F.Variables.isUndefined(pActionWhenValid, True) Then
        '
    ElseIf F.dictionaries.isNonEmptyDictionary(pActionWhenValid) Then
        If pValidation.getStatus Then
            Set dict = pActionWhenValid
            methodName = F.dictionaries.getItem(dict, TAG_METHOD_NAME)
            Set methodInvoker = F.dictionaries.getObject(dict, TAG_METHOD_INVOKER)
            Set methodParams = F.dictionaries.createSingleEntryDict(VALUE_____, getValue)
            Call F.Reflection.runMethod(Me, methodName, VBA.IIf(methodInvoker Is Nothing, getItem, methodInvoker), methodParams)
            Call pParent.updateLinesDisplayValues
        End If
    ElseIf F.Variables.isString(pActionWhenValid) Then
        If pValidation.getStatus Then
            Call getItem.runAction(VBA.CStr(pActionWhenValid), pValidationParams)
            Call pParent.updateLinesDisplayValues
        End If
    End If

End Sub

Public Sub runLinkedValidations()
    Dim validationLinks As VBA.Collection
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set validationLinks = getValidationLinks
    If Not validationLinks Is Nothing Then
        For Each value In validationLinks
            Call pParent.validateLineByProperty(VBA.CStr(value), True)
        Next value
    End If
    
End Sub

Private Sub updateValidationLabel()
    If pValidationLabel Is Nothing Then
        'This line is not editable, therefore it has no validation label.
    ElseIf pValidation Is Nothing Then
        'This line has not been validated yet.
    ElseIf pValidation.getStatus = True Then
        Call pValidationLabel.setVisible(False)
    Else
        With pValidationLabel
            Call .setVisible(True)
            Call .setCaption(pValidation.getMessage)
            Call .setState(pValidation.getStatus)
        End With
    End If
End Sub

Public Function isVisible() As Boolean
    If VBA.Len(pVisibilityProperty) Then
        isVisible = pParent.getEditItem.getValue(pVisibilityProperty)
    Else
        isVisible = True
    End If
End Function

Friend Sub setValidationLabelVisibility(value As Boolean)
    If pValidationLabel Is Nothing Then
        '
    ElseIf Not value Then
        Call pValidationLabel.setVisible(False)
    Else
        Call pValidationLabel.setVisible(Not pValidation.getStatus = True)
    End If
End Sub





'[EVENTS]
Private Sub pCheckBox_Change()
    Call ErrorManager.clear
    If Not pEventsDisabled Then
        Call getItem.updateValue(getValueKey, pCheckBox.value)
        Call validate(pCheckBox.value)
        Call runLinkedValidations
        Call pParent.updateValidation
        Call runChangeCallbackMethod
    End If
End Sub

Private Sub pComboBox_Change()
    Call ErrorManager.clear
    If Not pEventsDisabled Then
        Call getItem.updateValue(getValueKey, getComboBoxSelectedItem)
        Call validate(getComboBoxSelectedItem)
        Call runLinkedValidations
        Call pParent.updateValidation
        Call runChangeCallbackMethod
    End If
End Sub

Private Sub runChangeCallbackMethod()
    If VBA.Len(pChangeCallbackMethod) Then
        Call pParent.triggerChangeCallbackMethod(Me)
    End If
    If pUpdateView Then Call pParent.arrangeLines
End Sub

Private Function getComboBoxSelectedItem() As Variant
    Dim name As String
    '------------------------------------------------------------------------------------------------------
    name = pComboBox.value
    Call F.Variables.assign(getComboBoxSelectedItem, F.dictionaries.getItem(getItems, name))
End Function

Private Sub pTextBox_Change()
    Call ErrorManager.clear
    If Not pEventsDisabled Then
        Call getItem.updateValue(getValueKey, pTextBox.value)
        Call validate(pTextBox.value)
        Call runActionWhenValid
        Call runLinkedValidations
        Call pParent.updateValidation
    End If
End Sub

Private Sub pTextBox_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    Dim clipboardText As String
    '------------------------------------------------------------------------------------------------------
    
    If keyCode = vbKeyReturn Then
        If confirmAfterEnterClick Then
            Call pParent.triggerConfirm
        End If
    ElseIf pIsLocked And keyCode = vbKeyV And shift = 2 Then
        clipboardText = F.System.getClipboardContentAsString
        If VBA.Len(clipboardText) Then pTextBox.value = clipboardText
    End If
    
End Sub




Private Sub pActionButton_Click()
    Dim actionParams As Scripting.Dictionary
    Dim actionParent As Object
    Dim methodName As String
    Dim methodParams As Scripting.Dictionary
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Call ErrorManager.clear
    
    Set actionParams = F.dictionaries.getObject(pDictDefinition, TAG_ACTION_PARAMS)
    If F.dictionaries.isNonEmptyDictionary(actionParams) Then
        Set actionParent = F.dictionaries.getObject(actionParams, TAG_PARENT)
        If Not actionParent Is Nothing Then
            methodName = F.dictionaries.getItem(actionParams, TAG_METHOD_NAME)
            Set methodParams = F.objects.coalesce( _
                                    F.dictionaries.getObject(actionParams, TAG_METHOD_PARAMS), _
                                    F.dictionaries.Create(False))
            Call F.dictionaries.addItem(methodParams, VALUE_____, getValue, duplicateBehaviour_Override)
            Call F.dictionaries.addItem(methodParams, TAG_PARENT, pParent, duplicateBehaviour_Override)
            Call F.dictionaries.addItem(methodParams, TAG_WINDOW, pParent.getWindow, duplicateBehaviour_Override)
            value = VBA.CallByName(actionParent, methodName, VbMethod, methodParams)
            
            If VBA.Len(value) Then
                pTextBox.text = value
'                Call getItem.updateValue(getValueKey, value)
'                Call validate
'                Call runLinkedValidations
'                Call pParent.updateValidation
            End If
            
        End If
    End If
    
End Sub


Public Sub setFocus()
    Dim line As xxPropertyLine
    Dim textBox As MSForms.textBox
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    If TypeOf pValueControl Is MSForms.textBox Then
        Set textBox = pValueControl
        With textBox
            .SelStart = 0
            .SelLength = VBA.Len(.value)
        End With
        Call pValueControl.setFocus
    End If
    
End Sub
