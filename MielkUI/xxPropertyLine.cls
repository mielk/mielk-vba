VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xxPropertyLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "xxPropertyLine"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_FONT_FAMILY As String = "Segoe UI"
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_FONT_COLOR As Long = vbBlack
Private Const DEFAULT_FONT_BOLD As Boolean = False
Private Const DEFAULT_ALIGNMENT As Variant = xlCenter
'[Textbox]
Private Const VALID_TEXTBOX_BACK_COLOR As Long = 16777215
Private Const VALID_TEXTBOX_BORDER_COLOR As Long = 0
Private Const VALID_TEXTBOX_FONT_COLOR As Long = 0
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
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    pUuid = F.Crypto.createUuid
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
                                F.Dictionaries.getObject(pDictDefinition, TAG_VALIDATION_PARAMS))
End Function

Public Function createValidationParamsDictionary(value As Variant) As Scripting.Dictionary
    If F.Dictionaries.isNonEmptyDictionary(value) Then
        Set createValidationParamsDictionary = F.Cast.toDictionary(value)
    Else
        Set createValidationParamsDictionary = F.Dictionaries.Create(False)
    End If
    
    Call F.Dictionaries.addItem(createValidationParamsDictionary, TAG_PROPERTY, getProperty)
    
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




'[GETTERS]
Private Function getFrame() As MSForms.frame
    Set getFrame = pParent.getInnerFrame
End Function

Private Function getValidationFrame() As MSForms.frame
    Set getValidationFrame = pParent.getValidationFrame
End Function

Public Function getRight() As Single
    If UI.Forms.isControlValid(pActionButton) Then
        getRight = UI.Forms.getRight(pActionButton)
    ElseIf UI.Forms.isControlValid(pValueControl) Then
        getRight = UI.Forms.getRight(pValueControl)
    End If
End Function

Private Function getLabelCaption() As String
    getLabelCaption = F.Dictionaries.getItem(pDictDefinition, TAG_CAPTION)
End Function

Public Function getValueKey() As String
    getValueKey = F.Dictionaries.getItem(pDictDefinition, TAG_PROPERTY)
End Function

Public Function getValidationWidth() As String
    getValidationWidth = F.Dictionaries.getItem(pDictDefinition, TAG_VALIDATION_WIDTH)
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
            getTextAlign = .item(TAG_TEXT_ALIGN)
        Else
            getTextAlign = DEFAULT_ALIGNMENT
        End If
    End With
End Function

Public Function isReadOnly() As Boolean
    isReadOnly = F.Dictionaries.getItem(pDictDefinition, TAG_READ_ONLY)
End Function

Public Function getActionButtonPicture() As Variant
    Set getActionButtonPicture = F.Dictionaries.getObject(pDictDefinition, TAG_PICTURE)
End Function

Public Function getControlType() As ControlTypeEnum
    getControlType = F.Dictionaries.getItem(pDictDefinition, TAG_CONTROL_TYPE)
End Function

Private Function getItems() As Scripting.Dictionary
    Set getItems = F.Dictionaries.sortByKey(F.Dictionaries.getObject(pDictDefinition, TAG_ITEMS))
End Function

Private Function getValue() As Variant
    getValue = getItem.getValue(getValueKey)
End Function

Public Function getValidation() As ValidationObject
    Set getValidation = pValidation
End Function

Public Function getProperty() As String
    getProperty = F.Dictionaries.getItem(pDictDefinition, TAG_PROPERTY)
End Function

Public Function getItem() As IEditItem
    Set getItem = pParent.getEditItem
End Function

Public Function getStatus() As Boolean
    If pValidation Is Nothing Then
        getStatus = True
    Else
        getStatus = VBA.CBool(pValidation.getStatus)
    End If
End Function

Private Function getValidationLinks() As VBA.Collection
    Set getValidationLinks = F.Dictionaries.getObject(pDictDefinition, TAG_VALIDATION_LINKS)
End Function

Private Function confirmAfterEnterClick() As Boolean
    confirmAfterEnterClick = F.Dictionaries.getItem(pDictDefinition, TAG_CONFIRM_WITH_ENTER)
End Function



'[RENDERING]
Public Sub render()
    If Not pIsRendered Then
        pEventsDisabled = True
        Call insertLabel
        pEventsDisabled = False
    End If
End Sub
    
Private Sub insertLabel()
    Set pLabel = getFrame.controls.Add(CREATE_LABEL_ID)
    With pLabel
        .BackStyle = fmBackStyleTransparent
        .caption = getLabelCaption
        .ForeColor = pFontColor
        With .font
            .name = pFontFamily
            .size = pFontSize
            .Bold = pFontBold
        End With
        .width = pParent.getLabelWidth
        .AutoSize = True
        .AutoSize = False
        .width = pParent.getLabelWidth
    End With
End Sub

Private Sub insertValueControl()
    Select Case getControlType
        Case ControlType_TextBox:                       Call insertTextBox
        Case ControlType_CheckBox:                      Call insertCheckBox
        Case ControlType_ComboBox:                      Call insertComboBox
        Case ControlType_LabelWithActionButton:         Call insertLabelWithActionButton
    End Select
End Sub

Private Sub insertCheckBox()
    Set pCheckBox = getFrame.controls.Add(CREATE_CHECKBOX_ID)
    With pCheckBox
        .AutoSize = True
        .width = .height
        .value = getValue
    End With
    Set pValueControl = pCheckBox
End Sub

Private Sub insertComboBox()
    Set pComboBox = getFrame.controls.Add(CREATE_COMBOBOX_ID)
    With pComboBox
        .borderStyle = fmBorderStyleSingle
        .width = pParent.getValueWidth
        .TextAlign = getTextAlign
        With .font
            .name = pFontFamily
            .size = pFontSize
            .Bold = pFontBold
        End With
        Call UI.Forms.fillListBox(pComboBox, getItems)
        pComboBox.value = VBA.format(getValue, getNumberFormat)
    End With
    
    Set pValueControl = pComboBox
    Call updateValueControlView
    
End Sub

Private Sub insertTextBox()
    Set pTextBox = getFrame.controls.Add(CREATE_TEXTBOX_ID)
    With pTextBox
        .borderStyle = fmBorderStyleSingle
        .width = pParent.getValueWidth
        With .font
            .name = pFontFamily
            .size = pFontSize
            .Bold = pFontBold
        End With
        .TextAlign = getTextAlign
        .value = VBA.format(getValue, getNumberFormat)
        .Locked = isReadOnly
    End With
    
    Set pValueControl = pTextBox
    Call updateValueControlView
    
End Sub

Private Sub insertLabelWithActionButton()
    Set pTextBox = getFrame.controls.Add(CREATE_TEXTBOX_ID)
    With pTextBox
        .borderStyle = fmBorderStyleSingle
        .width = pParent.getValueWidth - DEFAULT_BUTTON_WIDTH - DEFAULT_BUTTON_LEFT_MARGIN
        With .font
            .name = pFontFamily
            .size = pFontSize
            .Bold = pFontBold
        End With
        .Locked = getIsLocked
        .TextAlign = getTextAlign
        .value = VBA.format(getValue, getNumberFormat)
    End With
    
    Set pActionButton = getFrame.controls.Add(CREATE_LABEL_ID)
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
    Dim status As Boolean:          status = getStatus
    '------------------------------------------------------------------------------------------------------
    
    If UI.Forms.isControlValid(pValueControl) Then
        If TypeOf pValueControl Is MSForms.textBox Or TypeOf pValueControl Is MSForms.comboBox Then
            With pValueControl
                .BackColor = VBA.IIf(isReadOnly, READ_ONLY_TEXTBOX_BACK_COLOR, _
                                VBA.IIf(status, VALID_TEXTBOX_BACK_COLOR, INVALID_TEXTBOX_BACK_COLOR))
                .borderColor = VBA.IIf(status, VALID_TEXTBOX_BORDER_COLOR, INVALID_TEXTBOX_BORDER_COLOR)
                .ForeColor = VBA.IIf(status, VALID_TEXTBOX_FONT_COLOR, INVALID_TEXTBOX_FONT_COLOR)
            End With
        End If
    End If
    
End Sub

Private Sub insertValidationLabel()
    Dim window As WTemplate
    Dim width As Single
    '------------------------------------------------------------------------------------------------------
    
    If UI.Forms.isControlValid(pValueControl) Then
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

Private Sub arrangeControls()
    Dim lineHeight As Single:               lineHeight = pParent.getLineHeight
    '------------------------------------------------------------------------------------------------------
    
    If pIsRendered Then
        Call UI.Forms.locateControl(pLabel, pLeft, pTop + (lineHeight - pLabel.height) / 2)
        Call UI.Forms.locateControl(pValueControl, _
                            UI.Forms.getRight(pLabel) + pParent.getLabelValueSpace, _
                            pTop + (lineHeight - pValueControl.height) / 2)
        If UI.Forms.isControlValid(pActionButton) Then
            Call UI.Forms.locateControl(pActionButton, _
                            UI.Forms.getRight(pValueControl) + DEFAULT_BUTTON_LEFT_MARGIN, _
                            pTop + (lineHeight - pActionButton.height) / 2)
        End If
        
        If Not pValidationLabel Is Nothing Then
            Call pValidationLabel.setTop(pValueControl.top + PIXEL_SIZE)
        End If
        
    End If
    
End Sub

Public Sub validate()
    Dim functions As Variant
    Dim varFunction As Variant
    Dim validation As ValidationObject
    '------------------------------------------------------------------------------------------------------
    
    functions = F.Dictionaries.getItem(pDictDefinition, TAG_VALIDATION_FUNCTION)
    If F.Variables.isPrimitive(functions) Then functions = VBA.Array(functions)
    
    Set pValidation = F.Create.ValidationObject(True)
    For Each varFunction In functions
        If Not F.Variables.isUndefined(varFunction, True, True) Then
            Set validation = getItem.getValidation(VBA.CStr(varFunction), pValidationParams)
            If Not validation.getStatus Then
                Set pValidation = validation
                Exit For
            End If
        End If
    Next varFunction
    
    Call updateValueControlView
    Call updateValidationLabel
    
End Sub

Public Sub runLinkedValidations()
    Dim validationLinks As VBA.Collection
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set validationLinks = getValidationLinks
    If Not validationLinks Is Nothing Then
        For Each value In validationLinks
            Call pParent.validateLineByProperty(VBA.CStr(value))
        Next value
    End If
    
End Sub

Private Sub updateValidationLabel()
    If pValidationLabel Is Nothing Then
        'This line is not editable, therefore it has no validation label.
    ElseIf pValidation Is Nothing Then
        'This line has not been validated yet.
    ElseIf pValidation.getStatus Then
        Call pValidationLabel.setVisible(False)
    Else
        With pValidationLabel
            Call .setVisible(True)
            Call .setCaption(pValidation.getMessage)
        End With
    End If
End Sub





'[EVENTS]
Private Sub pCheckBox_Change()
    Call ErrorManager.clear
    If Not pEventsDisabled Then
        Call getItem.updateValue(getValueKey, pCheckBox.value)
        Call validate
        Call runLinkedValidations
        Call pParent.updateValidation
    End If
End Sub

Private Function getComboBoxSelectedItem() As Variant
    Dim name As String
    '------------------------------------------------------------------------------------------------------
    name = pComboBox.value
    Call F.Variables.assign(getComboBoxSelectedItem, F.Dictionaries.getItem(getItems, name))
End Function

Private Sub pTextBox_Change()
    Call ErrorManager.clear
    If Not pEventsDisabled Then
        Call getItem.updateValue(getValueKey, pTextBox.value)
        Call validate
        Call runLinkedValidations
        Call pParent.updateValidation
    End If
End Sub

Private Sub pTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If confirmAfterEnterClick Then
            Call pParent.triggerConfirm
        End If
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
    
    Set actionParams = F.Dictionaries.getObject(pDictDefinition, TAG_ACTION_PARAMS)
    If F.Dictionaries.isNonEmptyDictionary(actionParams) Then
        Set actionParent = F.Dictionaries.getObject(actionParams, TAG_PARENT)
        If Not actionParent Is Nothing Then
            methodName = F.Dictionaries.getItem(actionParams, TAG_METHOD_NAME)
            Set methodParams = F.Objects.coalesce( _
                                    F.Dictionaries.getObject(actionParams, TAG_METHOD_PARAMS), _
                                    F.Dictionaries.Create(False))
            Call F.Dictionaries.addItem(methodParams, VALUE_____, getValue, duplicateBehaviour_Override)
            value = VBA.CallByName(actionParams, methodName, VbMethod, methodParams)
            
            If VBA.Len(value) Then
                pTextBox.text = value
                Call getItem.updateValue(getValueKey, value)
                Call validate
                Call runLinkedValidations
                Call pParent.updateValidation
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
