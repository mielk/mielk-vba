VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MStandarizerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Private Const CLASS_NAME As String = "MStandarizerView"
''[Caption tags] -------------------------------------------------------------------------------------------
'Private Const WINDOW_TITLE As String = "[StandarizationWindowTitle]"
'Private Const UNASSIGNED_ALIASES_TAG As String = "[UnassignedAliases]"
'Private Const ASSIGNED_ALIASES_TAG As String = "[AssignedAliases]"
'Private Const ASSIGN_BUTTON_CAPTION_TAG As String = "[AssignToStandardNameButtonCaption]"
'Private Const ASSIGN_BUTTON_UNDEFINED_CAPTION_TAG As String = "[AssignToUndefinedButtonCaption]"
'Private Const CREATE_NEW_ALIAS_CAPTION_TAG As String = "[CreateNewAliasButtonCaption]"
'Private Const CREATE_NEW_ALIAS_UNDEFINED_CAPTION_TAG As String = "[CreateNewAliasForUndefinedButtonCaption]"
'Private Const UNASSIGNED_ITEMS_WARNING As String = "[UnassignedItemsWarning]"
'Private Const NEW_NAME_WINDOW_TITLE_TAG As String = "[StandarizationNewNameWindowTitle]"
'Private Const NAME_TAG As String = "[Name]"
''[UI contants] --------------------------------------------------------------------------------------------
'Private Const DEFAULT_WIDTH As Single = 512
'Private Const DEFAULT_PADDING As Single = 6
'Private Const DEFAULT_SAVE_BUTTON_CAPTION As String = "Save"
'Private Const DEFAULT_CANCEL_BUTTON_CAPTION As String = "Cancel"
'Private Const LIST_WIDTH As Single = 300
'Private Const RAW_NAMES_LIST_HEIGHT As Single = 300
'Private Const STANDARD_NAMES_LIST_HEIGHT As Single = 240
''----------------------------------------------------------------------------------------------------------
'Private WithEvents pParent As MStandarizer
''[UI components] ------------------------------------------------------------------------------------------
'Private WithEvents pUserForm As MielkUI.GenericUserForm
'Private pMenuPanel As MSForms.Label
'Private WithEvents pStandardNamesList As MielkUI.BasicUIList
'Private WithEvents pRawNamesList As MielkUI.BasicUIList
'Private pVerticalSeparator As MSForms.Label
'Private WithEvents pNewAliasButton As MielkUI.BootstrapButton
'Private WithEvents pAssignButton As MielkUI.BootstrapButton
''[UI properties] ------------------------------------------------------------------------------------------
'Private pWidth As Single
'Private pSaveButtonCaption As String
'Private pCancelButtonCaption As String
'Private pPadding As Single
'Private pStandardNamesListHeadersDefinitions As MielkCore.ArrayList
'Private pRawNamesListHeadersDefinitions As MielkCore.ArrayList
''[State] --------------------------------------------------------------------------------------------------
'Private pIsRendered As Boolean
'Private pIsCancelled As Boolean
''Private WithEvents pActiveStandardName As MStandardName
''----------------------------------------------------------------------------------------------------------
'
'
'
''[INITIALIZERS]
'Private Sub Class_Initialize()
'    Call setDefaultValues
'End Sub
'
'Private Sub setDefaultValues()
'    pWidth = DEFAULT_WIDTH
'    pPadding = DEFAULT_PADDING
'    pSaveButtonCaption = DEFAULT_SAVE_BUTTON_CAPTION
'    pCancelButtonCaption = DEFAULT_CANCEL_BUTTON_CAPTION
'End Sub
'
'
'
''[SETTERS]
'Public Function setParent(value As MStandarizer) As MStandarizerView
'    Set setParent = Me
'    Set pParent = value
'End Function
'
'Public Function setWidth(value As Single) As MStandarizerView
'    Set setWidth = Me
'    pWidth = value
'End Function
'
'Public Function setSaveButtonCaption(value As String) As MStandarizerView
'    Set setSaveButtonCaption = Me
'    pSaveButtonCaption = value
'End Function
'
'Public Function setCancelButtonCaption(value As String) As MStandarizerView
'    Set setCancelButtonCaption = Me
'    pCancelButtonCaption = value
'End Function
'
'Public Function setRawNamesListHeaderDefinitions(value As MielkCore.ArrayList) As MStandarizerView
'    Set setRawNamesListHeaderDefinitions = Me
'    Set pRawNamesListHeadersDefinitions = value
'End Function
'
'Public Function setStandardNamesListHeaderDefinitions(value As MielkCore.ArrayList) As MStandarizerView
'    Set setStandardNamesListHeaderDefinitions = Me
'    Set pStandardNamesListHeadersDefinitions = value
'End Function
'
'
'
''[GETTERS]
'Public Function isCancelled() As Boolean
'    isCancelled = pIsCancelled
'End Function
'
'Private Function getActiveStandardName() As MStandardName
'    Set getActiveStandardName = pParent.getActiveStandardName
'End Function
'
'
'
''[RENDERING]
'Public Sub show(Optional mode As VBA.FormShowConstants = VBA.vbModeless)
'    Call createUserFormIfNotCreatedYet
'    Call pUserForm.center
'    Call pUserForm.render
'    Call render
'    Call validate
'    Call arrangeControls
'
'    'Feed data
'    Call updateRawNamesListContent
'    Call pAssignButton.disable
'    Call pStandardNamesList.setValues(pParent.getStandardNamesCaptions)
'
'    Call pUserForm.show(vbModal)
'
'End Sub
'
'Private Sub createUserFormIfNotCreatedYet()
'    If pUserForm Is Nothing Then
'        Set pUserForm = MielkUI.createUserForm(pWidth). _
'                                addValidationPanel(isDisplayed:=True). _
'                                addButtonsPanel(pSaveButtonCaption, pCancelButtonCaption). _
'                                setValidatorButtonsConnection(False). _
'                                setCaption(pParent.getText(WINDOW_TITLE))
'    End If
'End Sub
'
'Private Sub render()
'    If Not pIsRendered Then
'        Call insertRawNamesList
'        Call insertStandardNamesList
'        Call insertVerticalSeparator
'        Call insertAssignButton
'        Call insertNewAliasButton
'        pIsRendered = True
'    End If
'End Sub
'
'Private Sub insertStandardNamesList()
'    Set pStandardNamesList = MielkUI.createList(pUserForm.getCanvas). _
'                                setMultiSelectMode(fmMultiSelectSingle). _
'                                setHasHeaderPanel(True).setHeaderDefinitions(pStandardNamesListHeadersDefinitions). _
'                                setWidth(LIST_WIDTH).setHeight(STANDARD_NAMES_LIST_HEIGHT)
'    Call pStandardNamesList.render
'End Sub
'
'Private Sub insertRawNamesList()
'    Set pRawNamesList = MielkUI.createList(pUserForm.getCanvas). _
'                            setMultiSelectMode(fmMultiSelectMulti). _
'                            setHasHeaderPanel(True).setHeaderDefinitions(pRawNamesListHeadersDefinitions). _
'                            setWidth(LIST_WIDTH).setHeight(RAW_NAMES_LIST_HEIGHT)
'    Call pRawNamesList.render
'End Sub
'
'Private Sub insertVerticalSeparator()
'    Set pVerticalSeparator = pUserForm.AddControl(UIControlType_Label, "lblVerticalSeparator", True)
'    With pVerticalSeparator
'        .BackStyle = fmBackStyleOpaque
'        .BackColor = pRawNamesList.getBorderColor
'    End With
'End Sub
'
'Private Sub insertAssignButton()
'    Set pAssignButton = MielkUI.createBootstrapButton(pUserForm).setStyle(BootstrapStyle_Blue).setCaption(pParent.getText(ASSIGN_BUTTON_UNDEFINED_CAPTION_TAG))
'    Call pAssignButton.render
'End Sub
'
'Private Sub insertNewAliasButton()
'    Set pNewAliasButton = MielkUI.createBootstrapButton(pUserForm).setStyle(BootstrapStyle_Blue).setCaption(pParent.getText(CREATE_NEW_ALIAS_UNDEFINED_CAPTION_TAG))
'    Call pNewAliasButton.render
'End Sub
'
'Private Sub arrangeControls()
'    Dim validator As MielkUI.UIControl
'    Dim top As Single
'    Dim formWidth As Single
'    '------------------------------------------------------------------------------------------------------
'
'    Set validator = pUserForm.getValidationPanel
'    If Not validator Is Nothing Then top = validator.getBottom
'
'    Call pRawNamesList.Move(pPadding, top)
'    Call resizeControl(pVerticalSeparator, 1, pRawNamesList.getHeight + pPadding)
'    Call locateControl(pVerticalSeparator, pRawNamesList.getRight + pPadding, top)
'    Call pStandardNamesList.Move(getControlRight(pVerticalSeparator) + pPadding, top)
'
'    With pAssignButton
'        Call .Resize(LIST_WIDTH, 24)
'        Call .Move(Left:=pStandardNamesList.getLeft, top:=pStandardNamesList.getBottom + pPadding)
'    End With
'
'    With pNewAliasButton
'        Call .Resize(LIST_WIDTH, 24)
'        Call .Move(Left:=pStandardNamesList.getLeft, top:=pAssignButton.getBottom + pPadding)
'    End With
'
'    formWidth = pStandardNamesList.getRight + pPadding + USER_FORM_BORDER_WIDTH
'    With pUserForm
'        Call .getButtonsPanel.Move(top:=pRawNamesList.getBottom + pPadding)
'        Call .Resize(Width:=formWidth)
'        Call .refreshView
'    End With
'
'End Sub
'
'
'
'
''[Parent events]
'Public Sub updateActiveStandardName()
'    Call changeRawNamesListHeader(pParent.getActiveStandardName)
'    Call updateRawNamesListContent
'    Call clearStandardNamesListSelection
'End Sub
'
'Private Sub changeRawNamesListHeader(standardName As MStandardName)
'    Dim headerDef As MielkUI.BasicUIListHeaderDef
'    Dim caption As String
'    '------------------------------------------------------------------------------------------------------
'
'    Set headerDef = pRawNamesListHeadersDefinitions.getItem(1)
'    If Not headerDef Is Nothing Then
'        If standardName.isUnassigned Then
'            caption = pParent.getText(UNASSIGNED_ALIASES_TAG)
'        Else
'            caption = MielkCore.FormatString(pParent.getText(ASSIGNED_ALIASES_TAG), standardName.getName)
'        End If
'        Call headerDef.setCaption(caption)
'    End If
'
'End Sub
'
'Private Sub updateRawNamesListContent()
'    Dim values As VBA.Collection
'    '------------------------------------------------------------------------------------------------------
'
'    If Not pRawNamesList Is Nothing Then
'        Set values = pParent.getRawNamesForActiveStandardName
'        With pRawNamesList
'            Call .setValues(values)
'            Call .SetFocus(False, True)
'        End With
'        Call updateNewAliasButton
'    End If
'
'End Sub
'
'Private Sub clearStandardNamesListSelection()
'    If Not pStandardNamesList Is Nothing Then
'        With pStandardNamesList
'            Call .resetSelection(True)
'            Call .loseFocus
'        End With
'        Call updateAssignButton
'    End If
'End Sub
'
'Private Sub updateAssignButton()
'    Dim enable As Boolean
'    '------------------------------------------------------------------------------------------------------
'
'    If Not pStandardNamesList Is Nothing And Not pRawNamesList Is Nothing Then
'        If pStandardNamesList.isAnyItemSelected And pRawNamesList.isAnyItemSelected Then
'            'If pParent.getActiveStandardName.hasRawName(pRawNamesList.getSelectedValue) Then
'                enable = True
'            'End If
'        End If
'    End If
'
'    With pAssignButton
'        If enable Then
'            Call .setCaption(MielkCore.FormatString(pParent.getText(ASSIGN_BUTTON_CAPTION_TAG), pStandardNamesList.getSelectedValue))
'            Call .enable
'        Else
'            Call .setCaption(pParent.getText(ASSIGN_BUTTON_UNDEFINED_CAPTION_TAG))
'            Call .disable
'        End If
'    End With
'
'End Sub
'
'Private Sub updateNewAliasButton()
'    If pRawNamesList.isAnyItemSelected Then
'        With pNewAliasButton
'            Call .setCaption(pParent.getText(CREATE_NEW_ALIAS_CAPTION_TAG))
'            Call .enable
'        End With
'    Else
'        With pNewAliasButton
'            Call .setCaption(pParent.getText(CREATE_NEW_ALIAS_UNDEFINED_CAPTION_TAG))
'            Call .disable
'        End With
'    End If
'End Sub
'
'
'
'
'
'
'
'
'
''[ACTIONS]
'Private Sub changeAssignment()
'    Dim newStandardName As String
'    Dim rawNames As VBA.Collection
'    '------------------------------------------------------------------------------------------------------
'    newStandardName = pStandardNamesList.getSelectedValue
'    Set rawNames = MielkCore.createCollection(pRawNamesList.getSelectedValue)
'    Call pParent.changeAssignment(rawNames, newStandardName)
'End Sub
'
'
'Private Sub createNewAliasAndAssignActiveItems()
'    Dim name As String
'    Dim rawNames As VBA.Collection
'    '------------------------------------------------------------------------------------------------------
'
'    name = getNewNameFromUser(vbNullString)
'    If VBA.Len(name) Then
'        Call pParent.addStandardName(name)
'        Set rawNames = MielkCore.createCollection(pRawNamesList.getSelectedValue)
'        Call pParent.changeAssignment(rawNames, name)
'        Call validate
'    End If
'
'End Sub
'
'Private Function getNewNameFromUser(defaultValue As String) As String
'    Dim valuesEditor As MielkUI.WindowPropertiesEditor
'    '------------------------------------------------------------------------------------------------------
'
'    Set valuesEditor = MielkUI.createWindowPropertiesEditor. _
'                            setWidth(300). _
'                            setCaption(pParent.getText(NEW_NAME_WINDOW_TITLE_TAG)). _
'                            setTranslator(pParent). _
'                            setLines(MielkCore.createCollection(MielkUI.createPropertyEditLine(valuesEditor, NAME_TAG, pParent.getText(NAME_TAG), ValueEditType_TextBox).SetValue(defaultValue). _
'                                addValidationRule(MielkCore.getValidationFunction("notEmpty")).addValidationRule(getValidationFunction("uniqueName")). _
'                                addExternalValueRequiredForValidation(STANDARIZATION_STANDARD_NAMES_TAG, pParent.getStandardNames)))
'    With valuesEditor
'        Call .display(vbModal)
'        getNewNameFromUser = .getValues.item(NAME_TAG)
'    End With
'
'End Function
'
'
'
'
'
''[VALIDATING]
'Public Sub validate()
'    Call pUserForm.updateValidation("IsUndefinedEmpty", isUnassignedEmpty, pParent.getText(UNASSIGNED_ITEMS_WARNING))
'    Call arrangeControls
'End Sub
'
'Private Function isUnassignedEmpty() As Boolean
'    isUnassignedEmpty = pParent.isUnassignedEmpty
'End Function
'
'
'
'
'
''[MODEL EVENTS]
'Private Sub pParent_AfterActiveStandardNameChanged(standardName As MStandardName)
'    Call updateActiveStandardName
'End Sub
'
'Private Sub pParent_AfterAssignmentChanged()
'    Call updateRawNamesListContent
'    Call validate
'End Sub
'
'Private Sub pParent_AfterNewStandardNameAdded(standardName As MStandardName)
'    Call pStandardNamesList.AddItem(standardName.getName, standardName.getName)
'End Sub
'
'
'
'
''[CONTROL EVENTS]
'Private Sub pUserForm_Cancelled()
'    pIsCancelled = True
'    Call pParent.triggerCancel
'End Sub
'
'Private Sub pUserForm_Confirmed()
'    pIsCancelled = False
'    Call pParent.triggerSave
'End Sub
'
'Private Sub pStandardNamesList_AfterDoubleClick()
'    Dim selectedStandardName As String
'    '------------------------------------------------------------------------------------------------------
'    selectedStandardName = pStandardNamesList.getSelectedValue
'    Call pParent.setActiveStandardNameByCaption(selectedStandardName)
'End Sub
'
'Private Sub pStandardNamesList_AfterChange()
'    Call updateAssignButton
'End Sub
'
'Private Sub pRawNamesList_AfterChange()
'    Call updateNewAliasButton
'    Call updateAssignButton
'End Sub
'
'Private Sub pAssignButton_AfterClick(ByRef disposed As Boolean)
'    Call changeAssignment
'    Call clearStandardNamesListSelection
'End Sub
'
'Private Sub pAssignButton_AfterHover()
'    Call pUserForm.setHoveredButton(pAssignButton)
'End Sub
'
'Private Sub pNewAliasButton_AfterClick(ByRef disposed As Boolean)
'    Call createNewAliasAndAssignActiveItems
'End Sub
'
'Private Sub pNewAliasButton_AfterHover()
'    Call pUserForm.setHoveredButton(pNewAliasButton)
'End Sub
