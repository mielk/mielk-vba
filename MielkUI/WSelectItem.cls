VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSelectItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IListener

Private Const CLASS_NAME As String = "WSelectItem"
'[Control names] ------------------------------------------------------------------------------------------
Private Const LIST_BOX_NAME As String = "lstItems"
Private Const SEARCH_ICON_NAME As String = "lblSearchIcon"
Private Const SEARCH_TEXTBOX_NAME As String = "txtSearch"
Private Const SEARCH_CANCEL_BUTTON_NAME As String = "lblSearchCancel"
'[Validation label] ---------------------------------------------------------------------------------------
Private Const VALIDATION_LABEL_KEY As String = "lblValidationLabel"
Private Const VALIDATION_LABEL_HEIGHT As Single = 24
Private Const VALIDATION_LABEL_WIDTH As Single = 300
'[UI constants] -------------------------------------------------------------------------------------------
Private Const DEFAULT_LIST_BOX_WIDTH As Single = 200
Private Const DEFAULT_LIST_BOX_HEIGHT As Single = 200
Private Const DEFAULT_LIST_BOX_BORDER_COLOR As Long = 11842740
Private Const DEFAULT_HAS_TITLE_PANEL As Boolean = True
Private Const DEFAULT_MIN_ITEMS_SELECTED As Long = 0
Private Const DEFAULT_MAX_ITEMS_SELECTED As Long = 0
Private Const DEFAULT_VERTICAL_MARGIN As Single = 6
'[Search panel]
Private Const SEARCH_SECTION_LEFT_MARGIN As Single = 36
Private Const SEARCH_ICON_SIZE As Single = 12
Private Const SEARCH_CANCEL_BUTTON_SIZE As Single = 9
Private Const SEARCH_TEXTBOX_HEIGHT As Single = 18
Private Const SEARCH_TEXTBOX_WIDTH As Single = 150
Private Const SEARCH_TEXTBOX_FONT_SIZE As Single = 10
Private Const SEARCH_TEXTBOX_FONT_NAME As String = "Tahoma"
Private Const SEARCH_CONTROLS_SPACE As Single = 3
Private Const SEARCH_TEXTBOX_BORDER_COLOR As Long = 11842740
'[Add item button properties]
Private Const ADD_ITEM_BUTTON_TAG As String = "AddItem"
Private Const ADD_ITEM_BUTTON_HEIGHT As Single = 24
Private Const ADD_ITEM_BUTTON_WIDTH As Single = 90
Private Const ADD_ITEM_BUTTON_FONT_SIZE As Single = 10
Private Const ADD_ITEM_BUTTON_FONT_COLOR As Long = 1381653
Private Const ADD_ITEM_BUTTON_CAPTION As String = "SelectItem.AddItemButton"
Private Const ADD_ITEM_BUTTON_BACK_COLOR As Long = NEUTRAL_BACK_COLOR
Private Const ADD_ITEM_BUTTON_BORDER_COLOR As Long = NEUTRAL_BORDER_COLOR
'[UI components] ------------------------------------------------------------------------------------------
Private WithEvents pForm As WTemplate
Attribute pForm.VB_VarHelpID = -1
Private pSearchIcon As MSForms.label
Private WithEvents pSearchTextBox As MSForms.textBox
Attribute pSearchTextBox.VB_VarHelpID = -1
Private WithEvents pSearchCancelButton As MSForms.label
Attribute pSearchCancelButton.VB_VarHelpID = -1
'Private WithEvents pListBox As MSForms.listBox
Private WithEvents pList As MielkUI.UIList
Attribute pList.VB_VarHelpID = -1
Private WithEvents pAddButton As MielkUI.UIButton
Attribute pAddButton.VB_VarHelpID = -1
Private pValidationLabel As UIValidationLabel
'[UI properties] ------------------------------------------------------------------------------------------
Private pTitle As String
Private pListBoxWidth As Single
Private pListBoxHeight As Single
Private pListBoxBorderColor As Long
Private pLocationProps As DLocation
Private pListHeaders As VBA.Collection
'[State] --------------------------------------------------------------------------------------------------
Private pItems As Scripting.Dictionary
Private pCaptions As Scripting.Dictionary
Private pInitialSelection As Scripting.Dictionary
Private pSelectedItem As Variant
Private pHasTitlePanel As Boolean
Private pHasSearchPanel As Boolean
Private pMultiSelectAllowed As Boolean
Private pMinItemsSelected As Long
Private pMaxItemsSelected As Long
Private pSearchText As String
Private pIsRendered As Boolean
Private pCancelledByUser As Boolean
'----------------------------------------------------------------------------------------------------------
Private pAddingAllowed As Boolean
Private pAddingAction As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    '[Set initial values]
    pHasTitlePanel = DEFAULT_HAS_TITLE_PANEL
    pListBoxHeight = DEFAULT_LIST_BOX_HEIGHT
    pListBoxWidth = DEFAULT_LIST_BOX_WIDTH
    pListBoxBorderColor = DEFAULT_LIST_BOX_BORDER_COLOR
    pMinItemsSelected = DEFAULT_MIN_ITEMS_SELECTED
    pMaxItemsSelected = DEFAULT_MAX_ITEMS_SELECTED
End Sub

Private Sub Class_Terminate()
    Set pItems = Nothing
    Set pForm = Nothing
    Set pSearchIcon = Nothing
    Set pSearchTextBox = Nothing
    Set pSearchCancelButton = Nothing
    Set pList = Nothing
End Sub



'[SETTERS]
Public Function setTitle(value As String) As WSelectItem
    Set setTitle = Me
    pTitle = value
End Function

Public Function setListBoxHeight(value As Single) As WSelectItem
    Set setListBoxHeight = Me
    pListBoxHeight = value
End Function

Public Function setListBoxWidth(value As Single) As WSelectItem
    Set setListBoxWidth = Me
    pListBoxWidth = value
End Function

Public Function setItems(value As Scripting.Dictionary) As WSelectItem
    Set setItems = Me
    Set pItems = value
End Function

Public Function setCaptions(value As Scripting.Dictionary) As WSelectItem
    Set setCaptions = Me
    Set pCaptions = value
End Function

Public Function setListHeaders(value As VBA.Collection) As WSelectItem
    Set setListHeaders = Me
    Set pListHeaders = value
End Function

Public Function setInitialSelection(value As Variant) As WSelectItem
    Set setInitialSelection = Me
    Set pInitialSelection = value
End Function

Public Function setLocationProps(value As DLocation) As WSelectItem
    Set setLocationProps = Me
    Set pLocationProps = value
End Function

Public Function setHasTitlePanel(value As Boolean) As WSelectItem
    Set setHasTitlePanel = Me
    pHasTitlePanel = value
End Function

Public Function setHasSearchPanel(value As Boolean) As WSelectItem
    Set setHasSearchPanel = Me
    pHasSearchPanel = value
    Call updateSearchPanelVisibility
End Function

Public Function setMultiSelectAllowed(value As Boolean) As WSelectItem
    Set setMultiSelectAllowed = Me
    pMultiSelectAllowed = value
End Function

Public Function setMinItemsSelected(value As Long) As WSelectItem
    Set setMinItemsSelected = Me
    pMinItemsSelected = value
End Function

Public Function setMaxItemsSelected(value As Long) As WSelectItem
    Set setMaxItemsSelected = Me
    pMaxItemsSelected = value
End Function

Public Function setAddingAllowed(value As Boolean) As WSelectItem
    Set setAddingAllowed = Me
    pAddingAllowed = value
End Function

Public Function setAddingAction(value As Scripting.Dictionary) As WSelectItem
    Set setAddingAction = Me
    Set pAddingAction = value
End Function





'[GETTERS]
Public Function getForm() As WTemplate
    Set getForm = pForm
End Function

Public Function getSelectedItems() As VBA.Collection
    Dim selection As Variant
    '----------------------------------------------------------------------------------------------------------
    Call F.Variables.assign(selection, pList.getSelectedValue)
    If pCancelledByUser Then
        Set getSelectedItems = Nothing
    ElseIf F.Collections.isNonEmptyCollection(selection) Then
        Set getSelectedItems = selection
    ElseIf Not F.Variables.isUndefined(selection) Then
        Set getSelectedItems = F.Collections.Create(selection)
    Else
        Set getSelectedItems = New VBA.Collection
    End If
End Function

Public Function getSelectionDictionary() As Scripting.Dictionary
    Dim i As Long
    Dim key As Variant
    '----------------------------------------------------------------------------------------------------------
    
    Set getSelectionDictionary = F.dictionaries.Create(False)
    
    Stop
    
'    With pListBox
'        For i = 0 To .ListCount - 1
'            If .selected(i) Then
'                key = .list(i)
'                Call getSelectionDictionary.Add(key, F.dictionaries.getItem(pItems, key))
'            End If
'        Next i
'    End With
    
End Function

Public Function isCancelledByUser() As Boolean
    isCancelledByUser = pCancelledByUser
End Function




'[RENDERING]
Public Sub display(Optional locationProps As DLocation)
    Call createForm
    Call render
    Call updateSearchPanelVisibility
    Call pForm.arrangeComponents
    Call pForm.locateRelativeTo(locationProps)
    Call pForm.display(vbModal, False)
End Sub


Private Sub createForm()
    Set pForm = UI.Create.WindowTemplate
    With pForm
        Call .setCaption(pTitle)
        Call .setHasTitleBar(pHasTitlePanel)
        Call .setWarningsForUnsavedData(False)
        Call .setValidationPanelWidth(VALIDATION_LABEL_WIDTH)
        Call .render
        Call insertValidationLabel
    End With
End Sub

Private Sub insertValidationLabel()
    Set pValidationLabel = pForm.insertValidationLabel(VALIDATION_LABEL_KEY)
    With pValidationLabel
        Call .setHeight(VALIDATION_LABEL_HEIGHT)
    End With
End Sub



Private Sub render()
    If Not pIsRendered Then
        Call insertListBox
        Call insertAddButton
        Call insertSearchPanel
        pIsRendered = True
        Call arrangeControls
        Call validate
    End If
End Sub

Private Sub insertSearchPanel()
    Dim frame As MSForms.frame:         Set frame = pForm.getContentFrame
    '----------------------------------------------------------------------------------------------------------
    
    If pHasSearchPanel Then
        
        Set pSearchIcon = frame.controls.Add(CREATE_LABEL_ID, SEARCH_ICON_NAME, True)
        With pSearchIcon
            .width = SEARCH_ICON_SIZE
            .height = SEARCH_ICON_SIZE
            Set .picture = Icons.icoSearch.picture
        End With
        
        Set pSearchTextBox = frame.controls.Add(CREATE_TEXTBOX_ID, SEARCH_TEXTBOX_NAME, True)
        With pSearchTextBox
            .width = SEARCH_TEXTBOX_WIDTH
            .height = SEARCH_TEXTBOX_HEIGHT
            .borderStyle = fmBorderStyleSingle
            .borderColor = SEARCH_TEXTBOX_BORDER_COLOR
            .font.name = SEARCH_TEXTBOX_FONT_NAME
            .font.size = SEARCH_TEXTBOX_FONT_SIZE
        End With
        
        If pHasTitlePanel Then
            Set pSearchCancelButton = frame.controls.Add(CREATE_LABEL_ID, SEARCH_CANCEL_BUTTON_NAME, False)
            With pSearchCancelButton
                .width = SEARCH_CANCEL_BUTTON_SIZE
                .height = SEARCH_CANCEL_BUTTON_SIZE
                Set .picture = Icons.icoDelete.picture
                .MousePointer = fmMousePointerCustom
                Set .MouseIcon = Icons.getMousePointer
                .visible = False
            End With
        End If
        
    End If
    
End Sub

Private Sub insertListBox()
    Dim frame As MSForms.frame:             Set frame = pForm.getContentFrame
    '----------------------------------------------------------------------------------------------------------

    Set pList = UI.Create.list(frame). _
                            setMultiSelectMode(VBA.IIf(pMultiSelectAllowed, fmMultiSelectMulti, fmMultiSelectSingle)). _
                            setBorderColor(pListBoxBorderColor). _
                            setHasHeaderPanel(False). _
                            setItemsPreviewable(False). _
                            setHeaderDefinitions(pListHeaders). _
                            setWidth(pListBoxWidth). _
                            setHeight(pListBoxHeight)
    Call pList.render
    Call loadItemsToListBox

End Sub

Private Sub loadItemsToListBox()
    Dim key As Variant
    Dim items As Scripting.Dictionary
    '----------------------------------------------------------------------------------------------------------

    Call pList.removeAllItems
    If F.dictionaries.isNonEmptyDictionary(pItems) Then
        If VBA.Len(pSearchText) = 0 Then Set items = pItems Else Set items = filterItems
        Call pList.setValues(items, pInitialSelection, pCaptions)
    End If
    
End Sub

Private Sub insertAddButton()
    Dim frame As MSForms.frame:         Set frame = pForm.getContentFrame
    '----------------------------------------------------------------------------------------------------------

    If pAddingAllowed Then
        Set pAddButton = UI.Create.UIButton(frame)
        With pAddButton
            Call .setTag(ADD_ITEM_BUTTON_TAG)
            Call .setHeight(ADD_ITEM_BUTTON_HEIGHT)
            Call .setWidth(ADD_ITEM_BUTTON_WIDTH)
            Call .setFontSize(ADD_ITEM_BUTTON_FONT_SIZE)
            Call .setFontName(APP_FONT_FAMILY)
            Call .setFontBold(False)
            Call .setFontColor(ADD_ITEM_BUTTON_FONT_COLOR)
            Call .setCaption(MsgService.GetText(ADD_ITEM_BUTTON_CAPTION))
            Call .setBackgroundColor(ADD_ITEM_BUTTON_BACK_COLOR)
            Call .setBorderColor(ADD_ITEM_BUTTON_BORDER_COLOR)
            Call .setVisibility(True)
            Call .setListener(Me)
            Call .render
        End With
    End If
    
End Sub

Private Sub updateSearchPanelVisibility()
    If pIsRendered And pHasSearchPanel Then
        pSearchIcon.visible = pHasSearchPanel
        pSearchTextBox.visible = pHasSearchPanel
        If pHasTitlePanel Then
            pSearchCancelButton.visible = pHasSearchPanel
        End If
        Call arrangeControls
    End If
End Sub

Private Sub arrangeControls()
    Dim verticalMargin As Double:           verticalMargin = DEFAULT_VERTICAL_MARGIN
    Dim horizontalMargin As Double:         horizontalMargin = pForm.getContentFrameHorizontalMargin
    '----------------------------------------------------------------------------------------------------------
    
    If Not pSearchIcon Is Nothing Then
        With pSearchIcon
            .top = UI.Normalize.position(verticalMargin + (pSearchTextBox.height - .height) / 2)
            .left = UI.Normalize.position(horizontalMargin)
        End With
    End If
    
    If Not pSearchTextBox Is Nothing Then
        With pSearchTextBox
            .top = UI.Normalize.position(verticalMargin)
            .left = UI.Normalize.position(UI.Forms.getRight(pSearchIcon) + SEARCH_CONTROLS_SPACE)
            .width = calculateSearchTextBoxWidth
        End With
    End If
    
    If pHasTitlePanel Then
        If Not pSearchCancelButton Is Nothing Then
            With pSearchCancelButton
                .top = UI.Normalize.position(verticalMargin + (pSearchTextBox.height - .height) / 2)
                .left = UI.Normalize.position(UI.Forms.getRight(pSearchTextBox) + SEARCH_CONTROLS_SPACE)
            End With
        End If
    End If
    
    Call pList.move(UI.Normalize.position(horizontalMargin), calculateListBoxTop)

    Call pValidationLabel.setTop(pList.getTop)
    
    If Not pAddButton Is Nothing Then
        Call pAddButton.move(horizontalMargin, pList.getBottom + verticalMargin)
    End If
    
End Sub

Private Function calculateListBoxTop() As Single
    If pHasSearchPanel Then
        calculateListBoxTop = UI.Normalize.position(UI.Forms.getBottom(pSearchTextBox) + DEFAULT_VERTICAL_MARGIN)
    Else
        calculateListBoxTop = UI.Normalize.position(DEFAULT_VERTICAL_MARGIN)
    End If
End Function

Private Function calculateSearchTextBoxWidth() As Single
    If pHasTitlePanel Then
        calculateSearchTextBoxWidth = pListBoxWidth - 2 * SEARCH_CONTROLS_SPACE - SEARCH_ICON_SIZE - SEARCH_CANCEL_BUTTON_SIZE
    Else
        calculateSearchTextBoxWidth = pListBoxWidth - SEARCH_CONTROLS_SPACE - SEARCH_ICON_SIZE
    End If
End Function


Private Sub validate()
    Dim itemsSelected As Variant
    Dim i As Long
    Dim success As Boolean
    Dim validation As ValidationObject
    Dim selectionCounter As Long
    '----------------------------------------------------------------------------------------------------------

    With pList
        If pMultiSelectAllowed Then
            Call F.Variables.assign(itemsSelected, .getSelectedValue)
            If F.Collections.isNonEmptyCollection(itemsSelected) Then
                selectionCounter = itemsSelected.Count
            End If
            
            If pMinItemsSelected > 0 And selectionCounter < pMinItemsSelected Then
                Set validation = F.Create.ValidationObject(False, _
                                       F.Strings.format(MsgService.GetText("SelectItem.ValidationError.MinSelectionNotMetSingle"), pMinItemsSelected))
            ElseIf pMaxItemsSelected > 0 And selectionCounter > pMaxItemsSelected Then
                Set validation = F.Create.ValidationObject(False, _
                                       F.Strings.format(MsgService.GetText("SelectItem.ValidationError.MaxSelectionNotMetSingle"), pMaxItemsSelected))
            Else
                Set validation = F.Create.ValidationObject(True)
            End If
        ElseIf Not F.Variables.isUndefined(.getSelectedKey) Then
            Set validation = F.Create.ValidationObject(True)
        Else
            Set validation = F.Create.ValidationObject(False, MsgService.GetText("SelectItem.ValidationError.NoItemSelected"))
        End If
    End With
    
    If validation.getStatus = False Then
        Call pValidationLabel.setCaption(validation.getMessage)
        Call pValidationLabel.setVisible(True)
    Else
        Call pValidationLabel.setVisible(False)
    End If
    
    Call pForm.setValidation(validation)
    
End Sub

Private Sub setFocusOnList()
    Call pList.setFocus(False, True, False)
End Sub




Private Sub pForm_AfterDisplayed()
    Call setFocusOnList
End Sub

'[EVENTS]
Private Sub pForm_Cancel()
    Call ErrorManager.clear
    pCancelledByUser = True
End Sub

Private Sub pListBox_Change()
    Call ErrorManager.clear
    Call validate
End Sub

Private Sub pList_AfterDoubleClick()
    Call ErrorManager.clear
    If Not pMultiSelectAllowed Then
        Call pForm.hide
    End If
End Sub

Private Sub pForm_Confirm()
    Call ErrorManager.clear
End Sub




Private Sub pList_AfterChange()
    Call validate
End Sub

Private Sub pList_GiveFocusToTopControl()
    Call pList.updateActiveLine(Nothing)
    Call pSearchTextBox.setFocus
End Sub

'[SEARCH EVENTS]
Private Sub pSearchTextBox_Change()
    Call ErrorManager.clear
    Call updateSearching(pSearchTextBox.value)
End Sub

Private Sub pSearchTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDown Then
        Call pList.setFocus(False, True, False)
    ElseIf KeyCode = vbKeyTab Then
        Call pList.setFocus(False, True, False)
    End If
End Sub

Private Sub pSearchCancelButton_Click()
    Call ErrorManager.clear
    Call updateSearching(vbNullString)
End Sub

Private Sub updateSearching(value As String)
    pSearchText = value
    pSearchTextBox.value = value
    Call pList.filterItems(value)
End Sub

Private Function filterItems() As Scripting.Dictionary
    Dim key As Variant
    '----------------------------------------------------------------------------------------------------------
    
    Set filterItems = F.dictionaries.Create(False)
    For Each key In pItems.keys
        If VBA.InStr(1, key, pSearchText, vbTextCompare) Then
            Call filterItems.Add(key, pItems.item(key))
        End If
    Next key
    
End Function

Public Function execute() As VBA.Collection
    Call display(pLocationProps)
    Set execute = getSelectedItems
    Call Class_Terminate
End Function


Private Sub handleAddingItem()
    Dim methodName As String
    Dim methodInvoker As Object
    Dim item As Variant
    '----------------------------------------------------------------------------------------------------------
    Dim propName As String
    Dim key As String
    '----------------------------------------------------------------------------------------------------------
    
    If F.dictionaries.isNonEmptyDictionary(pAddingAction) Then
        methodName = F.dictionaries.getItem(pAddingAction, TAG_METHOD_NAME)
        Set methodInvoker = F.dictionaries.getObject(pAddingAction, TAG_METHOD_INVOKER)
        propName = F.dictionaries.getItem(pAddingAction, TAG_PROPERTY)
        
        If methodInvoker Is Nothing Then
            'missing invoker object
        ElseIf VBA.Len(methodName) = 0 Then
            'missing method name
        Else
            Set item = VBA.CallByName(methodInvoker, methodName, VbMethod, Nothing)
            If Not item Is Nothing Then
                key = item.getValue(propName)
                With pList
                    Call .addItem(key, item)
                    Call .selectItemByKey(key)
                    Call validate
                End With
            End If
        End If
        
    End If
    
End Sub




'[Implementation of IListener interface]

Private Sub IListener_catchEvent(Source As Variant, eventName As String, Optional properties As Scripting.IDictionary)
    If Source Is pAddButton Then
        Call handleAddingItem
    End If
End Sub
