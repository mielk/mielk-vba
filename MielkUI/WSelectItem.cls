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

Private Const CLASS_NAME As String = "WSelectItem"
'[Control names] ------------------------------------------------------------------------------------------
Private Const LIST_BOX_NAME As String = "lstItems"
Private Const SEARCH_ICON_NAME As String = "lblSearchIcon"
Private Const SEARCH_TEXTBOX_NAME As String = "txtSearch"
Private Const SEARCH_CANCEL_BUTTON_NAME As String = "lblSearchCancel"
'[UI constants] -------------------------------------------------------------------------------------------
Private Const DEFAULT_LIST_BOX_WIDTH As Single = 200
Private Const DEFAULT_LIST_BOX_HEIGHT As Single = 200
Private Const DEFAULT_LIST_BOX_BORDER_COLOR As Long = 11842740
Private Const DEFAULT_HAS_TITLE_PANEL As Boolean = True
Private Const DEFAULT_MIN_ITEMS_SELECTED As Long = 0
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
'[UI components] ------------------------------------------------------------------------------------------
Private WithEvents pForm As WTemplate
Attribute pForm.VB_VarHelpID = -1
Private pSearchIcon As MSForms.label
Private WithEvents pSearchTextBox As MSForms.textBox
Attribute pSearchTextBox.VB_VarHelpID = -1
Private WithEvents pSearchCancelButton As MSForms.label
Attribute pSearchCancelButton.VB_VarHelpID = -1
Private WithEvents pListBox As MSForms.listBox
Attribute pListBox.VB_VarHelpID = -1
'[UI properties] ------------------------------------------------------------------------------------------
Private pTitle As String
Private pListBoxWidth As Single
Private pListBoxHeight As Single
Private pListBoxBorderColor As Long
Private pLocationProps As DLocation
'[State] --------------------------------------------------------------------------------------------------
Private pItems As Scripting.Dictionary
Private pInitialSelection As Scripting.Dictionary
Private pSelectedItem As Variant
Private pHasTitlePanel As Boolean
Private pHasSearchPanel As Boolean
Private pMultiSelectAllowed As Boolean
Private pMinItemsSelected As Long
Private pSearchText As String
Private pIsRendered As Boolean
Private pCancelledByUser As Boolean
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    '[Set initial values]
    pHasTitlePanel = DEFAULT_HAS_TITLE_PANEL
    pListBoxHeight = DEFAULT_LIST_BOX_HEIGHT
    pListBoxWidth = DEFAULT_LIST_BOX_WIDTH
    pListBoxBorderColor = DEFAULT_LIST_BOX_BORDER_COLOR
    pMinItemsSelected = DEFAULT_MIN_ITEMS_SELECTED
End Sub

Private Sub Class_Terminate()
    Set pItems = Nothing
    Set pForm = Nothing
    Set pSearchIcon = Nothing
    Set pSearchTextBox = Nothing
    Set pSearchCancelButton = Nothing
    Set pListBox = Nothing
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

Public Function setInitialSelection(value As Scripting.Dictionary) As WSelectItem
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




'[GETTERS]
Public Function getForm() As WTemplate
    Set getForm = pForm
End Function

Public Function getSelectedItem() As Variant
    If pMultiSelectAllowed Then
        If VBA.IsEmpty(pSelectedItem) Then
            Set getSelectedItem = New VBA.Collection
        Else
            Set getSelectedItem = pSelectedItem
        End If
    Else
        Call F.Variables.assign(getSelectedItem, pSelectedItem)
    End If
End Function

Public Function getSelectionDictionary() As Scripting.Dictionary
    Dim i As Long
    Dim key As Variant
    '----------------------------------------------------------------------------------------------------------
    
    Set getSelectionDictionary = F.Dictionaries.Create(False)
    With pListBox
        For i = 0 To .ListCount - 1
            If .selected(i) Then
                key = .list(i)
                Call getSelectionDictionary.Add(key, F.Dictionaries.getItem(pItems, key))
            End If
        Next i
    End With
    
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
        Call .render
    End With
End Sub

Private Sub render()
    If Not pIsRendered Then
        Call insertListBox
        Call insertSearchPanel
        pIsRendered = True
        Call validate
        Call pListBox.setFocus
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
    Dim frame As MSForms.frame:         Set frame = pForm.getContentFrame
    '----------------------------------------------------------------------------------------------------------

    Set pListBox = frame.controls.Add(CREATE_LIST_ID, LIST_BOX_NAME, True)
    With pListBox
        .left = pForm.getContentFrameHorizontalMargin
        .top = pForm.getContentFrameVerticalMargin
        .width = pListBoxWidth
        .height = pListBoxHeight
        .SpecialEffect = fmSpecialEffectFlat
        .borderStyle = fmBorderStyleSingle
        .borderColor = pListBoxBorderColor
        .MultiSelect = VBA.IIf(pMultiSelectAllowed, fmMultiSelectMulti, fmMultiSelectSingle)
    End With

    Call loadItemsToListBox
    Call applyInitialSelection

End Sub

Private Sub loadItemsToListBox()
    Dim key As Variant
    Dim items As Scripting.Dictionary
    '----------------------------------------------------------------------------------------------------------
    
    Call pListBox.clear
    If F.Dictionaries.isNonEmptyDictionary(pItems) Then
        If VBA.Len(pSearchText) = 0 Then Set items = pItems Else Set items = filterItems
        For Each key In items.Keys
            Call pListBox.addItem(key)
        Next key
    End If
    
End Sub

Private Sub applyInitialSelection()
    Dim key As Variant
    Dim i As Long
    '----------------------------------------------------------------------------------------------------------
    
    If F.Dictionaries.isNonEmptyDictionary(pInitialSelection) Then
        With pListBox
            For i = 0 To .ListCount - 1
                key = .list(i)
                .selected(i) = pInitialSelection.Exists(key)
            Next i
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
    
    With pSearchIcon
        .top = F.Numerics.roundTo(verticalMargin + (pSearchTextBox.height - .height) / 2, PIXEL_SIZE)
        .left = F.Numerics.roundTo(horizontalMargin, PIXEL_SIZE)
    End With
    
    With pSearchTextBox
        .top = F.Numerics.roundTo(verticalMargin, PIXEL_SIZE)
        .left = F.Numerics.roundTo(UI.Forms.getRight(pSearchIcon) + SEARCH_CONTROLS_SPACE, PIXEL_SIZE)
        .width = calculateSearchTextBoxWidth
    End With
    
    If pHasTitlePanel Then
        With pSearchCancelButton
            .top = F.Numerics.roundTo(verticalMargin + (pSearchTextBox.height - .height) / 2, PIXEL_SIZE)
            .left = F.Numerics.roundTo(UI.Forms.getRight(pSearchTextBox) + SEARCH_CONTROLS_SPACE, PIXEL_SIZE)
        End With
    End If
    
    With pListBox
        .top = VBA.IIf(pHasSearchPanel, _
                    F.Numerics.roundTo(UI.Forms.getBottom(pSearchTextBox) + verticalMargin, PIXEL_SIZE), _
                    F.Numerics.roundTo(verticalMargin, PIXEL_SIZE))
        .left = F.Numerics.roundTo(horizontalMargin, PIXEL_SIZE)
    End With
    
End Sub

Private Function calculateSearchTextBoxWidth() As Single
    If pHasTitlePanel Then
        calculateSearchTextBoxWidth = pListBoxWidth - 2 * SEARCH_CONTROLS_SPACE - SEARCH_ICON_SIZE - SEARCH_CANCEL_BUTTON_SIZE
    Else
        calculateSearchTextBoxWidth = pListBoxWidth - SEARCH_CONTROLS_SPACE - SEARCH_ICON_SIZE
    End If
End Function


Private Sub validate()
    Dim itemsSelected As Long
    Dim i As Long
    Dim success As Boolean
    Dim validation As ValidationObject
    '----------------------------------------------------------------------------------------------------------
    
    With pListBox
        If pMultiSelectAllowed Then
            If pMinItemsSelected > 0 Then
                For i = 0 To .ListCount - 1
                    If .selected(i) Then itemsSelected = itemsSelected + 1
                Next i
                success = (itemsSelected >= pMinItemsSelected)
            Else
                success = True
            End If
        Else
            success = (.ListIndex >= 0)
        End If
    End With
    
    Set validation = F.Create.ValidationObject(success)
    Call pForm.setValidation(validation)
    
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

Private Sub pListBox_DblClick(ByVal cancel As MSForms.ReturnBoolean)
    Call ErrorManager.clear
    If Not pMultiSelectAllowed Then
        Call updateSelection
    End If
End Sub

Private Sub pForm_Confirm()
    Call ErrorManager.clear
    Call updateSelection
End Sub




Private Sub updateSelection()
    Dim index As Long
    Dim key As String
    '----------------------------------------------------------------------------------------------------------
    
    With pListBox
        If pMultiSelectAllowed Then
            Set pSelectedItem = New VBA.Collection
            For index = 0 To .ListCount - 1
                If .selected(index) Then
                    key = .list(index)
                    Call pSelectedItem.Add(F.Dictionaries.getItem(pItems, key))
                End If
            Next index
            If pForm.isVisible Then Call pForm.hide
        Else
            index = .ListIndex
            If index >= 0 Then
                key = .list(index)
                Call F.Variables.assign(pSelectedItem, F.Dictionaries.getItem(pItems, key))
                If pForm.isVisible Then Call pForm.hide
            End If
        End If
    End With
    
End Sub






'[SEARCH EVENTS]
Private Sub pSearchTextBox_Change()
    Call ErrorManager.clear
    Call updateSearching(pSearchTextBox.value)
End Sub

Private Sub pSearchCancelButton_Click()
    Call ErrorManager.clear
    Call updateSearching(vbNullString)
End Sub

Private Sub updateSearching(value As String)
    pSearchText = value
    pSearchTextBox.value = value
    Call loadItemsToListBox
End Sub

Private Function filterItems() As Scripting.Dictionary
    Dim key As Variant
    '----------------------------------------------------------------------------------------------------------
    
    Set filterItems = F.Dictionaries.Create(False)
    For Each key In pItems.Keys
        If VBA.InStr(1, key, pSearchText, vbTextCompare) Then
            Call filterItems.Add(key, pItems.item(key))
        End If
    Next key
    
End Function

Public Function execute() As Variant
    Call display(pLocationProps)
    Call F.Variables.assign(execute, getSelectedItem)
    Call Class_Terminate
End Function
