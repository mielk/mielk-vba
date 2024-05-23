VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IControl

Private Enum ControlTypeEnum
    ControlType_Frame = 1
    ControlType_FocusTextBox = 2
    ControlType_SelectAll = 3
    ControlType_DeselectAll = 4
End Enum

Private Enum VerticalAlignmentEnum
    VerticalAlignment_Top = 1
    VerticalAlignment_Middle = 2
    VerticalAlignment_Bottom = 3
End Enum

Private Const CLASS_NAME As String = "UIList"
'[UI constants] -------------------------------------------------------------------------------------------
Private Const DEFAULT_LINE_HEIGHT As Single = 15
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_FONT_FAMILY As String = APP_FONT_FAMILY
Private Const DEFAULT_PADDING As Single = 6
Private Const DEFAULT_BACK_COLOR As Long = VBA.vbWhite
Private Const DEFAULT_CURRENT_ITEM_BACK_COLOR As Long = 16053492
Private Const DEFAULT_SELECTED_ITEM_BACK_COLOR As Long = 13273922
Private Const DEFAULT_CURRENT_SELECTED_BACK_COLOR As Long = 14597527
Private Const DEFAULT_BORDER_COLOR As Long = &H8000000A
Private Const DEFAULT_MULTI_SELECT_MODE As Long = MSForms.fmMultiSelectSingle
Private Const DEFAULT_FOCUS_BORDER_COLOR As Long = 13470009
Private Const DEFAULT_MASS_SELECTION_BUTTON_SIZE As Single = 6
Private Const DEFAULT_SELECT_ALL_BUTTON_COLOR As Long = 13273922
Private Const DEFAULT_DESELECT_ALL_BUTTON_COLOR As Long = vbWhite
Private Const DEFAULT_HAS_HEADER_ROW As Boolean = True
Private Const DEFAULT_HEADER_HEIGHT As Single = 18
Private Const DEFAULT_SPACE_BETWEEN_HEADER_AND_LIST As Single = 3
Private Const DEFAULT_HEADER_FONT_SIZE As Single = 10
Private Const DEFAULT_HEADER_FONT_FAMILY As String = APP_FONT_FAMILY
Private Const DEFAULT_HEADER_FONT_COLOR As Long = vbBlack
Private Const DEFAULT_HEADER_FONT_BOLD As Boolean = True
Private Const DEFAULT_MAX_ITEMS_PER_PAGE As Long = 200
'[UI components] ------------------------------------------------------------------------------------------
Private pListener As IListener
Private pParent As MSForms.UserForm
Private WithEvents pFrame As MSForms.frame
Attribute pFrame.VB_VarHelpID = -1
Private WithEvents pFocusTextbox As MSForms.textBox
Attribute pFocusTextbox.VB_VarHelpID = -1
Private WithEvents pSelectAllButton As MSForms.label
Attribute pSelectAllButton.VB_VarHelpID = -1
Private WithEvents pDeselectAllButton As MSForms.label
Attribute pDeselectAllButton.VB_VarHelpID = -1
Private pHeadersFrame As MSForms.frame
Private pHeaderLine As UIListHeader
Private pHeaders As VBA.Collection
Private pLines() As UIListLine
Private pLinesByKey As Scripting.Dictionary
Private pKeysByLine As Scripting.Dictionary
Private pVisibleLinesList As VBA.Collection
'[UI position] --------------------------------------------------------------------------------------------
Private pWidth As Single
Private pTop As Single
Private pLeft As Single
Private pHeight As Single
Private pLineHeight As Single
Private pPadding As Single
'[UI look] ------------------------------------------------------------------------------------------------
Private pFontSize As Single
Private pFontFamily As String
Private pBackColor As Long
Private pSelectedItemBackColor As Long
Private pCurrentItemBackColor As Long
Private pCurrentSelectedBackColor As Long
Private pBorderColor As Long
Private pFocusBorderColor As Long
Private pMassSelectionButtonSize As Single
Private pSelectAllButtonColor As Long
Private pDeselectAllButtonColor As Long
Private pMaxItemsPerPage As Long
'[Header] -------------------------------------------------------------------------------------------------
Private pHasHeaderRow As Boolean
Private pSpaceBetweenHeaderAndList As Single
Private pHeaderHeight As Single
Private pHeaderFontSize As Single
Private pHeaderFontFamily As String
Private pHeaderFontBold As Boolean
Private pHeaderFontColor As Long
Private pHeaderDefinitions As VBA.Collection
'[Items] --------------------------------------------------------------------------------------------------
Private pUuid As String
Private pItems As Scripting.Dictionary
Private pFiltered As Scripting.Dictionary
Private pSortedKeys As VBA.Collection
Private pCaptions As Scripting.Dictionary
Private pSpecialFormats As Scripting.Dictionary
Private pConstKeysDict As Scripting.Dictionary
Private pCurrentPage As Long
Private pSelection As Scripting.Dictionary
Private pActiveLine As UIListLine
Private pPreviewItemKey As Variant
Private pFilter As String
'Private pSelectedLine As UIListLine
'Private pSelectedLines As Scripting.Dictionary
'[State] --------------------------------------------------------------------------------------------------
Private pMultiSelectMode As MSForms.fmMultiSelect
Private pItemsRemovable As Boolean
Private pItemsEditable As Boolean
Private pItemsPreviewable As Boolean
Private pIsRendered As Boolean
Private pHasVerticalScrollBar As Boolean
Private pHasMassSelectButtons As Boolean
Private pActionButtonsToLeft As Boolean
'[Events] -------------------------------------------------------------------------------------------------
Event AfterChange()
Event AfterDoubleClick()
Event AfterEditItemClick(item As Variant, ByRef newCaption As String)
Event AfterRemoveItemClick(item As Variant, ByRef confirmed As Boolean)
Event AfterPreviewChanged(key As Variant, item As Variant, index As Long)
Event GiveFocusToTopControl()
'----------------------------------------------------------------------------------------------------------




'[INITIALIZE]
Private Sub Class_Initialize()
    Call initializeContainers
    Call initializeDefaultValues
End Sub

Private Sub initializeContainers()
    Erase pLines
    Set pItems = F.dictionaries.Create(False)
    Set pFiltered = F.dictionaries.Create(False)
    Set pCaptions = F.dictionaries.Create(False)
    Set pSpecialFormats = F.dictionaries.Create(False)
    Set pLinesByKey = F.dictionaries.Create(False)
    Set pKeysByLine = F.dictionaries.Create(False)
    Set pVisibleLinesList = New VBA.Collection
    Set pHeaders = New VBA.Collection
    Set pSelection = F.dictionaries.Create(False)
    Set pConstKeysDict = F.dictionaries.Create(False)
    pCurrentPage = 1
    pFilter = vbNullString
End Sub

Private Sub initializeDefaultValues()
    pUuid = F.Crypto.createUUID
    pHasHeaderRow = DEFAULT_HAS_HEADER_ROW
    pHeaderHeight = DEFAULT_HEADER_HEIGHT
    pHeaderFontSize = DEFAULT_HEADER_FONT_SIZE
    pHeaderFontColor = DEFAULT_HEADER_FONT_COLOR
    pHeaderFontFamily = DEFAULT_HEADER_FONT_FAMILY
    pHeaderFontBold = DEFAULT_HEADER_FONT_BOLD
    pLineHeight = UI.Normalize.size(DEFAULT_LINE_HEIGHT)
    pFontSize = DEFAULT_FONT_SIZE
    pFontFamily = DEFAULT_FONT_FAMILY
    pPadding = DEFAULT_PADDING
    pBackColor = DEFAULT_BACK_COLOR
    pSelectedItemBackColor = DEFAULT_SELECTED_ITEM_BACK_COLOR
    pCurrentItemBackColor = DEFAULT_CURRENT_ITEM_BACK_COLOR
    pCurrentSelectedBackColor = DEFAULT_CURRENT_SELECTED_BACK_COLOR
    pBorderColor = DEFAULT_BORDER_COLOR
    pFocusBorderColor = DEFAULT_FOCUS_BORDER_COLOR
    pMultiSelectMode = DEFAULT_MULTI_SELECT_MODE
    pMassSelectionButtonSize = DEFAULT_MASS_SELECTION_BUTTON_SIZE
    pSelectAllButtonColor = DEFAULT_SELECT_ALL_BUTTON_COLOR
    pDeselectAllButtonColor = DEFAULT_DESELECT_ALL_BUTTON_COLOR
    pSpaceBetweenHeaderAndList = DEFAULT_SPACE_BETWEEN_HEADER_AND_LIST
    pMaxItemsPerPage = DEFAULT_MAX_ITEMS_PER_PAGE
    pHasMassSelectButtons = True
    pActionButtonsToLeft = True
End Sub






'[SETTERS]
Public Function setListener(value As IListener) As UIList
    Set setListener = Me
    Set pListener = value
End Function

Public Function setParent(value As Variant) As UIList
    Set setParent = Me
    Set pParent = value
End Function

Public Function setHasHeaderPanel(value As Boolean) As UIList
    Set setHasHeaderPanel = Me
    pHasHeaderRow = value
End Function

Public Function setHeaderDefinitions(value As Variant) As UIList
    Set setHeaderDefinitions = Me
    Set pHeaderDefinitions = F.Collections.createFrom(value)
End Function



'[Values]
Public Function setValues(value As Variant, _
                            Optional selection As Variant, _
                            Optional captions As Scripting.Dictionary, _
                            Optional constKeys As Scripting.Dictionary, _
                            Optional specialFormats As Scripting.Dictionary) As UIList
    
    Set setValues = Me
    Set pItems = F.dictionaries.createFrom(value, duplicateBehaviour_Skip)
    Set pFiltered = F.dictionaries.Clone(pItems)
    Call loadSelectionFromValue(selection)
    Set pCaptions = VBA.IIf(Not captions Is Nothing, captions, F.dictionaries.Create(False))
    Set pConstKeysDict = VBA.IIf(Not constKeys Is Nothing, constKeys, F.dictionaries.Create(False))
    Set pSpecialFormats = VBA.IIf(Not specialFormats Is Nothing, specialFormats, F.dictionaries.Create(False))
    If pIsRendered Then
        Call renderLines
        Call updatePagination
        Call sortItems(True)
        Call updateLinesItems
    End If
End Function

Private Sub loadSelectionFromValue(value As Variant)
    Dim key As Variant
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Set pSelection = F.dictionaries.Create(False)
    If VBA.IsMissing(value) Then
        'Nothing to add
    ElseIf F.Variables.isContainer(value) Then
        Set dict = F.dictionaries.createFrom(value)
        For Each key In dict
            If pItems.Exists(key) Then
                Call F.dictionaries.addItem(pSelection, key, pItems.item(key))
            End If
        Next key
    Else
        If pItems.Exists(key) Then
            Call F.dictionaries.addItem(pSelection, key, pItems.item(key))
        End If
    End If
End Sub

Public Function setMultiSelectMode(value As MSForms.fmMultiSelect) As UIList
    Set setMultiSelectMode = Me
    pMultiSelectMode = value
End Function

Public Function setMaxItemsPerPage(value As Long) As UIList
    Set setMaxItemsPerPage = Me
    pMaxItemsPerPage = value
End Function

Public Function setSelection(value As Variant) As UIList
    Set setSelection = Me
    Call loadSelectionFromValue(value)
    Call updateLinesSelectionStatus
End Function


'[UI Properties]
Public Function setFontSize(value As Single) As UIList
    Set setFontSize = Me
    pFontSize = value
    If pIsRendered Then Call refreshLines
End Function

Public Function setFontFamily(value As String) As UIList
    Set setFontFamily = Me
    pFontFamily = value
    If pIsRendered Then Call refreshLines
End Function


'[Position]
Public Function setTop(value As Single) As UIList
    Set setTop = Me
    pTop = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBottom(value As Single) As UIList
    Set setBottom = Me
    pTop = value - pHeight
    If pIsRendered Then Call arrangeControls
End Function

Public Function setLeft(value As Single) As UIList
    Set setLeft = Me
    pLeft = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setRight(value As Single) As UIList
    Set setRight = Me
    pLeft = value - pWidth
    If pIsRendered Then Call arrangeControls
End Function


'[Size]
Public Function setWidth(value As Single) As UIList
    Set setWidth = Me
    pWidth = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setHeight(value As Single) As UIList
    Set setHeight = Me
    pHeight = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setLineHeight(value As Single) As UIList
    Set setLineHeight = Me
    pLineHeight = value
End Function

Public Function setPadding(value As Single) As UIList
    Set setPadding = Me
    pPadding = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBackColor(value As Long) As UIList
    Set setBackColor = Me
    pBackColor = value
    If pIsRendered Then Call refreshFrameView
End Function

Public Function setSelectedItemBackColor(value As Long) As UIList
    Set setSelectedItemBackColor = Me
    pSelectedItemBackColor = value
    If pIsRendered Then Call refreshFrameView
End Function

Public Function setBorderColor(value As Long) As UIList
    Set setBorderColor = Me
    pBorderColor = value
    If pIsRendered Then Call refreshFrameView
End Function

Public Function setHoverBorderColor(value As Long) As UIList
    Set setHoverBorderColor = Me
    pFocusBorderColor = value
    If pIsRendered Then Call refreshFrameView
End Function

'[Mode]
Public Function setItemsRemovable(value As Boolean) As UIList
    Set setItemsRemovable = Me
    pItemsRemovable = value
End Function

Public Function setItemsEditable(value As Boolean) As UIList
    Set setItemsEditable = Me
    pItemsEditable = value
End Function

Public Function setItemsPreviewable(value As Boolean) As UIList
    Set setItemsPreviewable = Me
    pItemsPreviewable = value
End Function

Public Function setHasMassSelectionButtons(value As Boolean) As UIList
    Set setHasMassSelectionButtons = Me
    pHasMassSelectButtons = value
End Function

Public Function setActionButtonsToLeft(value As Boolean) As UIList
    Set setActionButtonsToLeft = Me
    pActionButtonsToLeft = value
End Function




'[GETTERS]
Public Function isValid() As Boolean
    isValid = UI.Forms.isControlValid(pFrame)
End Function

Public Function getLeft() As Single
    getLeft = pFrame.left
End Function

Public Function getRight() As Single
    If showMassSelectionButtons Then
        getRight = UI.Forms.getRight(pSelectAllButton)
    Else
        getRight = UI.Forms.getRight(pFrame)
    End If
End Function

Public Function getTop() As Single
    If pHasHeaderRow Then
        getTop = pHeadersFrame.top
    Else
        getTop = pFrame.top
    End If
End Function

Public Function getBottom() As Single
    getBottom = UI.Forms.getBottom(pFrame)
End Function

Public Function getWidth() As Single
    getWidth = pFrame.InsideWidth
End Function

Public Function getHeight() As Single
    If pHasHeaderRow Then
        getHeight = UI.Forms.getBottom(pFrame) - pHeadersFrame.top
    Else
        getHeight = pFrame.InsideHeight
    End If
End Function

Friend Function getPadding() As Single
    getPadding = pPadding
End Function

Friend Function getCanvas() As MSForms.UserForm
    Set getCanvas = pFrame
End Function

Friend Function showMassSelectionButtons() As Boolean
    showMassSelectionButtons = (pMultiSelectMode <> fmMultiSelectSingle And pHasMassSelectButtons)
End Function

Friend Function getHeaderFrame() As MSForms.frame
    Set getHeaderFrame = pHeadersFrame
End Function

Friend Function getHeaderFontSize() As Single
    getHeaderFontSize = pHeaderFontSize
End Function

Friend Function getHeaderFontFamily() As String
    getHeaderFontFamily = pHeaderFontFamily
End Function

Friend Function getHeaderFontColor() As Long
    getHeaderFontColor = pHeaderFontColor
End Function

Friend Function getHeaderFontBold() As Long
    getHeaderFontBold = pHeaderFontBold
End Function

Friend Function getBorderColor() As Long
    getBorderColor = pBorderColor
End Function

'[Mode]
Friend Function areItemsRemovable() As Boolean
    areItemsRemovable = pItemsRemovable
End Function

Friend Function areItemsEditable() As Boolean
    areItemsEditable = pItemsEditable
End Function

Friend Function areItemsPreviewable() As Boolean
    areItemsPreviewable = pItemsPreviewable
End Function

Friend Function hasVerticalScrollBar() As Boolean
    hasVerticalScrollBar = pHasVerticalScrollBar
End Function

Friend Function actionButtonsToLeft() As Boolean
    actionButtonsToLeft = pActionButtonsToLeft
End Function



'[CONST ITEMS]
Public Function setConstKeys(keys As Variant) As UIList
    Set setConstKeys = Me
    Set pConstKeysDict = F.dictionaries.createFrom(keys)
End Function

Public Function addConstKey(key As Variant) As UIList
    Set addConstKey = Me
    Call F.dictionaries.addItem(pConstKeysDict, key, key, duplicateBehaviour_Override)
End Function

Public Function isConstKey(key As Variant) As Boolean
    isConstKey = pConstKeysDict.Exists(key)
End Function




'[API]
Public Sub render()
    Call insertHeadersPanel
    Call insertFrame
    Call insertFocusTextBox
    Call insertSelectingLabels
    Call arrangeControls
    Call renderHeaders
    'Call renderLines
    pIsRendered = True
End Sub

Public Function Move(Optional left As Variant, Optional top As Variant) As UIList
    Set Move = Me
    If F.Variables.isOfNumericType(left) Then pLeft = left
    If F.Variables.isOfNumericType(top) Then pTop = top
    Call arrangeControls
End Function

Public Function resize(Optional width As Variant, Optional height As Variant) As UIList
    Set resize = Me
    If F.Math.isNotNegative(width) Then pWidth = width
    If F.Math.isNotNegative(height) Then pHeight = height
    Call arrangeControls
End Function

Public Sub remove()
    Call clear
End Sub





'[RENDERING]
Private Function getLabelName(controlType As ControlTypeEnum) As String
    Dim tag As String
    '----------------------------------------------------------------------------------------------------------
    tag = CLASS_NAME & "|" & pUuid & "|" & controlType
    getLabelName = "ctrl_" & VBA.left$(F.Crypto.SHA1HASH(tag), 25)
End Function

Private Sub insertHeadersPanel()
    If pHasHeaderRow Then
        Set pHeadersFrame = pParent.controls.Add(CREATE_FRAME_ID, getLabelName(ControlType_Frame), True)
        Call refreshHeadersPanel
    End If
End Sub

Private Sub insertFrame()
    Set pFrame = pParent.controls.Add(CREATE_FRAME_ID, getLabelName(ControlType_Frame), True)
    Call refreshFrameView
End Sub

Private Sub insertFocusTextBox()
    Set pFocusTextbox = pFrame.controls.Add(CREATE_TEXTBOX_ID, getLabelName(ControlType_FocusTextBox), True)
    Call refreshFocusTextBox
End Sub

Private Sub insertSelectingLabels()
    Const SELECT_ALL_TOOLTIP As String = "UiList.Buttons.SelectAllTooltip"
    Const UNSELECT_ALL_TOOLTIP As String = "UiList.Buttons.UnselectAllTooltip"
    '----------------------------------------------------------------------------------------------------------
    If showMassSelectionButtons Then
        Set pSelectAllButton = pParent.controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_SelectAll), True)
        Set pDeselectAllButton = pParent.controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_DeselectAll), True)
        Call refreshMassSelectionButton(pSelectAllButton, pSelectAllButtonColor, MsgService.getText(SELECT_ALL_TOOLTIP))
        Call refreshMassSelectionButton(pDeselectAllButton, pDeselectAllButtonColor, MsgService.getText(UNSELECT_ALL_TOOLTIP))
    End If
End Sub

Private Sub refreshHeadersPanel()
    With pHeadersFrame
        .backColor = pBackColor
        .borderStyle = MSForms.fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .caption = vbNullString
    End With
End Sub

Private Sub refreshFrameView()
    With pFrame
        .backColor = pBackColor
        .borderStyle = MSForms.fmBorderStyleSingle
        .borderColor = pBorderColor
        .SpecialEffect = fmSpecialEffectFlat
        .caption = vbNullString
    End With
End Sub

Private Sub refreshFocusTextBox()
    With pFocusTextbox
        .top = 0
        .left = -10
        .height = 1
        .width = 1
    End With
End Sub

Private Sub refreshMassSelectionButton(control As MSForms.label, backColor As Long, tooltip As String)
    With control
        .BackStyle = fmBackStyleOpaque
        .backColor = backColor
        .borderStyle = fmBorderStyleSingle
        .borderColor = &H8000000A
        .ControlTipText = tooltip
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = UI.Icons.getMousePointer
    End With
End Sub

Private Sub refreshLines()
    Stop
    Err.Raise 1
End Sub

Private Sub clear()
    On Error Resume Next
    With pParent.controls
        Call .remove(getLabelName(ControlType_Frame))
    End With
End Sub

Private Sub arrangeControls()
    Call resizeControls
    If pHasHeaderRow Then
        Call UI.Forms.locateControl(pHeadersFrame, pLeft, pTop)
        Call UI.Forms.locateControl(pFrame, pLeft, UI.Forms.getBottom(pHeadersFrame) + pSpaceBetweenHeaderAndList)
    Else
        Call UI.Forms.locateControl(pFrame, pLeft, pTop)
    End If
    If showMassSelectionButtons Then
        Call UI.Forms.locateControl(pSelectAllButton, UI.Forms.getRight(pFrame) + pMassSelectionButtonSize / 2, pFrame.top + pMassSelectionButtonSize / 2)
        Call UI.Forms.locateControl(pDeselectAllButton, pSelectAllButton.left, UI.Forms.getBottom(pSelectAllButton) + pMassSelectionButtonSize / 2)
    End If
End Sub

Private Sub resizeControls()
    If pHasHeaderRow Then
        Call UI.Forms.resizeControl(pHeadersFrame, pWidth, pHeaderHeight)
        Call UI.Forms.resizeControl(pFrame, pWidth, pHeight - pHeaderHeight - pSpaceBetweenHeaderAndList)
    Else
        Call UI.Forms.resizeControl(pFrame, pWidth, pFrame.InsideHeight)
    End If
    Call UI.Forms.resizeControl(pFocusTextbox, 1, pFrame.ScrollHeight)
    If showMassSelectionButtons Then
        Call UI.Forms.resizeControl(pSelectAllButton, pMassSelectionButtonSize, pMassSelectionButtonSize)
        Call UI.Forms.resizeControl(pDeselectAllButton, pMassSelectionButtonSize, pMassSelectionButtonSize)
    End If
End Sub

Private Sub renderLines()
    Dim i As Long
    Dim linesCounter As Long
    Dim currentLinesCounter As Long
    Dim startIndex As Long
    '------------------------------------------------------------------------------------------------------
    Dim line As UIListLine
    '------------------------------------------------------------------------------------------------------
    
    linesCounter = F.Math.getMin(pItems.Count, pMaxItemsPerPage)
    If linesCounter Then
        currentLinesCounter = F.arrays.getSize(pLines)
        
        If currentLinesCounter = 0 Then
            ReDim pLines(1 To linesCounter)
            startIndex = 1
        ElseIf currentLinesCounter < linesCounter Then
            ReDim Preserve pLines(1 To linesCounter)
            startIndex = currentLinesCounter + 1
        End If
    
        If startIndex Then
            For i = startIndex To linesCounter
                Set line = renderLine(i)
                Set pLines(i) = line
            Next i
            Call arrangeLines
        End If
    End If
    
End Sub

Private Function renderLine(index As Long) As UIListLine
    Set renderLine = New UIListLine
    With renderLine
        Call .setParent(Me)
        Call .setIndex(index)
        Call .setHeight(pLineHeight)
        Call .render
    End With
End Function

Private Sub resetLinesContainer()
    Dim line As UIListLine
    '------------------------------------------------------------------------------------------------------
    
    For Each line In pVisibleLinesList
        Call line.destroy
    Next line
    
    Erase pLines
    Set pLinesByKey = F.dictionaries.Create(False)
    Set pKeysByLine = F.dictionaries.Create(False)
    Set pSelection = F.dictionaries.Create(False)
    Set pActiveLine = Nothing
    
End Sub

Private Sub arrangeLines()
    Dim line As UIListLine
    Dim Y As Single:                                Y = 0.25
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    For i = LBound(pLines) To UBound(pLines)
        Set line = pLines(i)
        With line
            If line.isVisible Then
                Call .Move(top:=Y)
                Y = .getBottom
            End If
        End With
    Next i

End Sub








'[ITEMS]
Private Sub sortItems(Optional ascending As Boolean = True)
    Dim constItems As VBA.Collection
    Dim sortableItems As Scripting.Dictionary
    '----------------------------------------------------------------------------------------------------------
    Dim key As Variant
    '----------------------------------------------------------------------------------------------------------
    
    Set constItems = New VBA.Collection
    Set sortableItems = F.dictionaries.Create(False)
    
    '[Divide items into two groups - const and sortable]
    For Each key In pFiltered.keys
        If Not VBA.IsEmpty(key) Then
            If isConstKey(key) Then
                Call constItems.Add(key)
            Else
                Call F.dictionaries.addItem(sortableItems, getItemCaption(key), key)
            End If
        End If
    Next key
    
    Set pSortedKeys = F.Collections.Join(constItems, _
                    F.dictionaries.toCollection(F.dictionaries.sortByKey(sortableItems, ascending), DictPart_ValueOnly))
                    
End Sub

Private Sub updatePagination(Optional page As Long)
    If Not pHeaderLine Is Nothing Then
        Call pHeaderLine.updatePaginationStatus(page, countPages)
    End If
End Sub

Private Sub updateLinesItems(Optional pageIndex As Long = 1)
    Dim i As Long
    Dim index As Long
    Dim pagesOffset As Long
    Dim lastIndex As Long
    '------------------------------------------------------------------------------------------------------
    Dim key As Variant
    Dim line As UIListLine
    Dim lastVisibleLineIndex As Long
    '------------------------------------------------------------------------------------------------------
    
    Set pLinesByKey = F.dictionaries.Create(False)
    Set pKeysByLine = F.dictionaries.Create(False)
    
    pCurrentPage = pageIndex
    If Not pHeaderLine Is Nothing Then Call pHeaderLine.setCurrentPage(pCurrentPage)
    
    pagesOffset = (pCurrentPage - 1) * pMaxItemsPerPage
    lastIndex = F.Math.getMin(F.arrays.getSize(pLines), pMaxItemsPerPage)
    
    For i = 1 To lastIndex
        Set line = pLines(i)
        index = i + pagesOffset
        If index <= pFiltered.Count Then
            key = pSortedKeys(index)
            With line
                Call .setVisibility(True)
                Call .setItem(pFiltered.item(key))
                Call .setKey(key)
                Call .setCaption(getItemCaption(key))
                Call .setSelection(pSelection.Exists(key))
                Call .setSpecialFormat(F.dictionaries.getObject(pSpecialFormats, key))
                Call .setPreviewMode(F.Variables.equals(key, pPreviewItemKey))
            End With
            Call F.dictionaries.addItem(pLinesByKey, key, line)
            Call F.dictionaries.addItem(pKeysByLine, line.getUuid, key)
            lastVisibleLineIndex = i
        Else
            With line
                Call .clear
            End With
        End If
    Next i
    
    Call updateLinesSelectionStatus
    Call updateActiveLineAfterPageChange(lastVisibleLineIndex)
    Call showHideScrollBar(lastVisibleLineIndex)
    
End Sub

Private Sub updateActiveLineAfterPageChange(lastVisibleIndex As Long)
    Dim line As UIListLine
    '------------------------------------------------------------------------------------------------------
    If Not pActiveLine Is Nothing Then
        If pActiveLine.getIndex > lastVisibleIndex Then
            If lastVisibleIndex Then
                Set line = pLines(lastVisibleIndex)
                Call updateActiveLine(line)
            Else
                Call resetActiveLine
            End If
        End If
    End If
End Sub

Private Function getItemCaption(key As Variant) As String
    If pCaptions.Exists(key) Then
        getItemCaption = pCaptions.item(key)
    Else
        getItemCaption = key
    End If
End Function

Private Function countPages() As Long
    countPages = VBA.Int(pFiltered.Count / pMaxItemsPerPage) + VBA.IIf(pFiltered.Count Mod pMaxItemsPerPage > 0, 1, 0)
End Function

Private Sub showHideScrollBar(lastIndex As Long)
    Dim lastLine As UIListLine
    Dim bottom As Single
    '----------------------------------------------------------------------------------------------------------
    
    pHasVerticalScrollBar = False
    If F.arrays.isInBounds(lastIndex, 1, pLines) Then Set lastLine = pLines(lastIndex)
    If Not lastLine Is Nothing Then
        bottom = lastLine.getBottom
        If bottom > pFrame.height Then pHasVerticalScrollBar = True
    End If
    
    With pFrame
        If pHasVerticalScrollBar Then
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = bottom
        Else
            .ScrollHeight = 0
            .ScrollBars = fmScrollBarsNone
        End If
    End With
    
End Sub

Public Sub addItem(key As Variant, item As Variant, Optional caption As Variant)
    Dim line As UIListLine
    '------------------------------------------------------------------------------------------------------
    If Not VBA.IsMissing(caption) Then Call F.dictionaries.addItem(pCaptions, key, caption)
    Call F.dictionaries.addItem(pItems, key, item)
    Call filterItems(pFilter, False)
    Call sortItems
    Call updateLinesItems(pCurrentPage)
End Sub

Public Sub changeItemCaption(key As Variant, caption As String)
    Dim line As UIListLine
    '----------------------------------------------------------------------------------------------------------
    Call F.dictionaries.addItem(pCaptions, key, caption, duplicateBehaviour_Override)
    Set line = getLineByKey(key)
    If Not line Is Nothing Then Call line.setCaption(caption)
End Sub




'[HEADERS]
Public Sub renderHeaders()
    Dim headerDefinition As DListHeader
    Dim totalDeclaredWidth As Single
    '----------------------------------------------------------------------------------------------------------
    
    If pHasHeaderRow And Not pHeaderDefinitions Is Nothing Then
        totalDeclaredWidth = getAllHeadersWidth
        For Each headerDefinition In pHeaderDefinitions
            Set pHeaderLine = UI.Create.ListHeader(Me, headerDefinition)
            With pHeaderLine
                Call .setWidthShare(headerDefinition.getWidth / totalDeclaredWidth)
                Call .setTotalPages(countPages)
                Call .render
            End With
            Call pHeaderDefinitions.Add(pHeaderLine)
        Next headerDefinition
    End If
    
End Sub

Private Function getAllHeadersWidth() As Single
    Dim headerDef As DListHeader
    '----------------------------------------------------------------------------------------------------------
    For Each headerDef In pHeaderDefinitions
        getAllHeadersWidth = getAllHeadersWidth + headerDef.getWidth
    Next headerDef
End Function

Public Sub changeHeaderCaption(index As Long, caption As String)
    Dim headerDef As DListHeader
    '----------------------------------------------------------------------------------------------------------
    Set headerDef = F.Collections.getObjectByIndex(pHeaderDefinitions, index)
    If Not headerDef Is Nothing Then
        Call headerDef.setCaption(caption)
    End If
End Sub





'[SELECTING ITEMS]
Private Function countVisibleLines() As Long
    Dim line As UIListLine
    Dim i As Long
    '----------------------------------------------------------------------------------------------------------
    If F.arrays.isDefined(pLines) Then
        For i = LBound(pLines) To UBound(pLines)
            Set line = pLines(i)
            If Not line.isVisible Then
                Exit For
            Else
                countVisibleLines = i
            End If
        Next i
    Else
        countVisibleLines = 0
    End If

End Function

Friend Function getLineByKey(key As Variant) As UIListLine
    Set getLineByKey = F.dictionaries.getObject(pLinesByKey, key)
End Function

Friend Function getLineByIndex(index As Long) As UIListLine
    If F.arrays.isInBounds(index, 1, pLines) Then
        Set getLineByIndex = pLines(index)
    End If
End Function

Friend Sub triggerLineClick(line As UIListLine)
    Call setFocus
    Call updateActiveLine(line)
    Call changeSelection(line)
End Sub

Friend Sub triggerLineDoubleClick(line As UIListLine)
    Call triggerLineClick(line)
    RaiseEvent AfterDoubleClick
End Sub

Public Sub resetActiveLine()
    Call updateActiveLine(Nothing)
End Sub

Public Sub updateActiveItemByIndex(index As Long)
    Dim line As UIListLine
    '--------------------------------------------------------------------------------------------------
    Set line = getLineByIndex(index)
    If Not line Is Nothing Then
        Call updateActiveLine(line)
    End If
End Sub

Friend Sub updateActiveLine(line As UIListLine)
    If Not pActiveLine Is Nothing Then
        Call pActiveLine.setActiveState(False)
    End If
    Set pActiveLine = line
    If Not pActiveLine Is Nothing Then
        Call pActiveLine.setActiveState(True)
    End If
End Sub

Private Sub changeActiveRow(change As Integer)
    Dim visibleLines As Long
    Dim currentActiveIndex As Long
    Dim index As Long
    Dim line As UIListLine
    '--------------------------------------------------------------------------------------------------
    
    visibleLines = countVisibleLines
    
    If visibleLines Then
        
        If Not pActiveLine Is Nothing Then
            currentActiveIndex = pActiveLine.getIndex
        End If
        
        index = F.Math.getMax(1, F.Math.getMin(visibleLines, currentActiveIndex + change))
        Set line = pLines(index)
        Call updateActiveLine(line)
        
        If getItemVisibility(line) < 1 Then
            Call scrollToIndex(line, VBA.IIf(change < 0, VerticalAlignment_Top, VerticalAlignment_Bottom))
        End If
        
        If currentActiveIndex + change = 0 Then
            RaiseEvent GiveFocusToTopControl
        End If
        
    End If
    
End Sub

Friend Sub changeSelection(line As UIListLine)
    If Not line Is Nothing Then
        If pMultiSelectMode = fmMultiSelectSingle Then
            Call changeSelectionInSingleMode(line)
        ElseIf pMultiSelectMode = fmMultiSelectExtended Then
            'Err.Raise 1
        ElseIf pMultiSelectMode = fmMultiSelectMulti Then
            Call changeSelectionInMultiMode(line, Not line.isSelected)
        End If
        RaiseEvent AfterChange
    End If
End Sub

Private Sub changeSelectionInSingleMode(line As UIListLine)
    Dim previousSelectedKey As Variant
    Dim previousSelectedLine As UIListLine
    '--------------------------------------------------------------------------------------------------
    If Not line Is Nothing Then
        Call F.Variables.assign(previousSelectedKey, getSelectedKey)
        Set previousSelectedLine = F.dictionaries.getObject(pLinesByKey, previousSelectedKey)
        If Not previousSelectedLine Is Nothing Then Call previousSelectedLine.setSelection(False)
        Set pSelection = F.dictionaries.createSingleEntryDict(line.getKey, line.getItem)
        Call line.setSelection(True)
    End If
End Sub

Private Sub changeSelectionInMultiMode(line As UIListLine, selected As Boolean)
    If selected Then
        Call line.setSelection(True)
        Call F.dictionaries.addItem(pSelection, line.getKey, line.getItem)
    Else
        Call line.setSelection(False)
        Call F.dictionaries.removeItem(pSelection, line.getKey)
    End If
End Sub

Private Sub selectAll()
    Set pSelection = F.dictionaries.Clone(pFiltered)
    Call updateLinesSelectionStatus
    RaiseEvent AfterChange
End Sub

Private Sub deselectAll()
    Set pSelection = F.dictionaries.Create(False)
    Call updateLinesSelectionStatus
    RaiseEvent AfterChange
End Sub

Private Sub updateLinesSelectionStatus()
    Dim line As UIListLine
    Dim i As Long
    Dim key As Variant
    '--------------------------------------------------------------------------------------------------
    If F.arrays.isDefined(pLines) Then
        For i = LBound(pLines) To UBound(pLines)
            Set line = pLines(i)
            If Not line.isVisible Then
                'skip
            Else
                Call line.setSelection(pSelection.Exists(line.getKey))
            End If
        Next i
    End If
End Sub

Public Sub selectItemByKey(key As Variant)
    Dim line As UIListLine
    '--------------------------------------------------------------------------------------------------
    Set line = F.dictionaries.getObject(pLinesByKey, key)
    If Not line Is Nothing Then
        If pMultiSelectMode = fmMultiSelectSingle Then
            Call changeSelectionInSingleMode(line)
        ElseIf pMultiSelectMode = fmMultiSelectExtended Then
            Call changeSelectionInMultiMode(line, True)
        ElseIf pMultiSelectMode = fmMultiSelectMulti Then
            Call changeSelectionInMultiMode(line, True)
        End If
    End If
End Sub

Public Sub resetSelection(Optional resetActiveLine As Boolean = True)
    Call deselectAll
    If resetActiveLine Then Call updateActiveLine(Nothing)
End Sub



'[PREVIEW]
Friend Sub setPreviewLine(line As UIListLine)
    Dim previousPreviewLine As UIListLine
    '--------------------------------------------------------------------------------------------------
    If F.Variables.equals(line.getKey, pPreviewItemKey) Then
        'Nothing changed
    Else
    
        'Unselect previous preview line
        If Not F.Variables.isUndefined(pPreviewItemKey) Then
            Set previousPreviewLine = F.dictionaries.getObject(pLinesByKey, pPreviewItemKey)
            If Not previousPreviewLine Is Nothing Then
                Call previousPreviewLine.setPreviewMode(False)
            End If
        End If
                    
        'Select new preview line
        Call F.Variables.assign(pPreviewItemKey, line.getKey)
        Call line.setPreviewMode(True)
        
        RaiseEvent AfterPreviewChanged(getKeyForLine(line), line.getItem, line.getIndex)
    End If
End Sub

Private Function getKeyForLine(line As UIListLine) As Variant
    If pKeysByLine.Exists(line.getUuid) Then
        Call F.Variables.assign(getKeyForLine, pKeysByLine.item(line.getUuid))
    End If
End Function



'[FILTERING]
Public Sub filterItems(filter As String, Optional rerender As Boolean = True)
    Dim key As Variant
    Dim line As UIListLine
    Dim caption As String
    '--------------------------------------------------------------------------------------------------
    pFilter = filter
    If VBA.Len(pFilter) Then
        Set pFiltered = F.dictionaries.Create(False)
        For Each key In pItems.keys
            caption = getItemCaption(key)
            If VBA.InStr(1, caption, filter, vbTextCompare) > 0 Then
                Call F.dictionaries.addItem(pFiltered, key, pItems.item(key))
            End If
        Next key
    Else
        Set pFiltered = F.dictionaries.Clone(pItems)
    End If
    If rerender Then
        Call sortItems
        Call updateLinesItems
    End If
End Sub


'[CONTROL EVENTS]
Public Sub setFocus(Optional selectFirstItem As Boolean = False, _
                    Optional activateFirstItem As Boolean = False, _
                    Optional previewFirstItem As Boolean = False)
    Dim line As UIListLine
    '--------------------------------------------------------------------------------------------------
    If UI.Forms.isControlValid(pFocusTextbox) Then
        On Error Resume Next
        Call pFocusTextbox.setFocus
    End If
    pFrame.borderColor = pFocusBorderColor
    If pFiltered.Count > 0 Then
        Set line = getFirstVisibleLine
        If selectFirstItem Then Call changeSelection(line)
        If activateFirstItem Then Call updateActiveLine(line)
        If previewFirstItem Then Call setPreviewLine(line)
    End If
End Sub

Private Function getFirstVisibleLine() As UIListLine
    Dim i As Long
    Dim line As UIListLine
    '--------------------------------------------------------------------------------------------------
    If F.arrays.isDefined(pLines) Then
        For i = LBound(pLines) To UBound(pLines)
            Set line = pLines(i)
            If line.isVisible Then
                Set getFirstVisibleLine = line
                Exit For
            End If
        Next i
    End If
End Function

Public Sub loseFocus()
    pFrame.borderColor = pBorderColor
End Sub

Private Sub pSelectAllButton_Click()
    Call selectAll
End Sub

Private Sub pDeselectAllButton_Click()
    Call deselectAll
End Sub

Friend Sub editItem(line As UIListLine)
    Dim newCaption As String
    '--------------------------------------------------------------------------------------------------
    RaiseEvent AfterEditItemClick(line.getItem, newCaption)
    Call line.setCaption(newCaption)
End Sub

Friend Sub removeItem(line As UIListLine)
    Dim confirmed As Boolean
    '--------------------------------------------------------------------------------------------------
    
    Stop
    
    RaiseEvent AfterRemoveItemClick(line.getItem, confirmed)
    If confirmed Then
        Stop
'
'        '[Deselect]
'        If line.isSelected Then
'            Call changeSelection(line)
'            If F.Variables.equals(line, pSelectedLine) Then Set pSelectedLine = Nothing
'            RaiseEvent AfterChange
'        End If
'
'        Call line.destroy
'        Stop 'xxx
'        'Call F.dictionaries.removeItem(pItems, line.getKey)
'        'Call F.dictionaries.removeItem(pLinesByKey, line.getKey)
'        Call arrangeLines
    End If
End Sub


Public Sub removeAllItems()
    'Call UI.Forms.clearMsFrame(pFrame)
    Set pItems = F.dictionaries.Create(False)
    Set pFiltered = F.dictionaries.Create(False)
    Set pLinesByKey = F.dictionaries.Create(False)
    Set pKeysByLine = F.dictionaries.Create(False)
    Set pConstKeysDict = F.dictionaries.Create(False)
    Set pSelection = F.dictionaries.Create(False)
    Set pActiveLine = Nothing
    pPreviewItemKey = Empty
End Sub








'[SELECTION]
Public Function getSelectedValue() As Variant
    Dim selection As VBA.Collection
    '----------------------------------------------------------------------------------------------------------
    Set selection = F.dictionaries.toCollection(pSelection, DictPart_ValueOnly)
    If pMultiSelectMode = fmMultiSelectSingle Then
        If F.Collections.isNonEmptyCollection(selection) Then
            Call F.Variables.assign(getSelectedValue, selection.item(1))
        End If
    Else
        Set getSelectedValue = selection
    End If
End Function

Public Function getSelectedKey() As Variant
    Dim selection As VBA.Collection
    '----------------------------------------------------------------------------------------------------------
    Set selection = F.dictionaries.toCollection(pSelection, DictPart_KeyOnly)
    If pMultiSelectMode = fmMultiSelectSingle Then
        If F.Collections.isNonEmptyCollection(selection) Then
            Call F.Variables.assign(getSelectedKey, selection.item(1))
        End If
    Else
        Set getSelectedKey = selection
    End If
End Function

Public Function getSelectedKeyValue() As Scripting.Dictionary
    Set getSelectedKeyValue = pSelection
End Function

Public Function getSelectedIndex() As Variant
    Dim col As VBA.Collection
    Dim key As Variant
    Dim line As UIListLine
    '----------------------------------------------------------------------------------------------------------
    
    Set col = New VBA.Collection
    For Each key In pSelection.keys
        Set line = F.dictionaries.getObject(pLinesByKey, key)
        If Not line Is Nothing Then
            Call col.Add(line.getIndex)
        End If
    Next key
    
    If pMultiSelectMode = fmMultiSelectSingle Then
        If F.Collections.isNonEmptyCollection(col) Then
            getSelectedIndex = col.item(1)
        End If
    Else
        Set getSelectedIndex = col
    End If
End Function

Public Function isAnyItemSelected() As Boolean
    isAnyItemSelected = F.dictionaries.isNonEmptyDictionary(pSelection)
End Function

Public Function getKeyForValue(value As Variant) As Variant
    getKeyForValue = F.dictionaries.getKeyByItem(pItems, value)
End Function

Public Function getValueForKey(key As Variant) As Variant
    If pItems.Exists(key) Then
        Call F.Variables.assign(getValueForKey, pItems.item(key))
    End If
End Function

Public Function getValueByIndex(index As Long) As Variant
    Dim line As UIListLine
    '----------------------------------------------------------------------------------------------------------
    Set line = getLineByIndex(index)
    If Not line Is Nothing Then
        Call F.Variables.assign(getValueByIndex, line.getItem)
    End If
End Function













'[NAVIGATION]
Friend Sub changePage(pageIndex As Long)
    If F.Math.isInRange(pageIndex, 1, countPages) Then
        Call updateLinesItems(pageIndex)
    End If
End Sub


'[KEY EVENTS]
Private Sub pFocusTextbox_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    Select Case keyCode
        '[Navigation]
        Case vbKeyUp                    'Up
            Call changeActiveRow(-1)
            keyCode = 0
        Case vbKeyPageUp                'PageUp
            Call changeActiveRow(-getVisibleRowsCounter + 1)
            keyCode = 0
        Case vbKeyDown                  'Down
            Call changeActiveRow(1)
            keyCode = 0
        Case vbKeyLeft                  'Left
            Call changePage(pCurrentPage - 1)
            keyCode = 0
        Case vbKeyRight                 'Right
            Call changePage(pCurrentPage + 1)
            keyCode = 0
        Case vbKeyPageDown              'PageDown
            Call changeActiveRow(getVisibleRowsCounter - 1)
            keyCode = 0
        Case vbKeyHome                  'Home
            Call changeActiveRow(-pFiltered.Count)
            keyCode = 0
        Case vbKeyEnd                   'End
            Call changeActiveRow(pFiltered.Count)
            keyCode = 0
        Case vbKeySpace
            Call changeSelection(pActiveLine)
        Case vbKeyReturn
            Call changeSelection(pActiveLine)
            RaiseEvent AfterDoubleClick
        Case vbKeyTab
            Debug.Print 1
        
'        Case vbKeyLeft
'            Call mouseScroll(-SINGLE_SCROLL_PIXELS)
'        Case vbKeyRight
'            Call mouseScroll(SINGLE_SCROLL_PIXELS)
'            KeyCode = 0
'        Case 65, 97     'CTRL + A, a
'            If Shift = 2 Then
'                'If areAllSelected Then Call deselectAll Else Call SelectAll
'            End If
'        Case 70, 102    'CTRL + F, f
'            If Shift = 2 Then
'                'RaiseEvent FindWindowActivated
'            End If
    End Select
End Sub

Private Function getVisibleRowsCounter() As Integer
    getVisibleRowsCounter = VBA.Int(pFrame.InsideHeight / pLineHeight)
End Function




'Results:
'1  Item visible
'0  Item not visible
' -1 < x < 0    Item partially visible (on the top of the list)
' 0 < x < 1     Item partially visible (on the bottom of the list)
Private Function getItemVisibility(item As UIListLine) As Single
    Dim index As Integer
    Dim itemTop As Single
    Dim itemBottom As Single
    Dim listVisibleTop As Single
    Dim listVisibleBottom As Single
    '--------------------------------------------------------------------------------------------------
    
    If Not item.isVisible Then
        getItemVisibility = 0
    Else
        With item
            itemTop = .getTop
            itemBottom = .getBottom
        End With
        With pFrame
            listVisibleTop = .scrollTop
            listVisibleBottom = .scrollTop + .InsideHeight
        End With
    End If
    
    If itemTop < listVisibleTop Then
        If itemBottom > listVisibleTop Then
            getItemVisibility = (itemBottom - listVisibleTop) / -pLineHeight
        End If
    ElseIf itemTop < listVisibleBottom Then
        If itemBottom < listVisibleBottom Then
            getItemVisibility = 1
        Else
            getItemVisibility = (listVisibleBottom - itemTop) / pLineHeight
        End If
    End If
    
End Function

Private Sub scrollToIndex(line As UIListLine, position As VerticalAlignmentEnum)
    Dim totalScroll As Single
    Dim lineTop As Single
    Dim scrollTop As Single
    '------------------------------------------------------------------------------------------------------
    
    totalScroll = pFrame.ScrollHeight
    lineTop = line.getTop
    
    Select Case position
        Case VerticalAlignment_Top:        scrollTop = lineTop
        Case VerticalAlignment_Middle:     scrollTop = lineTop + pLineHeight / 2 - pFrame.height / 2
        Case VerticalAlignment_Bottom:     scrollTop = lineTop + pLineHeight - pFrame.height
    End Select
    
    
    'Check if new scroll top is in range of scroll height.
    If scrollTop < 0 Then
        scrollTop = 0
    ElseIf scrollTop > totalScroll - pFrame.height Then
        scrollTop = totalScroll - pFrame.height
    End If
    
    On Error Resume Next
    pFrame.scrollTop = scrollTop

End Sub



Private Sub pFrame_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    pFocusTextbox.top = pFrame.scrollTop + ActualDy
End Sub






'[Implementation of UIControl]
Private Sub IControl_setVisibility(value As Boolean)
    Stop
    'Call setVisibility(value)
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
    Call resize(width, height)
End Sub

Private Sub IControl_alignVertically(baseCtrl As MSForms.IControl)
    Stop
    'Call alignVertically(baseCtrl)
End Sub

Private Function IControl_move(Optional x As Variant, Optional Y As Variant) As IControl
    Set IControl_move = Move(x, Y)
End Function

Private Function IControl_offset(Optional x As Single, Optional Y As Single) As IControl
    Stop
    'Set IControl_offset = offset(x, y)
End Function

Private Sub IControl_triggerEvent(eventId As Long)
    Stop
End Sub

