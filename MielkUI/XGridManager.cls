VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XGridManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IGrid
Implements IFilterable

Private Const CLASS_NAME As String = "XGridManager"
'----------------------------------------------------------------------------------------------------------
Private Const DELETE_ITEM_TITLE_TAG As String = "ConfirmDeleteItem_WindowTitle"
Private Const DELETE_ITEM_QUESTION_TAG As String = "ConfirmDeleteItem_Question"
Private Const DELETE_ITEM_SUCCESS_TAG As String = "DeleteItem_Success"
Private Const RECORDS_COUNTER_TAG As String = "Caption_recordsCounter"
Private Const TOOLTIP_EXPORT_TO_EXCEL As String = "Tooltip_ExportToExcel"
Private Const TOOLTIP_DEACTIVATE_ALL_FILTERS As String = "Tooltip_DeactivateAllFilters"
'[Control names] ------------------------------------------------------------------------------------------
Private Const HEADER_FRAME_NAME As String = "frmHeader"
Private Const RECORDS_FRAME_NAME As String = "frmRecords"
Private Const FOOTER_FRAME_NAME As String = "frmFooter"
Private Const FOOTER_BORDER_NAME As String = "lblFooterTopBorder"
Private Const SEARCH_ICON_NAME As String = "lblSearchIcon"
Private Const SEARCH_TEXTBOX_NAME As String = "txtSearch"
Private Const SEARCH_CANCEL_BUTTON_NAME As String = "lblSearchCancelButton"
Private Const CHANGE_FILTERS_STATE_BUTTON_NAME As String = "lblChangeFiltersStateButton"
Private Const EXPORT_TO_EXCEL_BUTTON_NAME As String = "lblExcelIcon"
Private Const DEACTIVATE_ALL_FILTERS_BUTTON_NAME As String = "lblDeactivateAllFilters"
'----------------------------------------------------------------------------------------------------------
Private Const HEADER_LINE_HEIGHT As Single = 24
Private Const HEADER_LINE_BACK_COLOR As Long = 16119285
Private Const COLUMNS_SPACE As Single = 6
Private Const LINES_SPACE As Single = 0 '3
Private Const VERTICAL_PADDING As Single = 3
Private Const HORIZONTAL_PADDING As Single = 6
Private Const BACK_COLOR As Long = vbWhite
Private Const DEFAULT_SORT_ASCENDING As Boolean = True
Private Const PROGRESS_BAR_UPDATE_STEP As Long = 10
'----------------------------------------------------------------------------------------------------------
Private pWindow As IGridParentWindow
Private pFrame As MSForms.frame
Private pHeadersContainer As MSForms.frame
Private pRecordsContainer As MSForms.frame
Private pFooterContainer As MSForms.frame
Private pColumns As Scripting.Dictionary
Private pHeaderLine As xxGridHeaderLine
'[State] --------------------------------------------------------------------------------------------------
Private pIsRendered As Boolean
'----------------------------------------------------------------------------------------------------------
Private pItems As VBA.Collection
Private pLines As Scripting.Dictionary
'[Sorting] ------------------------------------------------------------------------------------------------
Private pSortColumn As DColumn
Private pSortAscending As Boolean
'[Filtering] ----------------------------------------------------------------------------------------------
Private pFiltered As Scripting.Dictionary
Private pFilters As Scripting.Dictionary
Private pGlobalFilter As String
'[Captions] -----------------------------------------------------------------------------------------------
Private pDeleteItemTitle As String
Private pDeleteItemQuestion As String
Private pDeleteSuccessInfo As String
'[Footer panel] -------------------------------------------------------------------------------------------
Private Const FOOTER_PANEL_HEIGHT As Single = 24
Private Const FOOTER_PANEL_BACK_COLOR As Long = 16119285
Private Const FOOTER_PANEL_BORDER_COLOR As Long = 5526612
Private pFooterTopBorder As MSForms.label
'[Search box] ---------------------------------------------------------------------------------------------
Private Const SEARCH_ICON_SIZE As Single = 12
Private Const SEARCH_CANCEL_BUTTON_SIZE As Single = 12
Private Const SEARCH_TEXTBOX_HEIGHT As Single = 18
Private Const SEARCH_TEXTBOX_WIDTH As Single = 120

Private Const SEARCH_TEXTBOX_FONT_SIZE As Single = 10
Private Const SEARCH_TEXTBOX_FONT_NAME As String = "Tahoma"
Private Const SEARCH_CONTROLS_SPACE As Single = 3
Private Const SEARCH_TEXTBOX_BORDER_COLOR As Long = 11842740
Private Const DEACTIVATE_FILTERS_BUTTON_RIGHT_MARGIN As Single = 18
Private pSearchIcon As MSForms.label
Private WithEvents pSearchBox As MSForms.textBox
Attribute pSearchBox.VB_VarHelpID = -1
Private WithEvents pSearchCancelButton As MSForms.label
Attribute pSearchCancelButton.VB_VarHelpID = -1
Private pSearchText As String
Private WithEvents pChangeAllFiltersState As MSForms.label
Attribute pChangeAllFiltersState.VB_VarHelpID = -1
Private pFiltersDeactivated As Boolean
'[Records counter]
Private Const FOOTER_CONTAINER_LEFT_PADDING As Single = 6
Private Const RECORDS_COUNTER_RIGHT_MARGIN As Single = 18
Private Const RECORDS_COUNTER_WIDTH As Single = 100
Private Const RECORDS_COUNTER_FONT_SIZE As Single = 8
Private Const RECORDS_COUNTER_FONT_NAME As String = APP_FONT_FAMILY
Private pRecordsCounterLabel As MSForms.label
Private pHasRecordsCounter As Boolean
'[Export to Excel]
Private Const EXCEL_EXPORT_ICON_SIZE As Single = 12
Private WithEvents pExportToExcelButton As MSForms.label
Attribute pExportToExcelButton.VB_VarHelpID = -1
'----------------------------------------------------------------------------------------------------------
Event action(actionName As String, actionParams As Scripting.Dictionary)
Event AfterItemEdit(item As IGridItem, ByRef updateAfterward As Boolean)
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Initialize()
    Call setDefaultValues
End Sub

Private Sub setDefaultValues()
    '[Captions]
    pDeleteItemTitle = MsgService.getText(DELETE_ITEM_TITLE_TAG)
    pDeleteItemQuestion = MsgService.getText(DELETE_ITEM_QUESTION_TAG)
    pDeleteSuccessInfo = MsgService.getText(DELETE_ITEM_SUCCESS_TAG)
End Sub

Private Sub Class_Terminate()
    Call destroy
End Sub




'[SETTERS]
Public Function setWindow(value As IGridParentWindow) As XGridManager
    Set setWindow = Me
    Set pWindow = value
End Function

Public Function setFrame(value As MSForms.frame) As XGridManager
    Set setFrame = Me
    Set pFrame = value
End Function

Public Function setItems(items As VBA.Collection) As XGridManager
    Set setItems = Me
    Set pItems = items
End Function

Public Function setColumns(columns As Scripting.Dictionary) As XGridManager
    Set setColumns = Me
    Set pColumns = columns
End Function

Public Function setFilters(filters As Scripting.Dictionary) As XGridManager
    Set setFilters = Me
    Set pFilters = filters
End Function

Public Function setSortParameters(sortColumnKey As String, sortAscending As Boolean) As XGridManager
    Set setSortParameters = Me
    Set pSortColumn = F.dictionaries.getObject(pColumns, sortColumnKey)
    pSortAscending = sortAscending
End Function

'[Captions]
Public Function setDeleteItemWindowTitle(value As String) As XGridManager
    Set setDeleteItemWindowTitle = Me
    pDeleteItemTitle = value
End Function

Public Function setDeleteItemQuestion(value As String) As XGridManager
    Set setDeleteItemQuestion = Me
    pDeleteItemQuestion = value
End Function

Public Function setDeleteSuccessInfo(value As String) As XGridManager
    Set setDeleteSuccessInfo = Me
    pDeleteSuccessInfo = value
End Function

Public Function setHasRecordsCounter(value As Boolean) As XGridManager
    Set setHasRecordsCounter = Me
    pHasRecordsCounter = value
End Function




'[GETTERS]
Public Function getFrame() As MSForms.frame
    Set getFrame = pFrame
End Function

Public Function getColumnsCollection() As VBA.Collection
    Set getColumnsCollection = F.dictionaries.toCollection(pColumns)
End Function

Public Function getHorizontalPadding() As Single
    getHorizontalPadding = HORIZONTAL_PADDING
End Function

Public Function getColumnsSpace() As Single
    getColumnsSpace = COLUMNS_SPACE
End Function

Public Function getColumn(name As String) As DColumn
    Set getColumn = F.dictionaries.getObject(pColumns, name)
End Function

Public Function getFilters() As Scripting.Dictionary
    Set getFilters = pFilters
End Function





'[LOADING]
Public Sub load()
    Call bindFiltersToColumns
    Call updateFilteredCollection
    Call render
End Sub




'[RENDERING]
Private Sub render()
    If Not pIsRendered Then
        Call insertContainers
        Call applyColumnsLeftPosition
        Call renderHeaderLine
        Call renderItemLines
        Call renderFooter
        Call arrangeGrid
        Call insertVerticalLines
        Call arrangeComponents
        pIsRendered = True
    End If
End Sub

Private Sub insertContainers()
    If Not UI.controls.isValid(pHeadersContainer) Then
        Set pHeadersContainer = pFrame.controls.add(CREATE_FRAME_ID, HEADER_FRAME_NAME)
        With pHeadersContainer
            .height = HEADER_LINE_HEIGHT
            .width = pFrame.InsideWidth
            .top = 0
            .left = 0
            .SpecialEffect = fmSpecialEffectFlat
            .backColor = HEADER_LINE_BACK_COLOR
            .borderStyle = fmBorderStyleNone
        End With
    End If
    
    If Not UI.controls.isValid(pRecordsContainer) Then
        Set pRecordsContainer = pFrame.controls.add(CREATE_FRAME_ID, RECORDS_FRAME_NAME)
        With pRecordsContainer
            .height = pFrame.InsideHeight - HEADER_LINE_HEIGHT - FOOTER_PANEL_HEIGHT
            .width = pFrame.InsideWidth
            .top = HEADER_LINE_HEIGHT
            .left = 0
            .SpecialEffect = fmSpecialEffectFlat
            .borderStyle = fmBorderStyleNone
        End With
    End If
    
    If Not UI.controls.isValid(pFooterContainer) Then
        Set pFooterContainer = pFrame.controls.add(CREATE_FRAME_ID, FOOTER_FRAME_NAME)
        With pFooterContainer
            .height = FOOTER_PANEL_HEIGHT
            .width = pFrame.InsideWidth
            .top = FOOTER_PANEL_HEIGHT
            .left = 0
            .SpecialEffect = fmSpecialEffectFlat
            .borderStyle = fmBorderStyleNone
        End With
    End If
    
End Sub

Private Sub applyColumnsLeftPosition()
    Dim left As Single
    Dim key As Variant
    Dim column As DColumn
    '------------------------------------------------------------------------------------------------------
    
    left = getHorizontalPadding
    For Each key In pColumns.keys
        Set column = pColumns.item(key)
        Call column.setLeft(left)
        left = left + column.getWidth + getColumnsSpace
    Next key
    
End Sub

Private Sub renderHeaderLine()
    Set pHeaderLine = UI.Create.GridHeaderLine(Me, pHeadersContainer)
    Call pHeaderLine.render
End Sub

Private Sub renderItemLines()
    Dim item As IGridItem
    Dim counter As Long
    Dim totalItems As Long
    '------------------------------------------------------------------------------------------------------
    
    Set pLines = F.dictionaries.Create(False)
    
    '[Item lines]
    totalItems = pItems.Count
    For Each item In pItems
        Call addLineForItem(item)
        counter = counter + 1
        If counter Mod PROGRESS_BAR_UPDATE_STEP = 0 Then Call ProgressBar.setProgress(counter / totalItems)
    Next item
    
End Sub

Private Sub addLineForItem(item As IGridItem)
    Dim line As xxGridLine
    '------------------------------------------------------------------------------------------------------
    Set line = UI.Create.GridLine(Me, pRecordsContainer, item)
    Call line.render
    Call F.dictionaries.addItem(pLines, item.getKey, line, duplicateBehaviour_ThrowError)
End Sub

Public Sub insertVerticalLines()

End Sub

Public Sub arrangeComponents()
    If UI.controls.isValid(pHeadersContainer) Then
        With pHeadersContainer
            .height = HEADER_LINE_HEIGHT
            .width = pFrame.InsideWidth
            .top = 0
            .left = 0
        End With
    End If
    
    If UI.controls.isValid(pRecordsContainer) Then
        With pRecordsContainer
            .height = pFrame.InsideHeight - HEADER_LINE_HEIGHT - FOOTER_PANEL_HEIGHT
            .width = pFrame.InsideWidth
            .top = HEADER_LINE_HEIGHT
            .left = 0
        End With
    End If
    
    If UI.controls.isValid(pFooterContainer) Then
        With pFooterContainer
            .height = FOOTER_PANEL_HEIGHT
            .width = pFrame.InsideWidth
            .top = pFrame.InsideHeight - FOOTER_PANEL_HEIGHT + PIXEL_SIZE
            .left = 0
            .backColor = FOOTER_PANEL_BACK_COLOR
        End With
    End If
    
End Sub





'[LINES]
Private Function getLine(id As Variant) As xxGridLine
    Set getLine = F.dictionaries.getObject(pLines, id)
End Function

Private Function getLineByItem(item As Object) As xxGridLine
    Dim line As xxGridLine
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    For Each key In pLines
        Set line = pLines.item(key)
        If line.getItem Is item Then
            Set getLineByItem = line
            Exit For
        End If
    Next key
    
End Function

'[SORTING]
Public Function isAscendingSortOrder() As Boolean
    isAscendingSortOrder = pSortAscending
End Function

Public Function getSortColumn() As DColumn
    Set getSortColumn = pSortColumn
End Function

Public Sub sort(column As DColumn)
    pSortAscending = calculateSortAscending(column)
    Set pSortColumn = column
    RaiseEvent action(ACTION_SORT, getSortParams)
    Call arrangeGrid
End Sub

Private Function getSortParams() As Scripting.Dictionary
    Set getSortParams = F.dictionaries.createWithItems(False, _
                            KeyValue(TAG_ORDER, pSortAscending), _
                            KeyValue(TAG_COLUMN, pSortColumn.getName))
End Function

Private Function calculateSortAscending(column As DColumn) As Boolean
    Dim activeSortColumnClicked As Boolean
    '------------------------------------------------------------------------------------------------------
    
    If Not pSortColumn Is Nothing Then
        If pSortColumn Is column Then
            activeSortColumnClicked = True
        End If
    End If
    
    If activeSortColumnClicked Then
        calculateSortAscending = Not pSortAscending
    Else
        calculateSortAscending = DEFAULT_SORT_ASCENDING
    End If
    
End Function

Private Sub arrangeGrid()
    Dim arr As Variant
    Dim i As Long
    Dim item As IGridItem
    Dim line As xxGridLine
    Dim top As Single:          top = VERTICAL_PADDING
    '------------------------------------------------------------------------------------------------------
    
    Call showHideLines
    arr = getSortedItems
    If F.arrays.isDefined(arr) Then
        For i = LBound(arr, 2) To UBound(arr, 2)
            Set item = arr(1, i)
            Set line = F.dictionaries.getObject(pLines, item.getKey)
            With line
                Call .hideOverflow
                Call .setTop(top)
                Call .setHasBorder(i > 1)
                top = .getBottom + LINES_SPACE
            End With
        Next i
    End If
    
    Call updateScrollBars(top)
    
End Sub

Private Sub showHideLines()
    Dim key As Variant
    Dim line As xxGridLine
    Dim visibility As Boolean
    '------------------------------------------------------------------------------------------------------
    
    For Each key In pLines
        Set line = pLines.item(key)
        visibility = pFiltered.Exists(key)
        Call line.setVisibility(visibility)
    Next key
    
End Sub

Private Function getSortedItems() As Variant
    Dim arr As Variant
    Dim key As Variant
    Dim i As Long
    Dim item As IGridItem
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    If F.dictionaries.isNonEmptyDictionary(pFiltered) Then
        ReDim arr(1 To 2, 1 To pFiltered.Count)
        For Each key In pFiltered.keys
            i = i + 1
            Set item = pFiltered.item(key)
            Set arr(1, i) = item
            Call F.Variables.assign(value, item.getValue(pSortColumn.getSortValuePropertyName))
            If VBA.IsNumeric(value) Then
                arr(2, i) = value
            Else
                arr(2, i) = F.Strings.stringify(value)
            End If
        Next key
        Call F.arrays.sort(arr, pSortAscending, 2, False)
        getSortedItems = arr
    End If
    
End Function



'[SCALING]
Private Sub updateScrollBars(y As Single)
    With pRecordsContainer
        If y > .InsideHeight Then
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = y
        Else
            .ScrollBars = fmScrollBarsNone
            .scrollTop = 0
        End If
    End With
End Sub

Public Sub adjustSize()
    If pIsRendered Then
        Call arrangeComponents
        Call adjustColumnWidths
        Call adjustLinesLayout
        Call arrangeFooterComponents
        Call arrangeGrid
    End If
End Sub

Private Sub adjustColumnWidths()
    Dim column As DColumn
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    Dim totalScalableWidth As Single
    Dim totalNotScalableWidth As Single
    Dim availableWidth As Single
    Dim factor As Double
    Dim width As Single
    '------------------------------------------------------------------------------------------------------
    
    For Each key In pColumns.keys
        Set column = pColumns.item(key)
        If column.isScalable Then
            totalScalableWidth = totalScalableWidth + column.getWidth
        Else
            totalNotScalableWidth = totalNotScalableWidth + column.getWidth
        End If
    Next key
    
    availableWidth = (pRecordsContainer.InsideWidth - totalNotScalableWidth - (pColumns.Count - 1) * COLUMNS_SPACE) - 2 * HORIZONTAL_PADDING
    factor = availableWidth / totalScalableWidth
    
    For Each key In pColumns.keys
        Set column = pColumns.item(key)
        If column.isScalable Then
            width = F.Numerics.roundTo(column.getWidth * factor, PIXEL_SIZE)
            Call column.setWidth(width)
        End If
    Next key
    
End Sub

Private Sub adjustLinesLayout()
    Dim key As Variant
    Dim line As xxGridLine
    '------------------------------------------------------------------------------------------------------
    
    Call pHeaderLine.arrangeControls
    For Each key In pLines.keys
        Set line = pLines.item(key)
        Call line.arrangeColumnsLayout
    Next key
    
End Sub



'[FILTERING]
Public Sub editFilter(header As xxGridHeader)
    Dim filter As IFilter
    '------------------------------------------------------------------------------------------------------
    
    RaiseEvent action(ACTION_FILTER_EDIT, _
                    F.dictionaries.createWithItems(False, KeyValue(TAG_COLUMN, header.getColumn.getName)))
    
    Set filter = header.getColumn.getFilter
    If Not filter Is Nothing Then
        With filter
            Call .setColumnName(header.getColumn.getLabel)
            Call .edit
        End With
    End If
    
End Sub

Public Sub removeFilter(header As xxGridHeader)
    Dim filter As IFilter
    '------------------------------------------------------------------------------------------------------
    
    RaiseEvent action(ACTION_FILTER_OFF, _
                    F.dictionaries.createWithItems(False, KeyValue(TAG_COLUMN, header.getColumn.getName)))
                    
    Set filter = header.getColumn.getFilter
    If Not filter Is Nothing Then
        Call filter.clear
        'call updateFilter(filter)
    End If
    
End Sub

Private Function getAvailableValuesForFilter(key As String) As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Dim item As IGridItem
    Dim column As DColumn
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set dict = F.dictionaries.Create(False)
    
    Set column = F.dictionaries.getObject(pColumns, key)
    If Not column Is Nothing Then
        For Each item In pItems
            If Not isItemIncludedInGlobalFilter(item) Then
                'skip
            ElseIf Not isItemIncludedInFilters(item, key) Then
                'skip
            Else
                Call F.Variables.assign(value, _
                    modifyValueWithFilterRegex(item.getValue(column.getValuePropertyName), column.getFilterRegex))
                If F.Variables.isPrimitive(value) Then
                    Call F.dictionaries.addItem(dict, value, value, duplicateBehaviour_Skip)
                ElseIf F.Variables.isContainer(value) Then
                    Call F.dictionaries.appendItemsFromContainer(dict, value, duplicateBehaviour_Skip)
                End If
            End If
        Next item
    End If
    
    Set getAvailableValuesForFilter = F.dictionaries.sortByKey(dict, True)
    
End Function

Private Function getFilterRegex(key As String) As String
    Dim column As DColumn
    '------------------------------------------------------------------------------------------------------
    Set column = getColumn(key)
    If Not column Is Nothing Then
        getFilterRegex = column.getFilterRegex
    End If
End Function

Private Function modifyValueWithFilterRegex(value As Variant, regex As String) As Variant
    Dim col As VBA.Collection
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(regex) = 0 Then
        Call F.Variables.assign(modifyValueWithFilterRegex, value)
    ElseIf F.Variables.isPrimitive(value) Then
        modifyValueWithFilterRegex = F.regex.getFirstGroupMatch(VBA.CStr(value), regex)
    ElseIf F.Variables.isContainer(value) Then
        Set col = F.Collections.createFrom(value)
        Set modifyValueWithFilterRegex = New VBA.Collection
        For Each item In col
            Call modifyValueWithFilterRegex.add(F.regex.getFirstGroupMatch(VBA.CStr(item), regex))
        Next item
    End If
    
End Function


Private Function isItemIncludedInFilters(item As IGridItem, Optional excludedFilter As String) As Boolean
    Dim keyFilter As Variant
    Dim filter As IFilter
    Dim column As DColumn
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    isItemIncludedInFilters = True
    
    For Each keyFilter In pFilters.keys
        Set filter = pFilters.item(keyFilter)
        If filter.isActive Then
            If filter.getKey <> excludedFilter Then
                Set column = F.dictionaries.getObject(pColumns, filter.getKey)
                Call F.Variables.assign(value, item.getValue(column.getValuePropertyName))
                If Not filter.match(value) Then
                    isItemIncludedInFilters = False
                    Exit For
                End If
            End If
        End If
    Next keyFilter
    
End Function

Private Sub updateFilter(filter As IFilter)
    RaiseEvent action(ACTION_FILTER_ON, getFilteronParams(filter))
    
    Call updateFilteredCollection
    Call arrangeGrid
    Call updateFilterIcon(filter)
End Sub

Private Function getFilteronParams(filter As IFilter) As Scripting.Dictionary
    Set getFilteronParams = F.dictionaries.createWithItems(False, _
                                KeyValue(TAG_COLUMN, pSortColumn.getName), _
                                KeyValue(TAG_VALUES, filter.getSelectionString))
End Function

Public Sub loadFiltersSettings(filters As Scripting.Dictionary)
    Dim key As Variant
    Dim sourceFilter As IFilter
    Dim destFilter As IFilter
    '------------------------------------------------------------------------------------------------------
    Dim isActive As Boolean
    Dim selection As Variant
    '------------------------------------------------------------------------------------------------------
    
    If F.dictionaries.isNonEmptyDictionary(filters) Then
        For Each key In filters.keys
            Set sourceFilter = filters.item(key)
            With sourceFilter
                isActive = .isActive
                Set selection = .getSelection
            End With
            
            Set destFilter = F.dictionaries.getObject(pFilters, key)
            If Not destFilter Is Nothing Then
                With destFilter
                    Call .setIsActive(isActive)
                    Call .setSelection(selection)
                End With
            End If
            
        Next key
    End If
    
    Call updateFilters
    
End Sub

Private Sub updateFilters()
    Dim key As Variant
    Dim filter As IFilter
    '------------------------------------------------------------------------------------------------------
    
    Call updateFilteredCollection
    Call arrangeGrid
    
    For Each key In pFilters.keys
        Set filter = pFilters.item(key)
        Call updateFilterIcon(filter)
    Next key
    
End Sub

Private Sub updateFilteredCollection()
    Dim item As IGridItem
    Dim line As xxGridLine
    '------------------------------------------------------------------------------------------------------
    
    Set pFiltered = F.dictionaries.Create(False)
    For Each item In pItems
        Set line = F.dictionaries.getObject(pLines, item.getKey)
        If isItemIncludedInGlobalFilter(item) Then
            If isItemIncludedInFilters(item, vbNullString) Then
                Call F.dictionaries.addItem(pFiltered, item.getKey, item)
            End If
        End If
    Next item
    
    Call updateRecordsCounter
    
End Sub

Private Function getActiveFilters() As VBA.Collection
    Dim filter As IFilter
    '------------------------------------------------------------------------------------------------------
    
    Set getActiveFilters = New VBA.Collection
    For Each filter In pFilters
        If filter.isActive Then
            Call getActiveFilters.add(filter)
        End If
    Next filter
    
End Function

Private Sub updateFilterIcon(filter As IFilter)
    Dim key As String
    Dim header As xxGridHeader
    '------------------------------------------------------------------------------------------------------
    
    key = filter.getKey
    Set header = pHeaderLine.getHeader(key)
    If Not header Is Nothing Then
        Call header.refreshFilterIcon
    End If
    
End Sub

Private Sub bindFiltersToColumns()
    Dim key As Variant
    Dim column As DColumn
    Dim filter As IFilter
    '------------------------------------------------------------------------------------------------------
    
    For Each key In pFilters
        Set column = F.dictionaries.getObject(pColumns, key)
        If Not column Is Nothing Then
            Set filter = pFilters.item(key)
            Call column.setFilter(filter)
        End If
    Next key
    
End Sub

Public Sub setGlobalFilter(value As String)
    pGlobalFilter = value
    Call updateFilteredCollection
    Call arrangeGrid
    
    If VBA.Len(value) Then
        RaiseEvent action(ACTION_FILTER_ON, F.dictionaries.createWithItems(False, KeyValue(TAG_FILTER, value)))
    Else
        RaiseEvent action(ACTION_FILTER_OFF, F.dictionaries.Create)
    End If
    
End Sub

Public Sub removeGlobalFilter()
    Call setGlobalFilter(vbNullString)
End Sub

Private Function isItemIncludedInGlobalFilter(item As IGridItem) As Boolean
    Dim column As DColumn
    Dim value As Variant
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(pGlobalFilter) Then
        For Each key In pColumns
            Set column = pColumns.item(key)
            If column.getControlType <> ControlType_Icon Then
                Call F.Variables.assign(value, item.getValue(column.getValuePropertyName))
                
                If F.Collections.isCollection(value) Then
                    value = F.Collections.toString(F.Cast.toCollection(value))
                End If
                
                If VBA.InStr(1, value, pGlobalFilter, vbTextCompare) Then
                    isItemIncludedInGlobalFilter = True
                    Exit For
                End If
            End If
        Next key
    Else
        isItemIncludedInGlobalFilter = True
    End If
    
End Function



'[ITEM ACTIONS]
Public Sub runItemAction(item As IGridItem, actionName As String)
    Select Case actionName
        Case ACTION_DELETE:         Call deleteItem(item)
        Case ACTION_EDIT:           Call editItem(item)
        Case Else:                  Call item.run(actionName)
    End Select
End Sub

Private Sub deleteItem(item As IGridItem)
    Dim question As String
    Dim logParams As Scripting.Dictionary
    Dim actionName As String
    '------------------------------------------------------------------------------------------------------
    
    Set logParams = F.dictionaries.createWithItems(False, KeyValue(TAG_ITEM, item.getDeleteInfoString), _
                                                          KeyValue(TAG_STEP, TAG_START))
    actionName = item.getActionName(ACTION_DELETE)
    RaiseEvent action(actionName, logParams)
    
    question = F.Strings.format(pDeleteItemQuestion, item.getDeleteInfoString)
    
    If F.Dialogs.getUserConfirmation(question, pDeleteItemTitle) Then
        Call F.dictionaries.addItem(logParams, TAG_STEP, TAG_CONFIRM)
        RaiseEvent action(actionName, logParams)
        
        Call item.run(ACTION_DELETE)
        Call F.Collections.removeItem(pItems, item)
        Call F.dictionaries.removeItem(pFiltered, item.getKey)
        Call arrangeGrid
        Call VBA.MsgBox(F.Strings.format(pDeleteSuccessInfo, item.getDeleteInfoString), vbInformation, App.getName)
    Else
        Call F.dictionaries.addItem(logParams, TAG_STEP, TAG_CANCEL)
        RaiseEvent action(actionName, logParams)
    End If
        
End Sub

Private Sub editItem(gridItem As IGridItem)
    Dim item As IEditItem
    Dim id As Long
    Dim validation As ValidationObject
    Dim line As xxGridLine
    Dim updateAfterward As Boolean
    '------------------------------------------------------------------------------------------------------
    
    '[Convert the given grid item into IEditItem object. If the given item doesn't implement this
    ' interface is cannot be edited].
    On Error Resume Next
    Set item = gridItem
    If item Is Nothing Then
        Call VBA.Err.Raise(1, description:="Item doesn't implement IEditItem interface, therefore it cannot be edited")
    Else
        id = item.getId
    End If
    
    If Not DEV_MODE Then On Error GoTo ErrHandler Else On Error GoTo 0
    
    Call pWindow.sendToBackground
    Set validation = item.edit
    Call pWindow.restoreFromBackground
    
    If validation.getStatus Then
        Set line = getLineByItem(item)
        If Not line Is Nothing Then
            Call line.updateValues
            Call arrangeGrid
        End If
        
        RaiseEvent AfterItemEdit(gridItem, updateAfterward)
        If updateAfterward Then Call updateAllItems
        
    ElseIf VBA.Len(validation.getMessage) = 0 Then
        ' Cancelled by user
    Else
        Call VBA.Err.Raise(1, description:=validation.getMessage)
    End If
    
ErrHandler:
    
End Sub


Private Sub updateAllItems()
    Dim key As Variant
    Dim line As xxGridLine
    '------------------------------------------------------------------------------------------------------
    For Each key In pLines.keys
        Set line = pLines.item(key)
        Call line.updateValues
    Next key
End Sub

Public Sub addItem(item As IGridItem)
    Call pItems.add(item)
    Call addLineForItem(item)
    Call updateFilteredCollection
    Call arrangeGrid
End Sub

Public Sub updateLines()
    Dim key As Variant
    Dim line As xxGridLine
    '------------------------------------------------------------------------------------------------------
    
    For Each key In pLines.keys
        Set line = pLines.item(key)
        Call line.updateValues
    Next key
    Call arrangeGrid
    
End Sub




'[FOOTER PANEL]
Private Sub renderFooter()
    Set pFooterTopBorder = pFooterContainer.controls.add(CREATE_LABEL_ID, FOOTER_BORDER_NAME, True)
    With pFooterTopBorder
        .width = pFooterContainer.InsideWidth
        .height = PIXEL_SIZE
        .left = 0
        .top = 0
        .borderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleOpaque
        .backColor = FOOTER_PANEL_BORDER_COLOR
    End With
    
    If pHasRecordsCounter Then
        Set pRecordsCounterLabel = pFooterContainer.controls.add(CREATE_LABEL_ID, SEARCH_ICON_NAME, True)
        With pRecordsCounterLabel
            .font.name = RECORDS_COUNTER_FONT_NAME
            .font.size = RECORDS_COUNTER_FONT_SIZE
            .caption = "jf"
            .width = RECORDS_COUNTER_WIDTH
            .AutoSize = True
            .AutoSize = False
            .width = RECORDS_COUNTER_WIDTH
            .BackStyle = fmBackStyleTransparent
        End With
        Call updateRecordsCounter
    End If
    
    Set pSearchIcon = pFooterContainer.controls.add(CREATE_LABEL_ID, SEARCH_ICON_NAME, True)
    With pSearchIcon
        .width = SEARCH_ICON_SIZE
        .height = SEARCH_ICON_SIZE
        .BackStyle = fmBackStyleTransparent
        Set .picture = Icons.icoSearch.picture
    End With
    
    Set pSearchBox = pFooterContainer.controls.add(CREATE_TEXTBOX_ID, SEARCH_TEXTBOX_NAME, True)
    With pSearchBox
        .width = SEARCH_TEXTBOX_WIDTH
        .height = SEARCH_TEXTBOX_HEIGHT
        .borderStyle = fmBorderStyleSingle
        .borderColor = SEARCH_TEXTBOX_BORDER_COLOR
        .font.name = SEARCH_TEXTBOX_FONT_NAME
        .font.size = SEARCH_TEXTBOX_FONT_SIZE
    End With
    
    Set pSearchCancelButton = pFooterContainer.controls.add(CREATE_LABEL_ID, SEARCH_CANCEL_BUTTON_NAME, False)
    With pSearchCancelButton
        .BackStyle = fmBackStyleTransparent
        .width = SEARCH_CANCEL_BUTTON_SIZE
        .height = SEARCH_CANCEL_BUTTON_SIZE
        Set .picture = Icons.icoDelete.picture
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = Icons.getMousePointer
    End With
    
    Set pChangeAllFiltersState = pFooterContainer.controls.add(CREATE_LABEL_ID, CHANGE_FILTERS_STATE_BUTTON_NAME, True)
    With pChangeAllFiltersState
        .BackStyle = fmBackStyleTransparent
        .width = SEARCH_CANCEL_BUTTON_SIZE
        .height = SEARCH_CANCEL_BUTTON_SIZE
        Set .picture = Icons.icoRemoveFilters.picture
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = Icons.getMousePointer
        .ControlTipText = MsgService.getText(TOOLTIP_DEACTIVATE_ALL_FILTERS)
    End With
    
    Set pExportToExcelButton = pFooterContainer.controls.add(CREATE_LABEL_ID, EXPORT_TO_EXCEL_BUTTON_NAME, True)
    With pExportToExcelButton
        .BackStyle = fmBackStyleTransparent
        .width = EXCEL_EXPORT_ICON_SIZE
        .height = EXCEL_EXPORT_ICON_SIZE
        Set .picture = Icons.icoExcel.picture
        .MousePointer = fmMousePointerCustom
        Set .MouseIcon = Icons.getMousePointer
        .ControlTipText = MsgService.getText(TOOLTIP_EXPORT_TO_EXCEL)
    End With
    
    Call arrangeFooterComponents
    
End Sub

Private Sub arrangeFooterComponents()
    Dim panelHeight As Single
    Dim left As Single
    '------------------------------------------------------------------------------------------------------
    
    panelHeight = pFooterContainer.InsideHeight
    
    '[Top border]
    With pFooterTopBorder
        .top = 0
        .left = 0
        .width = pFooterContainer.InsideWidth + 2 * PIXEL_SIZE
    End With
    
    left = FOOTER_CONTAINER_LEFT_PADDING
    
    '[Records counter]
    If pHasRecordsCounter Then
        Call UI.controls.locate(pRecordsCounterLabel, _
                    FOOTER_CONTAINER_LEFT_PADDING, _
                    (panelHeight - pRecordsCounterLabel.height) / 2)
        left = UI.controls.getRight(pRecordsCounterLabel) + RECORDS_COUNTER_RIGHT_MARGIN
    End If
                    
    '[Search controls]
    Call UI.controls.locate(pSearchIcon, left, (panelHeight - SEARCH_ICON_SIZE) / 2)
    Call UI.controls.locate(pSearchBox, _
                    UI.controls.getRight(pSearchIcon) + SEARCH_CONTROLS_SPACE, _
                    (panelHeight - SEARCH_TEXTBOX_HEIGHT) / 2)
    Call UI.controls.locate(pSearchCancelButton, _
                    UI.controls.getRight(pSearchBox) + SEARCH_CONTROLS_SPACE, _
                    (panelHeight - SEARCH_CANCEL_BUTTON_SIZE) / 2)
    
    '[Deactivate all filters]
    Call UI.controls.locate(pChangeAllFiltersState, _
                    UI.controls.getRight(pSearchCancelButton) + DEACTIVATE_FILTERS_BUTTON_RIGHT_MARGIN, _
                    (panelHeight - SEARCH_CANCEL_BUTTON_SIZE) / 2)
    
    '[Export to Excel]
    Call UI.controls.locate(pExportToExcelButton, _
                    pFooterContainer.InsideWidth - EXCEL_EXPORT_ICON_SIZE - FOOTER_CONTAINER_LEFT_PADDING, _
                    (panelHeight - EXCEL_EXPORT_ICON_SIZE) / 2)
    
End Sub





'[EVENTS]

Private Sub pSearchBox_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    Call ErrorManager.clear
    If keyCode = vbKeyReturn Then
        Call setSearchText(pSearchBox.value)
    End If
End Sub

Private Sub pSearchCancelButton_Click()
    Call ErrorManager.clear
    Call setSearchText(vbNullString)
End Sub

Private Sub setSearchText(value As String)
    pSearchText = value
    pSearchBox.value = value
    pSearchCancelButton.visible = (VBA.Len(pSearchText) > 0)
    Call setGlobalFilter(pSearchText)
End Sub

Private Sub updateRecordsCounter()
    Dim text As String
    '------------------------------------------------------------------------------------------------------
    If pHasRecordsCounter Then
        text = MsgService.getText(RECORDS_COUNTER_TAG)
        text = F.Strings.format(text, pFiltered.Count, pItems.Count)
        If UI.controls.isValid(pRecordsCounterLabel) Then
            pRecordsCounterLabel.caption = text
        End If
    End If
End Sub

Private Sub pChangeAllFiltersState_Click()
    Call ErrorManager.clear
    Call changeFiltersState(False)
End Sub

Private Sub changeFiltersState(value As Boolean)
    Dim key As Variant
    Dim filter As IFilter
    '------------------------------------------------------------------------------------------------------
    
    Call ErrorManager.clear
    For Each key In pFilters.keys
        Set filter = pFilters.item(key)
        Call filter.setIsActive(value)
    Next key
    
    Call updateFilters
    
End Sub


'[EXPORT TO EXCEL]
Private Sub pExportToExcelButton_Click()
    Call ErrorManager.clear
    Call exportToExcel
End Sub

Public Sub exportToExcel()
    Dim arr As Variant
    Dim xls As Excel.Application
    Dim wkb As Excel.Workbook
    Dim wks As Excel.Worksheet
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    RaiseEvent action(ACTION_EXPORT_TO_EXCEL, F.dictionaries.Create)
    
    arr = generateitemsArray
    If F.arrays.isDefined(arr) Then
        Set xls = VBA.CreateObject("Excel.Application")
        With xls
            .visible = False
            .SheetsInNewWorkbook = 1
            Set wkb = .Workbooks.add
            Set wks = wkb.Worksheets(1)
            Set rng = F.Ranges.pasteData(arr, wks.Cells(1, 1), True)
            If F.Ranges.isValid(rng) Then
                Call rng.EntireColumn.AutoFit
            End If
            .visible = True
        End With
    End If
    
End Sub

Private Function generateitemsArray(Optional includeHeaders As Boolean = True) As Variant
    Dim exportableColumns As VBA.Collection
    Dim column As DColumn
    Dim arrSortedItems As Variant
    '------------------------------------------------------------------------------------------------------
    Dim items As VBA.Collection
    Dim arr As Variant
    Dim i As Long
    Dim j As Long
    Dim item As IGridItem
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set exportableColumns = F.Collections.filterObjects(F.dictionaries.toCollection(pColumns), "{isExportable}")
    
    If F.Collections.isNonEmptyCollection(exportableColumns) Then
        arrSortedItems = getSortedItems
        If F.arrays.isDefined(arrSortedItems) Then
            Set items = F.Collections.fromArray(F.arrays.getRow(arrSortedItems, 1, False))
            If F.Collections.isNonEmptyCollection(items) Then
                ReDim arr(1 To exportableColumns.Count, VBA.IIf(includeHeaders, 0, 1) To items.Count)
                
                If includeHeaders Then
                    Call fillArrayWithHeaders(arr, exportableColumns)
                End If
                
                For Each item In items
                    i = i + 1
                    For j = 1 To exportableColumns.Count
                        Set column = exportableColumns.item(j)
                        Call F.Variables.assign(value, item.getValue(column.getValuePropertyName))
                        
                        If F.Variables.isPrimitive(value) Then
                            arr(j, i) = value
                        ElseIf F.Collections.isCollection(value) Then
                            arr(j, i) = F.Collections.toString(F.Cast.toCollection(value))
                        Else
                            Stop
                        End If
                        
                    Next j
                Next item
                
            End If
            
        Else
            ReDim arr(1 To exportableColumns.Count, 1 To 1)
            If includeHeaders Then
                Call fillArrayWithHeaders(arr, exportableColumns)
            End If
        End If
    End If
    
    generateitemsArray = arr
    
End Function


Public Sub fillArrayWithHeaders(ByRef arr As Variant, columns As VBA.Collection)
    Dim i As Long
    Dim column As DColumn
    '------------------------------------------------------------------------------------------------------
    
    For i = 1 To columns.Count
        Set column = columns.item(i)
        arr(i, LBound(arr, 2)) = column.getLabel
    Next i
    
End Sub



Public Sub destroy()
    If UI.controls.isValid(pFrame) Then
        With pFrame.controls
            If UI.controls.isValid(pHeadersContainer) Then
                Call .remove(pHeadersContainer.name)
            End If
            If UI.controls.isValid(pRecordsContainer) Then
                Call .remove(pRecordsContainer.name)
            End If
        End With
    End If
    
    Set pColumns = Nothing
    Set pHeaderLine = Nothing
    Set pItems = Nothing
    Set pLines = Nothing
    
End Sub











'[Implementation of IFilterable interface]
Private Sub IFilterable_applyFilter(filter As IFilter)
    Call updateFilter(filter)
End Sub

Private Sub IFilterable_removeFilter(filter As IFilter)
    Call updateFilter(filter)
End Sub

Private Function IFilterable_getAvailableValues(key As String) As Scripting.IDictionary
    Set IFilterable_getAvailableValues = getAvailableValuesForFilter(key)
End Function

Private Function IFilterable_getFilterRegex(key As String) As String
    IFilterable_getFilterRegex = getFilterRegex(key)
End Function

