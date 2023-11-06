VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoManager_Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRepoManager

Private Const CLASS_NAME As String = "RepoManager_Sheet"
'----------------------------------------------------------------------------------------------------------
Private Const ADD_BUTTON As String = "btnAdd"
Private Const CLOSE_BUTTON As String = "btnClose"
Private Const DELETE_ICON As String = "icoDelete"
Private Const EDIT_ICON As String = "icoEdit"
Private Const IS_EDIT_SHEET_NAME As String = "IS_EDIT_SHEET"
Private Const REPO_TAG_NAME As String = "REPO_TAG"
Private Const PROTECTION_PASSWORD As String = "a" 'RepoManager_ProtectionPassword"
'----------------------------------------------------------------------------------------------------------
Private Const COL_DELETE As Long = 2
Private Const COL_EDIT As Long = 3
Private Const COL_ID As Long = 5
Private Const COL_FIRST_DATA_COLUMN As Long = 6
Private Const ROW_HEADERS As Long = 3
'[View properties] ----------------------------------------------------------------------------------------
Private Const ROWS_HEIGHT As Single = 18
'----------------------------------------------------------------------------------------------------------
Private pRepo As Repository2
Private pParent As IRepoParent
Private pBook As Excel.Workbook
Private WithEvents pSheet As Excel.Worksheet
Attribute pSheet.VB_VarHelpID = -1
Private WithEvents pListener As ActionListener
Attribute pListener.VB_VarHelpID = -1
'----------------------------------------------------------------------------------------------------------
Private pColumns As VBA.Collection
Private pHeaders As Variant
Private pPropertiesByIndex As Variant
'----------------------------------------------------------------------------------------------------------
Private pName As String
Private pIdField As String
Private pItems As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------
Private pRange As Excel.Range
'[State] --------------------------------------------------------------------------------------------------
Private pSortingAllowed As Boolean
Private pSortColumnIndex As Long
Private pSortAscending As Boolean
'----------------------------------------------------------------------------------------------------------



Private Sub Class_Initialize()
    Set pBook = MApp.getRepoEditBook
    If pBook Is Nothing Then
        Set pBook = Excel.Workbooks("utd.xlsm")
    End If
    Set pListener = ActionListener
End Sub

Friend Sub destroy()
    Call removeSheet
    Set pRepo = Nothing
    Set pParent = Nothing
    Set pBook = Nothing
    Set pSheet = Nothing
    Set pListener = Nothing
    Set pRange = Nothing
End Sub

Private Sub removeSheet()
    If F.Books.isValid(pBook) Then
        Call F.Sheets.Delete(pBook, pSheet.name)
    End If
End Sub



'[Setters]
Public Function setRepo(value As Repository2) As RepoManager_Sheet
    Set setRepo = Me
    Set pRepo = value
    Set pParent = pRepo.getParent
    pName = pParent.getName
End Function




'[Rendering]
Public Sub show()
    Excel.Application.ScreenUpdating = False
    Set pSheet = insertSheet
    Call prepareSheet
    Excel.Application.ScreenUpdating = True
End Sub

Private Function insertSheet() As Excel.Worksheet
    Call removePreviousSheets
    Call wksEdit.Copy(Before:=pBook.Worksheets(1))
    Set insertSheet = pBook.Worksheets(1)
    insertSheet.name = pParent.getSheetName
    Call F.ExcelNames.addValue(insertSheet, REPO_TAG_NAME, pParent.getName)
End Function

Private Sub removePreviousSheets()
    Dim sheetToBeRemoved As VBA.Collection
    Dim wks As Excel.Worksheet
    Dim wksName As Variant
    Dim repoTag As String:                      repoTag = pParent.getName
    '------------------------------------------------------------------------------------------------------
    
    Set sheetToBeRemoved = New VBA.Collection
    
    For Each wks In pBook.Worksheets
        If shouldBeDeleted(wks, repoTag) Then
            Call sheetToBeRemoved.Add(wks.name)
        End If
    Next wks
    
    Excel.Application.DisplayAlerts = False
    For Each wksName In sheetToBeRemoved
        Set wks = pBook.Worksheets(wksName)
        Call wks.Delete
    Next wksName
    Excel.Application.DisplayAlerts = True
    
End Sub

Private Function shouldBeDeleted(wks As Excel.Worksheet, repoTag As String) As Boolean
    Dim isEditSheet As Variant
    Dim TAG As String
    '------------------------------------------------------------------------------------------------------
    isEditSheet = F.ExcelNames.getValue(wks, IS_EDIT_SHEET_NAME)
    If isEditSheet Then
        TAG = F.ExcelNames.getValue(wks, REPO_TAG_NAME)
        shouldBeDeleted = (TAG = repoTag)
    End If
End Function


Private Function prepareSheet() As Excel.Worksheet
    
    '[Hide/show add button]
    If Not pParent.isAddingAllowed Then
        'Call f.Sheets.removeShape(pSheet, ADD_BUTTON)
        pSheet.rows(1).Hidden = True
    Else
        Call updateAddButtonCaption
    End If
    
    '[Hide/show delete column]
    If Not pParent.isDeletingAllowed Then
        'Call removeDeleteIcon
        pSheet.columns(COL_DELETE).EntireColumn.Hidden = True
    End If
    
    Call insertColumns
    Call alignCloseButton
    Call insertValues
    
End Function

Private Sub updateAddButtonCaption()
    Dim btn As Excel.shape
    '------------------------------------------------------------------------------------------------------
    Set btn = F.Sheets.getShape(pSheet, ADD_BUTTON)
    If Not btn Is Nothing Then
        With btn.DrawingObject
            .caption = VBA.UCase$(MsgService.getText(.caption))
        End With
    End If
End Sub

Private Sub insertColumns()
    Dim columnDto As Scripting.Dictionary
    Dim index As Long
    Dim propName As String
    '------------------------------------------------------------------------------------------------------
    
    Set pColumns = pParent.getColumns
    
    index = COL_FIRST_DATA_COLUMN
    ReDim pPropertiesByIndex(index To index + pColumns.Count - 1)
    ReDim pHeaders(index To index + pColumns.Count - 1)
    
    For Each columnDto In pColumns
        propName = F.dictionaries.getItem(columnDto, TAG_COLUMN_PROPERTY)
        pPropertiesByIndex(index) = propName
        Call applyColumnProperties(index, columnDto)
        index = index + 1
    Next columnDto
    
End Sub

Private Sub alignCloseButton()
    Dim lastColumn As Long
    Dim rng As Excel.Range
    Dim columnRight As Single
    Dim shp As Excel.shape
    '------------------------------------------------------------------------------------------------------
    
    lastColumn = UBound(pPropertiesByIndex)
    Set rng = pSheet.columns(lastColumn)
    If F.Ranges.isValid(rng) Then
        columnRight = rng.left + rng.width
        Set shp = F.Sheets.getShape(pSheet, CLOSE_BUTTON)
        If Not shp Is Nothing Then
            With shp
                .top = 0
                .left = columnRight - .width
            End With
        End If
    End If
    
End Sub

Private Sub applyColumnProperties(index As Long, dto As Scripting.Dictionary)
    Dim width As Single
    Dim label As String
    Dim alignment As Variant
    Dim format As String
    Dim indent As Single
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    If index Then
        Set rng = pSheet.columns(index)
    
        width = F.dictionaries.getItem(dto, TAG_COLUMN_WIDTH)
        label = MsgService.getText(F.dictionaries.getItem(dto, TAG_COLUMN_LABEL))
        alignment = convertAlignTextToEnum(VBA.CStr(F.dictionaries.getItem(dto, TAG_COLUMN_ALIGNMENT)))
        format = F.dictionaries.getItem(dto, TAG_COLUMN_FORMAT)
        indent = F.dictionaries.getItem(dto, TAG_COLUMN_INDENT)
        
        If width > 0 Then rng.EntireColumn.ColumnWidth = width
        rng.Cells(ROW_HEADERS, 1).value = label
        If Not F.Variables.isUndefined(alignment, True) Then rng.HorizontalAlignment = alignment
        If VBA.Len(format) Then rng.NumberFormat = format
        If indent Then rng.IndentLevel = indent
        
        pHeaders(index) = label
        
    End If
End Sub

Private Sub insertValues()
    Dim key As Variant
    Dim item As IEditItem
    '------------------------------------------------------------------------------------------------------
    Dim arr As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim propName As String
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Call setupValues
    
    If pItems.Count Then
        rowIndex = ROW_HEADERS + 1
        ReDim arr(LBound(pPropertiesByIndex) - 1 To UBound(pPropertiesByIndex), _
                    rowIndex To rowIndex + pItems.Count - 1)
    
    
        For Each key In pItems.Keys
            Set item = pItems.item(key)
            For colIndex = LBound(arr, 1) To UBound(arr, 1)
                If colIndex = LBound(arr, 1) Then
                    propName = pIdField
                Else
                    propName = pPropertiesByIndex(colIndex)
                End If
                value = item.getValue(propName)
                arr(colIndex, rowIndex) = value
            Next colIndex
            rowIndex = rowIndex + 1
        Next key
        
        Set pRange = F.Ranges.pasteData(arr, pSheet.Cells(ROW_HEADERS + 1, COL_ID))
        pRange.EntireRow.Hidden = False
        
        Call copyActionIcons
        Call applySorting
        Call applyFilter
    End If
    
    Call formatRange
    Call trimSheet
    Call setProtection
    
End Sub

Private Sub setupValues()
    pIdField = pParent.getIdField
    If VBA.Len(pIdField) = 0 Then
        Call VBA.Err.Raise(1, Description:="Missing ID field for repository " & pParent.getName)
    End If
    Set pItems = pRepo.getItemsDictionary(pIdField)
End Sub

Private Sub copyActionIcons(Optional onlyLastRow As Boolean = False)
    Dim rng As Excel.Range
    Dim lastRow As Long
    Dim sh As Excel.shape
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    With pSheet
        lastRow = F.Ranges.getRangeLastRow(pRange)
        If onlyLastRow Then
            Set rng = .Range( _
                            .Cells(lastRow - 1, COL_DELETE), _
                            .Cells(lastRow, COL_EDIT))
        Else
            Set rng = .Range( _
                            .Cells(ROW_HEADERS + 1, COL_DELETE), _
                            .Cells(lastRow, COL_EDIT))
        End If
    End With
    Call rng.FillDown
    
    
    '[Rename delete and edit icons]
    For Each sh In pSheet.Shapes
        If F.Strings.startsWith(sh.name, EDIT_ICON) Then
            sh.name = EDIT_ICON & "_" & F.Math.increment(i)
        ElseIf F.Strings.startsWith(sh.name, DELETE_ICON) Then
            sh.name = DELETE_ICON & "_" & F.Math.increment(i)
        End If
    Next sh
    
End Sub

Private Sub applySorting()
    Dim sortingRules As Scripting.Dictionary
    Dim arr As Variant
    Dim i As Long
    Dim columnIndex As Long
    '------------------------------------------------------------------------------------------------------
    
    Set sortingRules = pParent.getSortingRules
    If F.dictionaries.isNonEmptyDictionary(sortingRules) Then
        arr = F.dictionaries.toArray(sortingRules, DictPart_KeyAndValue)
        For i = UBound(arr, 1) To LBound(arr, 1) Step -1
            columnIndex = F.arrays.getItemIndex(pPropertiesByIndex, arr(i, LBound(arr, 2)))
            If F.arrays.isInBounds(columnIndex, 1, pPropertiesByIndex) Then
                pSortColumnIndex = columnIndex
                pSortAscending = arr(i, UBound(arr, 2))
                Call sortValuesArray
            End If
        Next i
    End If
    
End Sub

Private Sub formatRange()
    If F.Ranges.isValid(pRange) Then
        With pRange
            .rows.EntireRow.RowHeight = ROWS_HEIGHT
        End With
    End If
End Sub

Private Sub trimSheet()
    Dim rngColumns As Excel.Range
    Dim rngRows As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    With pSheet
        If F.Ranges.isValid(pRange) Then
            Set rngColumns = .Range(.Cells(1, F.Ranges.getRangeLastColumn(pRange) + 1), .Cells(1, .columns.Count))
            Set rngRows = .Range(.Cells(F.Ranges.getRangeLastRow(pRange) + 1, 1), .Cells(.rows.Count, 1))
        Else
            Set rngColumns = .Range(.Cells(1, UBound(pPropertiesByIndex) + 1), .Cells(1, .columns.Count))
            Set rngRows = .Range(.Cells(ROW_HEADERS + 1, 1), .Cells(.rows.Count, 1))
        End If
        rngColumns.EntireColumn.Hidden = True
        rngRows.EntireRow.Hidden = True
    End With
    
End Sub


Private Sub applyFilter()
    If pParent.isFilteringAllowed Then
        pSheet.AutoFilterMode = False
        With pRange
            Call .Offset(-1).resize(.rows.Count + 1).AutoFilter
        End With
    End If
End Sub

Private Sub setProtection()
    Call pSheet.Protect( _
                    Password:=PROTECTION_PASSWORD, _
                    DrawingObjects:=True, _
                    Contents:=True, _
                    Scenarios:=True, _
                    UserInterfaceOnly:=True, _
                    AllowFormattingCells:=False, _
                    AllowFormattingColumns:=False, _
                    AllowFormattingRows:=False, _
                    AllowInsertingColumns:=False, _
                    AllowInsertingRows:=False, _
                    AllowInsertingHyperlinks:=False, _
                    AllowDeletingColumns:=False, _
                    AllowDeletingRows:=False, _
                    AllowSorting:=True, _
                    AllowFiltering:=True, _
                    AllowUsingPivotTables:=False)
End Sub


Friend Sub updateValues()
    
    Excel.Application.ScreenUpdating = True
    Call pSheet.Unprotect(PROTECTION_PASSWORD)
        
    'Remove previous range (except for the first row)
    If Not F.Ranges.isValid(pRange) Then
    ElseIf pRange.rows.Count <= 1 Then
    Else
        Call pRange.Offset(1).resize(pRange.rows.Count - 1).EntireRow.Delete
    End If
    
    Call insertValues
        
    Call pSheet.Protect(PROTECTION_PASSWORD)
    Excel.Application.ScreenUpdating = True
    
End Sub




'[Actions]
Private Sub editItem(row As Long)
    Dim id As Long
    Dim item As IEditItem
    Dim validation As ValidationObject
    '[Updating view] --------------------------------------------------------------------------------------
    Dim arr As Variant
    '------------------------------------------------------------------------------------------------------

    On Error Resume Next
    id = pSheet.Cells(row, COL_ID)
    Set item = F.dictionaries.GetObject(pItems, id)
    On Error GoTo 0
    
    If Not item Is Nothing Then
        Set validation = item.edit
        If validation.getStatus Then
            Call pRepo.updateItemKeys(item)
            arr = getItemLineValues(item)
            Call F.Ranges.pasteData(arr, pSheet.Cells(row, COL_ID), sheetPassword:=PROTECTION_PASSWORD)
            Call updateAffectedRows(item.getObjectsAffected)
            Call sortValuesArray
        End If
    End If

End Sub

Private Function getItemLineValues(item As IEditItem) As Variant
    Dim arr As Variant
    Dim colIndex As Long
    Dim propName As String
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    ReDim arr(LBound(pPropertiesByIndex) - 1 To UBound(pPropertiesByIndex))
    arr(LBound(arr, 1)) = item.getId
    For colIndex = LBound(arr) + 1 To UBound(arr)
        propName = pPropertiesByIndex(colIndex)
        value = item.getValue(propName)
        arr(colIndex) = value
    Next colIndex
    
    getItemLineValues = arr
    
End Function

Private Sub deleteItem(row As Long)
    Const DELETE_ITEM_TITLE As String = "GridManager.DeleteItem.WindowTitle"
    Const DELETE_ITEM_QUESTION As String = "GridManager.DeleteItem.ConfirmationQuestion"
    '------------------------------------------------------------------------------------------------------
    Dim id As Long
    Dim item As IEditItem
    Dim question As String
    Dim answer As Boolean
    '------------------------------------------------------------------------------------------------------
    
    Excel.Application.ScreenUpdating = False
    
    On Error Resume Next
    id = pSheet.Cells(row, COL_ID)
    Set item = F.dictionaries.GetObject(pItems, id)
    On Error GoTo 0
    
    If Not item Is Nothing Then
        question = F.Strings.format(MsgService.getText(DELETE_ITEM_QUESTION), item.getName)
        answer = F.Dialogs.getUserConfirmation(question, MsgService.getText(DELETE_ITEM_TITLE))
        If answer Then
            Call item.Delete
            Call pRepo.removeItem(item)
            Call setupValues
            Call updateViewAfterDeleting(row)
        End If
    End If
    
    Excel.Application.ScreenUpdating = True
    
End Sub

Private Sub updateViewAfterDeleting(row As Long)
    If pRange.rows.Count <= 1 Then
        Call pSheet.rows(row).ClearContents
        Set pRange = Nothing
        Call trimSheet
    Else
        Call F.Sheets.removeShapesByCell(pSheet, row)
        Call pSheet.rows(row).Delete
        Call trimSheet
    End If
End Sub

Private Sub addItem()
    Dim item As IEditItem
    Dim validation As ValidationObject
    '[Updating view] --------------------------------------------------------------------------------------
    Dim arr As Variant
    '------------------------------------------------------------------------------------------------------

    Excel.Application.ScreenUpdating = False

    Set item = pParent.createNewItem
    
    If Not item Is Nothing Then
        Call setupValues
        arr = getItemLineValues(item)
        Call addLineToRange(arr)
        Call updateAffectedRows(item.getObjectsAffected)
        Call sortValuesArray
    End If

    Excel.Application.ScreenUpdating = True

End Sub

Private Sub addLineToRange(arr As Variant)
    Dim rng As Excel.Range
    Dim row As Long
    '------------------------------------------------------------------------------------------------------
    
    If F.Ranges.isValid(pRange) Then
        With pRange
            Set rng = .Offset(.rows.Count).resize(1)
            rng.EntireRow.Hidden = False
            Set pRange = .resize(.rows.Count + 1)
        End With
    Else
        With pSheet
            Set pRange = .Range( _
                            .Cells(ROW_HEADERS + 1, COL_ID), _
                            .Cells(ROW_HEADERS + 1, UBound(pPropertiesByIndex) + 1))
            pRange.EntireRow.Hidden = False
        End With
    End If
    
    row = F.Ranges.getRangeLastRow(pRange)
    
    Call copyActionIcons(True)
    Call F.Ranges.pasteData(arr, pSheet.Cells(row, COL_ID), sheetPassword:=PROTECTION_PASSWORD)
    
End Sub



Private Sub updateAffectedRows(items As VBA.Collection)
    Dim item As IEditItem
    Dim row As Long
    Dim arr As Variant
    '------------------------------------------------------------------------------------------------------
            
    If F.Collections.isNonEmptyCollection(items) Then
        For Each item In items
            row = findItemRow(item)
            If row Then
                arr = getItemLineValues(item)
                Call F.Ranges.pasteData(arr, pSheet.Cells(row, COL_ID), sheetPassword:=PROTECTION_PASSWORD)
            End If
        Next item
    End If
    
End Sub

Private Function findItemRow(item As IEditItem) As Long
    Dim id As Long
    Dim rngIds As Excel.Range
    Dim arrIds As Variant
    '------------------------------------------------------------------------------------------------------
    
    id = item.getId
    Set rngIds = pRange.resize(, 1)
    arrIds = F.arrays.fromRange(rngIds)
    findItemRow = F.arrays.getItemIndex(arrIds, id) + ROW_HEADERS

End Function



'[Sorting]
Private Sub sortValuesArray(Optional columnIndex As Long, Optional ByRef Cancel As Boolean)
    Dim rngSortArea As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    If columnIndex = 0 Then
        'Re-sort after adding new item. Don't change order.
    ElseIf columnIndex = pSortColumnIndex Then
        pSortAscending = Not pSortAscending
    Else
        pSortAscending = True
        pSortColumnIndex = columnIndex
    End If
    
    If pSortColumnIndex Then
        With pRange
            Set rngSortArea = .Offset(-1, 0).resize(.rows.Count + 1, .columns.Count)
        End With
        Call rngSortArea.Sort( _
                    key1:=pSheet.Cells(ROW_HEADERS, pSortColumnIndex), _
                    Order1:=VBA.IIf(pSortAscending, xlAscending, xlDescending), _
                    header:=xlYes)
        Call updateSortIndicator
    End If
    
    Cancel = True
End Sub

Private Sub updateSortIndicator()
    Const SORT_INDICATOR_FONT_FAMILY As String = "Wingdings 3"
    Const SORT_ASCENDING_SYMBOL As String = "p"
    Const SORT_DESCENDING_SYMBOL As String = "q"
    '------------------------------------------------------------------------------------------------------
    Dim rng As Excel.Range
    Dim cell As Excel.Range
    Dim text As String
    '------------------------------------------------------------------------------------------------------
    
    With pSheet
        Set rng = .Range(.Cells(ROW_HEADERS, LBound(pHeaders)), .Cells(ROW_HEADERS, UBound(pHeaders)))
        Call F.Ranges.pasteData(pHeaders, rng, sheetPassword:=PROTECTION_PASSWORD)
        Call pSheet.Unprotect(PROTECTION_PASSWORD)
        Call setProtection
        
        '[Add sorting indicator to the current sorting column]
        Set cell = .Cells(ROW_HEADERS, pSortColumnIndex)
        text = pHeaders(pSortColumnIndex) & " " & VBA.IIf(pSortAscending, SORT_ASCENDING_SYMBOL, SORT_DESCENDING_SYMBOL)
        With cell
            .Value2 = text
            .Characters(Start:=VBA.Len(text), Length:=1).Font.name = SORT_INDICATOR_FONT_FAMILY
        End With
    End With
    
End Sub




'[Catching events]
Private Sub pListener_AfterClick(callerName As String, sheet As Worksheet, row As Long, column As Long)
    If Not sheet Is pSheet Then
    ElseIf callerName = ADD_BUTTON Then
        Call addItem
    ElseIf callerName = CLOSE_BUTTON Then
        Call MApp.unregisterRepoManager(Me)
    ElseIf column = COL_EDIT Then
        Call editItem(row)
    ElseIf column = COL_DELETE Then
        Call deleteItem(row)
    End If
End Sub

Private Sub pSheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.row = ROW_HEADERS Then
        If pParent.isSortingAllowed Then
            Call sortValuesArray(Target.column, Cancel)
        End If
    End If
End Sub





'[Implementation of IRepoManager]

Private Function IRepoManager_getName() As String
    IRepoManager_getName = pName
End Function

Private Function IRepoManager_setRepo(value As Repository2) As IRepoManager
    Set IRepoManager_setRepo = setRepo(value)
End Function

Private Sub IRepoManager_show()
    Call show
End Sub

Private Sub IRepoManager_destroy()
    Call destroy
End Sub

Private Sub IRepoManager_updateValues()
    Call updateValues
End Sub
