VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Const CLASS_NAME As String = "FRanges"
    '----------------------------------------------------------------------------------------------------------
Private mrBuild As Excel.Range
    '----------------------------------------------------------------------------------------------------------


Public Function getNonEmptyRange(wks As Excel.Worksheet, _
                    Optional startFromFirstRow As Boolean = False, _
                    Optional startFromFirstColumn As Boolean = False) As Excel.Range
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim row As Long
    Dim col As Long
    Dim i As Long
    Dim arr As Variant
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    lastRow = wks.Cells.Find(what:="*", After:=wks.Cells(1, 1), lookAt:=xlPart, lookIn:=xlFormulas, _
                            searchOrder:=xlByRows, searchDirection:=xlPrevious, matchCase:=False).row
    If lastRow = 0 Then GoTo EmptySheetException
                    
    lastCol = wks.Cells.Find(what:="*", After:=wks.Cells(1, 1), lookAt:=xlPart, lookIn:=xlFormulas, _
                            searchOrder:=xlByColumns, searchDirection:=xlPrevious, matchCase:=False).column
    If lastCol = 0 Then GoTo EmptySheetException
    On Error GoTo 0
    
    With wks
        arr = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
    End With
    
    If startFromFirstRow Then
        firstRow = 1
    Else
        For row = 1 To lastRow
            If Not VBA.IsEmpty(arr(row, lastCol)) Then
                firstRow = row
                Exit For
            End If
            col = wks.Cells(row, lastCol).End(xlToLeft).column
            If col > 1 Then
                firstRow = row
                Exit For
            ElseIf Not VBA.IsEmpty(wks.Cells(row, col)) Then
                firstRow = row
                Exit For
            End If
        Next row
    End If
    
    If startFromFirstColumn Then
        firstCol = 1
    Else
        For col = 1 To lastCol
            If Not VBA.IsEmpty(arr(lastRow, col)) Then
                firstCol = col
                Exit For
            End If
            row = wks.Cells(lastRow, col).End(xlUp).row
            If row > 1 Then
                firstCol = col
                Exit For
            ElseIf Not VBA.IsEmpty(wks.Cells(row, col)) Then
                firstCol = col
                Exit For
            End If
        Next col
    End If
    
    
    If firstRow <= lastRow And firstCol <= lastCol Then
        With wks
            Set getNonEmptyRange = .Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol))
        End With
    End If

'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
EmptySheetException:
    GoTo ExitPoint

End Function

Private Sub yieldBounds(rng As Excel.Range, ByRef firstRow As Long, ByRef firstCol As Long, ByRef lastRow As Long, ByRef lastCol As Long)
    Dim subrng As Excel.Range
    Dim row As Long
    Dim col As Long
    '------------------------------------------------------------------------------------------------------
    
    If isValid(rng) Then
    
        'First row
        For Each subrng In rng.Areas
            If firstRow = 1 Then Exit For
            row = subrng.row
            If row < firstRow Then firstRow = row
        Next subrng
        
        'First column
        For Each subrng In rng.Areas
            If firstCol = 1 Then Exit For
            col = subrng.column
            If col < firstCol Then firstCol = col
        Next subrng
        
        'Last row
        For Each subrng In rng.Areas
            row = subrng.row + subrng.rows.count - 1
            If row > lastRow Then lastRow = row
        Next subrng
        
        'Last column
        For Each subrng In rng.Areas
            col = subrng.column + subrng.columns.count - 1
            If col > lastCol Then lastCol = col
        Next subrng
        
    End If
    
End Sub

Public Function getNonEmptyCells(rng As Excel.Range) As Excel.Range
    Dim rngFormulas As Variant 'Excel.Range
    Dim rngConstants As Variant 'Excel.Range
    '----------------------------------------------------------------------

    On Error Resume Next
    With rng
        Set rngFormulas = .SpecialCells(xlCellTypeFormulas)
        Set rngConstants = .SpecialCells(xlCellTypeConstants)
    End With
    On Error GoTo 0

    Set getNonEmptyCells = F.ranges.join(F.Collections.Create(rngFormulas, rngConstants))

End Function

Public Function join(ParamArray ranges() As Variant) As Excel.Range
    Dim arrRanges As Variant:   arrRanges = ranges
    Dim colRanges As VBA.Collection
    Dim varRange As Variant
    Dim rng As Excel.Range
    Dim parent As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------

    Set colRanges = New VBA.Collection

    For Each varRange In ranges
        If VBA.TypeName(varRange) = "Range" Then
            Call colRanges.Add(varRange)
        ElseIf F.Variables.isContainer(varRange) Then
            Set colRanges = F.Collections.join(colRanges, varRange)
        End If
    Next varRange

    For Each varRange In colRanges
        If VBA.IsObject(varRange) Then
            If varRange Is Nothing Then
            
            ElseIf TypeOf varRange Is Excel.Range Then
                Set rng = varRange
                If isValid(rng) Then
                    If join Is Nothing Then
                        Set join = rng
                        Set parent = rng.parent
                    Else
                        If Not rng.parent Is parent Then Call VBA.err.raise(1, description:="DifferentWorksheetsException")
                        Set rng = varRange
                        Set join = rng.Application.Union(join, rng)
                    End If
                End If
            End If
        End If
    Next varRange
    
End Function



'Public Function getRangeFirstColumn(rng As Excel.Range) As Long
'    Dim subrange As Excel.Range
'    '------------------------------------------------------------------------------------------------------
'
'    Call throwExceptionIfRangeIsInvalid(rng)
'    getRangeFirstColumn = rng.parent.columns.Count + 1
'    For Each subrange In rng.Areas
'        getRangeFirstColumn = Excel.Application.WorksheetFunction.Min(getRangeFirstColumn, subrange.column)
'    Next subrange
'
'End Function
'
'Public Function getRangeFirstRow(rng As Excel.Range) As Long
'    Dim subrange As Excel.Range
'    '------------------------------------------------------------------------------------------------------
'
'    If isRangeValid(rng) Then
'        getRangeFirstRow = rng.parent.rows.Count + 1
'        For Each subrange In rng.Areas
'            getRangeFirstRow = Excel.Application.WorksheetFunction.Min(getRangeFirstRow, subrange.row)
'        Next subrange
'    End If
'
'End Function
'
'Public Function getRangeLastColumn(rng As Excel.Range) As Long
'    Dim subrange As Excel.Range
'    Dim subrangeLastColumn As Long
'    '------------------------------------------------------------------------------------------------------
'
'    If isRangeValid(rng) Then
'        For Each subrange In rng.Areas
'            subrangeLastColumn = subrange.column + subrange.columns.Count - 1
'            getRangeLastColumn = Excel.Application.WorksheetFunction.Max(getRangeLastColumn, subrangeLastColumn)
'        Next subrange
'    End If
'
'End Function
'
'Public Function getRangeLastRow(rng As Excel.Range) As Long
'    Dim subrange As Excel.Range
'    Dim subrangeLastRow As Long
'    '------------------------------------------------------------------------------------------------------
'
'    If isRangeValid(rng) Then
'        For Each subrange In rng.Areas
'            subrangeLastRow = subrange.row + subrange.rows.Count - 1
'            getRangeLastRow = Excel.Application.WorksheetFunction.Max(getRangeLastRow, subrangeLastRow)
'        Next subrange
'    End If
'
'End Function
'


Public Function findCellsByContent(rng As Range, what As Variant, Optional lookIn As XlFindLookIn = xlValues, Optional lookAt As XlLookAt = xlWhole, _
                        Optional searchOrder As XlSearchOrder = xlByColumns, Optional searchDirection As XlSearchDirection = xlNext, _
                        Optional matchCase As Boolean = False, Optional matchByte As Boolean = False, Optional searchFormat As Boolean = False) As Excel.Range
    Dim searchResult As Range
    Dim firstMatch As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    With rng
        Set searchResult = .Find(what, , lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte, searchFormat)
        If Not searchResult Is Nothing Then
            Set firstMatch = searchResult
            'firstMatch = searchResult.address
            
            Do
                If findCellsByContent Is Nothing Then
                    Set findCellsByContent = searchResult
                Else
                    Set findCellsByContent = join(findCellsByContent, searchResult)
                End If
                Set searchResult = .FindNext(searchResult)
                
                If searchResult Is Nothing Then
                    Exit Do
                ElseIf Not rng.Application.intersect(searchResult, firstMatch) Is Nothing Then
                    Exit Do
                End If
            Loop
            
        End If
    End With
    
End Function

Public Function findRowByLabel(wks As Excel.Worksheet, label As String) As Long
    Dim cell As Excel.Range
    '------------------------------------------------------------------------------------------------------
    With wks
        Set cell = .Cells.Find(what:=label, After:=.Cells(1, 1), lookAt:=xlWhole, lookIn:=xlFormulas, searchOrder:=xlByRows, searchDirection:=xlNext, matchCase:=False)
        If Not cell Is Nothing Then findRowByLabel = cell.row
    End With
End Function

Public Function findColumnByLabel(wks As Excel.Worksheet, label As String) As Long
    Dim cell As Excel.Range
    '------------------------------------------------------------------------------------------------------
    With wks
        Set cell = .Cells.Find(what:=label, After:=.Cells(1, 1), lookAt:=xlWhole, lookIn:=xlFormulas, searchOrder:=xlByRows, searchDirection:=xlNext, matchCase:=False)
        If Not cell Is Nothing Then findColumnByLabel = cell.column
    End With
End Function

Public Sub setInsideBorders(rng As Excel.Range, Optional color As Long = 14277081)
    With rng.Borders(xlInsideHorizontal)
        .color = color
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .color = color
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Public Sub clearInsideBorders(rng As Excel.Range)
    If F.ranges.isValid(rng) Then
        With rng.Borders(xlInsideHorizontal)
            .LineStyle = xlLineStyleNone
        End With
        With rng.Borders(xlInsideVertical)
            .LineStyle = xlLineStyleNone
        End With
    End If
End Sub



'If the given range is a few cells merged in one area - return only the first of them.
Public Function getFirstCell(rng As Excel.Range) As Excel.Range
    Dim objName As Excel.name
    Dim firstCell As Excel.Range
    Dim mergedArea As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    If rng.Cells.count > 1 Then
        Set firstCell = rng.Cells(1, 1)
        If firstCell.MergeCells Then
            Set mergedArea = firstCell.MergeArea
            If F.strings.compareStrings(mergedArea.address, rng.address) Then
                Set getFirstCell = firstCell
            End If
        End If
    End If
    
    If getFirstCell Is Nothing Then
        Set getFirstCell = rng
    End If
    
End Function

Public Function isValid(rng As Excel.Range) As Boolean
    Dim lngRangeRow As Long
    '------------------------------------------------------------------------------------------------------
    On Error Resume Next
    lngRangeRow = rng.row
    If lngRangeRow > 0 Then isValid = True
End Function

Public Function pasteData(data As Variant, initRange As Excel.Range, Optional transponeData As Boolean = True, _
                        Optional clearSheet As Boolean = False, Optional sheetPassword As Variant) As Excel.Range
    Dim arr As Variant
    Dim wks As Excel.Worksheet
    Dim rng As Excel.Range
    Dim enableEventsCache As Boolean
    Dim protected As Boolean
    Dim filter As ExcelFilter
    '------------------------------------------------------------------------------------------------------
    Dim firstRow As Long
    Dim firstCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    '------------------------------------------------------------------------------------------------------

    
    With Application
        enableEventsCache = .EnableEvents
        .EnableEvents = False
    End With
    
    If Not isValid(initRange) Then Call Exceptions.InvalidRange.raise
    
    arr = F.arrays.to2DArray(data)
    If F.arrays.countDimensions(arr) <> 2 Then Call Exceptions.InvalidDimension.raise
    
    With initRange
        
        Set wks = .parent
        
        If wks.ProtectContents Then
            protected = True
            Call wks.Unprotect(sheetPassword)
        End If
        
        If clearSheet Then
            Call wks.Cells.ClearContents
        Else
            Set filter = F.Create.ExcelFilter(wks)
            Call filter.readFromSheet
        End If
        
        If transponeData Then arr = F.arrays.transposeArray(arr)
                                                                                                        
        firstRow = .row
        firstCol = .column
        lastRow = .row + F.arrays.getSize(arr, 1) - 1
        lastCol = .column + F.arrays.getSize(arr, 2) - 1
                                                                                                        
    End With
        
    With wks
        Set rng = .Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol))
    End With
    
    If Not filter Is Nothing Then Call filter.removeFilter
    rng = arr
    If Not filter Is Nothing Then Call filter.applyFilter
    
    Set pasteData = rng

'==========================================================================================================
ExitPoint:
    
    Application.EnableEvents = enableEventsCache
    If protected Then Call wks.Protect(sheetPassword)
    
    Exit Function

End Function

Public Function getCellAddress(cell As Excel.Range, Optional rowLocked As Boolean = True, Optional columnLocked As Boolean = True)
    Dim defaultAddress As String
    Dim columnIndex As String
    Dim row As Long
    '------------------------------------------------------------------------------------------------------
    defaultAddress = cell.address
    columnIndex = F.strings.substring(defaultAddress, "$", "$")
    row = VBA.CLng(F.strings.substring(defaultAddress, columnIndex & "$", vbNullString))
    getCellAddress = VBA.IIf(columnLocked, "$", vbNullString) & columnIndex & VBA.IIf(rowLocked, "$", vbNullString) & row
End Function

Public Sub clearFormatConditions(rng As Excel.Range)
    Dim fc As Excel.FormatCondition
    '------------------------------------------------------------------------------------------------------
    If isValid(rng) Then
        For Each fc In rng.FormatConditions
            Call fc.delete
        Next fc
    End If
End Sub

Public Sub pasteAsValues(rng As Excel.Range, Optional protectionPassword As String)
    Dim wks As Excel.Worksheet
    Dim isProtected As Boolean
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo Exception
    
    If isValid(rng) Then
        With rng
            Set wks = .parent
            isProtected = wks.ProtectContents
            If isProtected Then Call wks.Unprotect(protectionPassword)
            Call .Copy
            Call .PasteSpecial(Paste:=xlPasteValues, operation:=xlNone, SkipBlanks:=False, transpose:=False)
            If isProtected Then Call wks.Protect(protectionPassword)
        End With
    End If

'==========================================================================================================
ExitPoint:
    Exit Sub

    '----------------------------------------------------------------------------------------------------------
Exception:
    'Error handling
    GoTo ExitPoint

End Sub





Public Function getFirstNonEmptyRow(wks As Excel.Worksheet, _
                      Optional startRow As Long, Optional startCol As Long, _
                      Optional endRow As Long, Optional endCol As Long, _
                      Optional ignoreHiddenCells As Boolean = False) As Long
    Const METHOD_NAME As String = "firstNonEmptyRow"
    '------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim lngRowStart As Long
    Dim lngRowEnd As Long
    Dim lngNonBlanks As Long
    Dim lngColStart As Long
    Dim lngColEnd As Long
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------


    'Checks if the given worksheet is valid and can be referred to. -------------------------------------|
    'If not, the code jumps to the label IllegalSheetException, where you can define your own           '|
    'error handling rules for this exception.                                                           '|
    If Not F.Sheets.isValid(wks) Then GoTo IllegalSheetException                                             '|
    '----------------------------------------------------------------------------------------------------|


    'Calculate the actual range to be searched including optional parameters passed to the function. ----|
    If startCol > 0 And startCol <= wks.columns.count Then lngColStart = startCol Else lngColStart = 1  '|
    If endCol > 0 And endCol <= wks.columns.count Then lngColEnd = endCol Else _
                                                                        lngColEnd = wks.columns.count   '|
    If startRow > 0 And startRow <= wks.rows.count Then lngRowStart = startRow Else lngRowStart = 1     '|
    If endRow > 0 And endRow <= wks.rows.count Then lngRowEnd = endRow Else lngRowEnd = wks.rows.count  '|
    '----------------------------------------------------------------------------------------------------|




Retry:

    '----------------------------------------------------------------------------------------------------|
    If endRow > 0 And endRow <= wks.rows.count Then                                                     '|
        lngRow = endRow                                                                                 '|
    Else                                                                                                '|
        lngRow = wks.rows.count                                                                         '|
    End If                                                                                              '|
    '----------------------------------------------------------------------------------------------------|


    '----------------------------------------------------------------------------------------------------|
    Do                                                                                                  '|
        Set rng = wks.Range(wks.Cells(lngRowStart, lngColStart), wks.Cells(lngRow, lngColEnd))          '|
        lngNonBlanks = wks.Application.WorksheetFunction.CountA(rng)                                    '|
                                                                                                        '|
                                                                                                        '|
        '--------------------------------------------------------------------------------------------|  '|
        If lngNonBlanks Then                                                                        '|  '|
                                                                                                    '|  '|
            If lngRow = lngRowStart Then Exit Do                                                    '|  '|
            lngRowEnd = lngRow                                                                      '|  '|
            lngRow = lngRowStart + ((lngRow - lngRowStart - 1) / 2)                                 '|  '|
                                                                                                    '|  '|
        Else                                                                                        '|  '|
            lngRowStart = lngRow + 1                                                                '|  '|
            lngRow = lngRowEnd                                                                      '|  '|
                                                                                                    '|  '|
            '------------------------------------------------------------------------------------|  '|  '|
            If lngRowStart > lngRowEnd Then                                                     '|  '|  '|
                lngRow = 0                                                                      '|  '|  '|
                Exit Do                                                                         '|  '|  '|
            End If                                                                              '|  '|  '|
            '------------ [If lngRowStart > lngRowEnd Then] -------------------------------------|  '|  '|
                                                                                                    '|  '|
        End If                                                                                      '|  '|
        '---------------- [If lngNonBlanks Then] ----------------------------------------------------|  '|
                                                                                                        '|
    Loop                                                                                                '|
    '----------------------------------------------------------------------------------------------------|


    '----------------------------------------------------------------------------------------------------|
    If lngRow Then                                                                                      '|
                                                                                                        '|
        '--------------------------------------------------------------------------------------------|  '|
        If ignoreHiddenCells And wks.rows(lngRow).Hidden Then                                       '|  '|
            lngRowStart = nextVisibleRow(wks, lngRow, Excel.xlDown)                                 '|  '|
            If lngRowStart <= wks.rows.count Then GoTo Retry                                        '|  '|
        Else                                                                                        '|  '|
            getFirstNonEmptyRow = lngRow                                                            '|  '|
        End If                                                                                      '|  '|
        '---------------- [If ignoreHiddenCells And wks.rows(lngRow).Hidden Then] -------------------|  '|
                                                                                                        '|
    End If                                                                                              '|




'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function

Public Function firstNonEmptyColumn(wks As Excel.Worksheet, _
                   Optional startRow As Long, Optional startCol As Long, _
                   Optional endRow As Long, Optional endCol As Long, _
                   Optional ignoreHiddenCells As Boolean = False) As Long
    Const METHOD_NAME As String = "firstNonEmptyColumn"
    '------------------------------------------------------------------------------------------------------
    Dim lngCol As Long
    Dim lngColStart As Long
    Dim lngColEnd As Long
    Dim lngNonBlanks As Long
    Dim lngRowStart As Long
    Dim lngRowEnd As Long
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------


    'Checks if the given worksheet is valid and can be referred to. -------------------------------------|
    'If not, the code jumps to the label IllegalSheetException, where you can define your own           '|
    'error handling rules for this exception.                                                           '|
    If Not F.Sheets.isValid(wks) Then GoTo IllegalSheetException                                            '|
    '----------------------------------------------------------------------------------------------------|
    
    
    'Calculate the actual range to be searched including optional parameters passed to the function. ----|
    If startCol > 0 And startCol <= wks.columns.count Then lngColStart = startCol Else lngColStart = 1  '|
    If endCol > 0 And endCol <= wks.columns.count Then lngColEnd = endCol Else _
                                                                        lngColEnd = wks.columns.count   '|
    If startRow > 0 And startRow <= wks.rows.count Then lngRowStart = startRow Else lngRowStart = 1     '|
    If endRow > 0 And endRow <= wks.rows.count Then lngRowEnd = endRow Else lngRowEnd = wks.rows.count  '|
    '----------------------------------------------------------------------------------------------------|


Retry:

    '----------------------------------------------------------------------------------------------------|
    If endCol > 0 And endCol <= wks.columns.count Then                                                  '|
        lngCol = endCol                                                                                 '|
    Else                                                                                                '|
        lngCol = wks.columns.count                                                                      '|
    End If                                                                                              '|
    '----------------------------------------------------------------------------------------------------|
    

    '----------------------------------------------------------------------------------------------------|
    Do                                                                                                  '|
                                                                                                        '|
        Set rng = wks.Range(wks.Cells(lngRowStart, lngColStart), wks.Cells(lngRowEnd, lngCol))          '|
        lngNonBlanks = wks.Application.WorksheetFunction.CountA(rng)                                    '|
                                                                                                        '|
        '--------------------------------------------------------------------------------------------|  '|
        If lngNonBlanks Then                                                                        '|  '|
                                                                                                    '|  '|
            If lngCol = lngColStart Then Exit Do                                                    '|  '|
            lngColEnd = lngCol                                                                      '|  '|
            lngCol = lngColStart + ((lngCol - lngColStart - 1) / 2)                                 '|  '|
                                                                                                    '|  '|
        Else                                                                                        '|  '|
                                                                                                    '|  '|
            lngColStart = lngCol + 1                                                                '|  '|
            lngCol = lngColEnd                                                                      '|  '|
                                                                                                    '|  '|
            '------------------------------------------------------------------------------------|  '|  '|
            If lngColStart > lngColEnd Then                                                     '|  '|  '|
                lngCol = 0                                                                      '|  '|  '|
                Exit Do                                                                         '|  '|  '|
            End If                                                                              '|  '|  '|
            '------ [If lngColStart > lngColEnd Then] -------------------------------------------|  '|  '|
                                                                                                    '|  '|
        End If                                                                                      '|  '|
        '---------- [If lngNonBlanks Then] ----------------------------------------------------------|  '|
                                                                                                        '|
    Loop                                                                                                '|
    '----------------------------------------------------------------------------------------------------|
        

    '----------------------------------------------------------------------------------------------------|
    If lngCol Then                                                                                      '|
                                                                                                        '|
        '--------------------------------------------------------------------------------------------|  '|
        If ignoreHiddenCells And wks.columns(lngCol).Hidden Then                                    '|  '|
            lngColStart = nextVisibleColumn(wks, lngCol, Excel.xlToRight)                           '|  '|
            If lngColStart <= wks.columns.count Then GoTo Retry                                     '|  '|
        Else                                                                                        '|  '|
            firstNonEmptyColumn = lngCol                                                            '|  '|
        End If                                                                                      '|  '|
        '--------------------------------------------------------------------------------------------|  '|
                                                                                                        '|
    End If                                                                                              '|
    '----------------------------------------------------------------------------------------------------|



'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function


Public Function lastNonEmptyRow(wks As Excel.Worksheet, _
                      Optional startRow As Long, Optional startCol As Long, _
                      Optional endRow As Long, Optional endCol As Long, _
                      Optional ignoreHiddenCells As Boolean = False) As Long
    Dim lngRow As Long
    Dim lngRowStart As Long
    Dim lngRowEnd As Long
    Dim lngNonBlanks As Long
    Dim lngColStart As Long
    Dim lngColEnd As Long
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------

    If Not F.Sheets.isValid(wks) Then GoTo IllegalSheetException
    
    If startCol > 0 And startCol <= wks.columns.count Then lngColStart = startCol Else lngColStart = 1
    If endCol > 0 And endCol <= wks.columns.count Then lngColEnd = endCol Else _
                                                                        lngColEnd = wks.columns.count
    If startRow > 0 And startRow <= wks.rows.count Then lngRowStart = startRow Else lngRowStart = 1
    If endRow > 0 And endRow <= wks.rows.count Then lngRowEnd = endRow Else lngRowEnd = wks.rows.count

Retry:

    lngRow = 1

    Do
        Set rng = wks.Range(wks.Cells(lngRow, lngColStart), wks.Cells(lngRowEnd, lngColEnd))
        lngNonBlanks = wks.Application.WorksheetFunction.CountA(rng)
                                                                                                        
        If lngNonBlanks Then
            If lngRow = lngRowEnd Then Exit Do
            lngRowStart = lngRow
            lngRow = lngRow + ((lngRowEnd - lngRow + 1) / 2)
        Else
            lngRowEnd = lngRow - 1
            lngRow = lngRowStart
                                                                                                      
            If lngRowStart > lngRowEnd Then
                lngRow = 0
                Exit Do
            End If
                                                                                                      
        End If
    Loop

    If lngRow Then
        If ignoreHiddenCells And wks.rows(lngRow).Hidden Then
            lngRowEnd = nextVisibleRow(wks, lngRow, Excel.xlUp)
            If lngRowEnd Then GoTo Retry
        Else
            lastNonEmptyRow = lngRow
        End If
    End If

'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function

Public Function lastNonEmptyColumn(wks As Excel.Worksheet, _
                   Optional startRow As Long, Optional startCol As Long, _
                   Optional endRow As Long, Optional endCol As Long, _
                   Optional ignoreHiddenCells As Boolean = False) As Long
    Const METHOD_NAME As String = "lastNonEmptyColumn"
    '------------------------------------------------------------------------------------------------------
    Dim xls As Excel.Application
    Dim lngCol As Long
    Dim lngColStart As Long
    Dim lngColEnd As Long
    Dim lngNonBlanks As Long
    Dim lngRowStart As Long
    Dim lngRowEnd As Long
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------

    If Not F.Sheets.isValid(wks) Then GoTo IllegalSheetException
    Set xls = wks.parent.Application
    
    If startCol > 0 And startCol <= wks.columns.count Then lngColStart = startCol Else lngColStart = 1
    If endCol > 0 And endCol <= wks.columns.count Then lngColEnd = endCol Else _
                                                                        lngColEnd = wks.columns.count
    If startRow > 0 And startRow <= wks.rows.count Then lngRowStart = startRow Else lngRowStart = 1
    If endRow > 0 And endRow <= wks.rows.count Then lngRowEnd = endRow Else lngRowEnd = wks.rows.count

Retry:
    lngCol = 1
    
    Do
        Set rng = wks.Range(wks.Cells(lngRowStart, lngCol), wks.Cells(lngRowEnd, lngColEnd))
        lngNonBlanks = xls.WorksheetFunction.CountA(rng)
                                                                                                        
        If lngNonBlanks Then
                                                                                                      
            If lngCol = lngColEnd Then Exit Do
            lngColStart = lngCol
            lngCol = lngCol + ((lngColEnd - lngCol + 1) / 2)
                                                                                                      
        Else
            lngColEnd = lngCol - 1
            lngCol = lngColStart
                                                                                                      
            If lngColStart > lngColEnd Then
                lngCol = 0
                Exit Do
            End If
                                                                                                      
        End If
                                                                                                        
    Loop

    If lngCol Then
        If ignoreHiddenCells And wks.columns(lngCol).Hidden Then
            lngColEnd = nextVisibleColumn(wks, lngCol, Excel.xlToLeft)
            If lngColEnd Then GoTo Retry
        Else
            lastNonEmptyColumn = lngCol
        End If
    End If


'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function

Public Function nextVisibleRow(wks As Excel.Worksheet, initialRow As Long, _
                                                                          direction As XlDirection) As Long
    Dim intOffset As Integer
    '------------------------------------------------------------------------------------------------------
                
    If Not F.Sheets.isValid(wks) Then GoTo IllegalSheetException
                
    Select Case direction
        Case Excel.xlUp:    intOffset = -1
        Case Excel.xlDown:  intOffset = 1
        Case Else
            nextVisibleRow = initialRow
            GoTo ExitPoint
    End Select

    nextVisibleRow = initialRow
    Do
        nextVisibleRow = nextVisibleRow + intOffset
        If Not wks.rows(nextVisibleRow).Hidden Then Exit Do
    Loop

'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function

Public Function nextVisibleColumn(wks As Excel.Worksheet, initialCol As Long, _
                                          direction As XlDirection) As Long
    Dim intOffset As Integer
    '------------------------------------------------------------------------------------------------------
                
    If Not F.Sheets.isValid(wks) Then GoTo IllegalSheetException
                
    Select Case direction
        Case Excel.xlToLeft:    intOffset = -1
        Case Excel.xlToRight:   intOffset = 1
        Case Else
            nextVisibleColumn = initialCol
            GoTo ExitPoint
    End Select


    nextVisibleColumn = initialCol
    Do
        nextVisibleColumn = nextVisibleColumn + intOffset
        If Not wks.columns(nextVisibleColumn).Hidden Then Exit Do
    Loop


'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function



Public Function intersect(ParamArray ranges() As Variant) As Excel.Range
    Dim arrRanges As Variant:               arrRanges = ranges
    Dim varRange As Variant
    Dim i As Integer
    Dim isInitialized As Boolean
    Dim rng As Excel.Range
    Dim subrng As Excel.Range
    Dim app As Excel.Application
    '------------------------------------------------------------------------------------------------------

    If F.arrays.isDefined(arrRanges) Then
        For i = LBound(arrRanges) To UBound(arrRanges)
            Call F.Variables.assign(varRange, arrRanges(i))

            If isNotEmptyRange(varRange) Then
                If Not isInitialized Then
                    Set rng = varRange
                    Set app = rng.Application
                    isInitialized = True
                Else
                    Set rng = app.intersect(rng, varRange)
                    If rng Is Nothing Then Exit For
                End If
            Else
                Set rng = Nothing
                Exit For
            End If

        Next i
    End If

    Set intersect = rng

End Function

Private Function isNotEmptyRange(rng As Variant) As Boolean
    Const METHOD_NAME As String = "isNotEmptyRange"
    '------------------------------------------------------------------------------------------------------
    Dim subrng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    If VBA.IsObject(rng) Then
        If Not rng Is Nothing Then
            If TypeOf rng Is Excel.Range Then
                Set subrng = rng
                isNotEmptyRange = F.ranges.isValid(subrng)
            End If
        End If
    End If
End Function


Public Function getNextEmptyCell(cell As Excel.Range, direction As Long) As Excel.Range
    Dim wks As Excel.Worksheet
    Dim i As Long
    Dim stepSize As Long
    Dim firstIndex As Long
    Dim lastIndex As Long
    Dim isHorizontalSearch As Boolean
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    If F.ranges.isValid(cell) Then
        Set wks = cell.parent
        
        If direction = xlToLeft Or direction = xlUp Then
            stepSize = -1
            lastIndex = 1
        ElseIf direction = xlToRight Or direction = xlDown Then
            stepSize = 1
            lastIndex = VBA.IIf(direction = xlToRight, wks.columns.count, wks.rows.count)
        End If
        
        isHorizontalSearch = (direction = xlToRight Or direction = xlToLeft)
        firstIndex = VBA.IIf(isHorizontalSearch, cell.column, cell.row)
        
        For i = firstIndex To lastIndex Step stepSize
            If isHorizontalSearch Then
                Set rng = wks.Cells(cell.row, i)
            Else
                Set rng = wks.Cells(i, cell.column)
            End If
            If VBA.IsEmpty(rng.value) Then
                Set getNextEmptyCell = rng
                Exit For
            End If
        Next i
        
    End If
    
End Function



Public Function insertColumns(startCell As Excel.Range, columnsCounter As Long) As Excel.Range
    Dim wks As Excel.Worksheet
    Dim rng As Excel.Range
    Dim column As Long
    '------------------------------------------------------------------------------------------------------
    
    Set wks = startCell.parent
    column = startCell.column
    
    Set rng = startCell.Resize(1, columnsCounter).EntireColumn
    Call rng.insert(Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove)
    
    With wks
        Set insertColumns = .Range(.Cells(1, column), .Cells(1, column + columnsCounter - 1)).EntireColumn
    End With
    
End Function

Public Function insertRows(sheet As Excel.Worksheet, topRow As Long, rowsCounter As Long, _
                           Optional Shift As Variant = xlDown, Optional clearFormat As Boolean = True) As Excel.Range
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    With sheet
        Set rng = .Range(.Cells(topRow, 1), .Cells(topRow + rowsCounter - 1, 1)).EntireRow
        Call rng.insert(VBA.IIf(Shift = xlDown, -4121, -4162))
        Set rng = .Range(.Cells(topRow, 1), .Cells(topRow + rowsCounter - 1, 1)).EntireRow
        If clearFormat Then Call rng.clear
    End With
    Set insertRows = rng
End Function


Public Function getRangeLastColumn(rng As Excel.Range) As Long
    Dim subrng As Excel.Range
    Dim lastRow As Long
    '------------------------------------------------------------------------------------------------------
    On Error Resume Next
    If F.ranges.isValid(rng) Then
        For Each subrng In rng.Areas
            With subrng
                lastRow = .column + .columns.count - 1
                If lastRow > getRangeLastColumn Then getRangeLastColumn = lastRow
            End With
        Next subrng
    End If
End Function

Public Function getRangeLastRow(rng As Excel.Range) As Long
    On Error Resume Next
    With rng
        getRangeLastRow = .row + .rows.count - 1
    End With
End Function







'borderDict - JSON object width definition of borders.
' Template:
'            {
'                "right": {
'                    "color": 16777215,
'                    "weight": "thin",
'                    "style": "continuous"
'                },
'                "left": {...},
'                "top": {...},
'                "bottom": {...},
'                "inside-horizontal": {...},
'                "inside-vertical": {...}
'            }
Public Sub applyBorders(rng As Excel.Range, bordersDict As Scripting.Dictionary, _
                            Optional skipOutsideBorders As Boolean = False, Optional skipInsideBorders As Boolean = False)
    Const BORDER_WEIGHT_TAG As String = "weight"
    Const BORDER_COLOR_TAG As String = "color"
    Const BORDER_STYLE_TAG As String = "style"
    '------------------------------------------------------------------------------------------------------
    Dim varKey As Variant
    Dim borderIndex As XlBordersIndex
    Dim borderProps As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If F.Dictionaries.isNonEmptyDictionary(bordersDict) Then
        For Each varKey In bordersDict.Keys
            borderIndex = convertBorderIndexNameToEnum(VBA.CStr(varKey))
            If (isInsideBorder(borderIndex) And Not skipInsideBorders) Or (isOutsideBorder(borderIndex) And Not skipOutsideBorders) Then
                Set borderProps = bordersDict.item(varKey)
                Call applyBorder(rng, borderIndex, borderProps)
                Call applyAdjacentBorder(rng, borderIndex, borderProps)
            End If
        Next varKey
    End If
    
End Sub

Public Sub applySpecificBorder(rng As Excel.Range, bordersDict As Scripting.Dictionary, borderIndex As XlBordersIndex)
    Const BORDER_WEIGHT_TAG As String = "weight"
    Const BORDER_COLOR_TAG As String = "color"
    Const BORDER_STYLE_TAG As String = "style"
    '------------------------------------------------------------------------------------------------------
    Dim borderIndexName As String
    Dim borderProps As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If F.Dictionaries.isNonEmptyDictionary(bordersDict) Then
        borderIndexName = convertBorderIndexEnumToName(borderIndex)
        Set borderProps = F.Dictionaries.getObject(bordersDict, borderIndexName)
        If Not borderProps Is Nothing Then
            Call applyBorder(rng, borderIndex, borderProps)
        Else
            rng.Borders(borderIndex).LineStyle = xlLineStyleNone
        End If
    End If
    
End Sub

Private Sub applyBorder(rng As Excel.Range, borderIndex As XlBordersIndex, borderProps As Scripting.Dictionary)
    Const BORDER_WEIGHT_TAG As String = "weight"
    Const BORDER_COLOR_TAG As String = "color"
    Const BORDER_STYLE_TAG As String = "style"
    '------------------------------------------------------------------------------------------------------
    With rng.Borders(borderIndex)
        .Weight = convertBorderWeightNameToEnum(F.Dictionaries.getItem(borderProps, BORDER_WEIGHT_TAG))
        .LineStyle = convertBorderStyleNameToEnum(F.Dictionaries.getItem(borderProps, BORDER_STYLE_TAG))
        .color = F.Dictionaries.getItem(borderProps, BORDER_COLOR_TAG)
    End With
End Sub

Private Sub applyAdjacentBorder(rng As Excel.Range, borderIndex As XlBordersIndex, borderProps As Scripting.Dictionary)
    Dim rngAdjacent As Excel.Range
    Dim oppositeBorderIndex As XlBordersIndex
    '------------------------------------------------------------------------------------------------------
    
    If borderIndex = xlEdgeBottom Then
        oppositeBorderIndex = xlEdgeTop
        Set rngAdjacent = rng.rows(rng.rows.count).Offset(1)
    ElseIf borderIndex = xlEdgeTop Then
        oppositeBorderIndex = xlEdgeBottom
        Set rngAdjacent = rng.rows(1).Offset(-1)
    ElseIf borderIndex = xlEdgeLeft Then
        oppositeBorderIndex = xlEdgeRight
        Set rngAdjacent = rng.columns(1).Offset(, -1)
    ElseIf borderIndex = xlEdgeRight Then
        oppositeBorderIndex = xlEdgeLeft
        Set rngAdjacent = rng.columns(rng.columns.count).Offset(, 1)
    End If
    
    If F.ranges.isValid(rngAdjacent) And oppositeBorderIndex > 0 Then
        Call applyBorder(rngAdjacent, oppositeBorderIndex, borderProps)
    End If
    
End Sub




'formatDict - JSON object width definition of borders.
' Template:
'            {
'                "right": {
'                    "color": 16777215,
'                    "weight": "thin",
'                    "style": "continuous"
'                },
'                "left": {...},
'                "top": {...},
'                "bottom": {...},
'                "inside-horizontal": {...},
'                "inside-vertical": {...}
'            }
Public Function applyFormatCondition(rng As Excel.Range, formatProps As Scripting.Dictionary) As Excel.FormatCondition
    Dim formatType As XlFormatConditionType:                formatType = xlExpression
    Dim operator As XlFormatConditionOperator:              operator = xlEqual
    Dim setFirstPriority As Boolean
    Dim stopIfTrue As Boolean
    Dim formula As Variant
    Dim backColor As Variant
    Dim FontColor As Variant
    Dim fontSize As Variant
    Dim FontBold As Variant
    '------------------------------------------------------------------------------------------------------
    
    If F.ranges.isValid(rng) Then
        
        '[Load properties]
        Call F.json.assignPropertyFromJSON(formatProps, formatType, "type", F.Reflection.getFunctionFullName("convertFormatConditionTypeToEnum", Excel.ThisWorkbook))
        Call F.json.assignPropertyFromJSON(formatProps, operator, "operator", F.Reflection.getFunctionFullName("convertFormatConditionOperatorToEnum", Excel.ThisWorkbook))
        Call F.json.assignPropertyFromJSON(formatProps, setFirstPriority, "setFirstPriority")
        Call F.json.assignPropertyFromJSON(formatProps, stopIfTrue, "stopIfTrue")
        Call F.json.assignPropertyFromJSON(formatProps, formula, "formula")
        Call F.json.assignPropertyFromJSON(formatProps, backColor, "interior.color")
        Call F.json.assignPropertyFromJSON(formatProps, FontColor, "font.color")
        Call F.json.assignPropertyFromJSON(formatProps, fontSize, "font.size")
        Call F.json.assignPropertyFromJSON(formatProps, FontBold, "font.bold")
        
        '[Apply properties to range]
        Set applyFormatCondition = rng.FormatConditions.Add(formatType, operator, formula)
        With applyFormatCondition
            If setFirstPriority Then Call .setFirstPriority
            .stopIfTrue = stopIfTrue
            If Not VBA.IsEmpty(backColor) Then .Interior.color = backColor
            If Not VBA.IsEmpty(FontColor) Then .Font.color = FontColor
            If Not VBA.IsEmpty(FontBold) Then .Font.Bold = FontBold
            If Not VBA.IsEmpty(fontSize) Then .Font.size = fontSize
        End With
        
    End If
    
End Function

Public Sub applyFormatFromDictionary(rng As Excel.Range, properties As Scripting.Dictionary)
    Dim dict As Scripting.Dictionary
    Dim key As Variant
    Dim prop As UIProp
    Dim dictBorders As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If F.ranges.isValid(rng) Then
        Set dict = F.Dictionaries.flatDictionary(properties)
        If F.Dictionaries.isNonEmptyDictionary(dict) Then
            For Each key In dict.Keys
                Set prop = UIProps.getPropByAlias(VBA.CStr(key))
                If Not prop Is Nothing Then
                    Call applyFormatForSingleProperty(rng, prop, prop.getConvertedValue(dict.item(key)))
                End If
            Next key
        End If
        
        '[Borders]
        On Error Resume Next
        Set dictBorders = UIProps.Borders.getValueFromPropertiesDictionary(properties)
        If Not DEV_MODE Then On Error GoTo ErrHandler
        If F.Dictionaries.isNonEmptyDictionary(dictBorders) Then Call applyBorders(rng, dictBorders)
        
    End If
    
ErrHandler:
    
End Sub

Public Sub applyFormatForSingleProperty(rng As Excel.Range, prop As UIProp, value As Variant)
    
    Select Case prop.getId
        Case UIProps.width.getId:                       rng.EntireColumn.ColumnWidth = value
        Case UIProps.height.getId:                      rng.EntireRow.rowHeight = value
        Case UIProps.BackgroundColor.getId:             rng.Interior.color = value
        Case UIProps.fontSize.getId:                    rng.Font.size = value
        Case UIProps.FontColor.getId:                   rng.Font.color = value
        Case UIProps.FontFamily.getId:                  rng.Font.name = value
        Case UIProps.FontBold.getId:                    rng.Font.Bold = value
        Case UIProps.FontItalic.getId:                  rng.Font.Italic = value
        Case UIProps.VerticalAlignment.getId:           rng.VerticalAlignment = value
        Case UIProps.HorizontalAlignment.getId:         rng.HorizontalAlignment = value
        Case UIProps.IndentLevel.getId:                 rng.IndentLevel = value
        Case UIProps.WrapText.getId:                    rng.WrapText = value
        Case UIProps.AutoFit.getId:                     Call rng.AutoFit
        Case UIProps.NumberFormat.getId:                rng.NumberFormat = value
    End Select
    
End Sub





Public Sub applyRowAutoFit(rng As Excel.Range, Optional minHeight As Single)
    Dim heights() As Single
    Dim height As Single
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    Dim subrng As Excel.Range
    Dim dictHeights As Scripting.Dictionary
    Dim colRows As VBA.Collection
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    If F.ranges.isValid(rng) Then
        If rng.Areas.count > 1 Then
            For Each subrng In rng.Areas
                Call applyRowAutoFit(subrng, minHeight)
            Next subrng
        Else
            Set dictHeights = F.Dictionaries.Create(False)
            
            With rng
                ReDim heights(1 To .rows.count)
                For i = LBound(heights) To UBound(heights)
                    heights(i) = .rows(i).rowHeight
                    Set subrng = F.ranges.join(subrng, .rows(i))
                Next i
                
                Call subrng.EntireRow.AutoFit
                
                For i = LBound(heights) To UBound(heights)
                    height = F.Math.getMax(heights(i), minHeight)
                    If .rows(i).rowHeight < height Then
                        Set colRows = F.Dictionaries.getObject(dictHeights, height)
                        If colRows Is Nothing Then
                            Set colRows = New VBA.Collection
                            Call dictHeights.Add(height, colRows)
                        End If
                        Call colRows.Add(rng.rows(i).EntireRow)
                    End If
                Next i
                
                
                
                'Restore minimum height
                For Each key In dictHeights.Keys
                    Set subrng = F.ranges.join(dictHeights.item(key))
                    If F.ranges.isValid(subrng) Then
                        subrng.EntireRow.rowHeight = VBA.CSng(key)
                    End If
                Next key
                
                'Hide rows that were hidden before
                For i = LBound(heights) To UBound(heights)
                    If heights(i) = 0 Then
                        .rows(i).EntireRow.Hidden = True
                    End If
                Next i
                
                
            End With
        
        End If
    End If
    
End Sub

Public Sub clearSheetFormatConditions(wks As Excel.Worksheet, Optional protectionPassword As String)
    Dim isProtected As Boolean:             isProtected = wks.ProtectContents
    Dim count As Long
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If isProtected Then Call wks.Unprotect(protectionPassword)
    
    With wks.Cells.FormatConditions
        count = .count
        For i = count To 1 Step -1
            Call .item(i).delete
        Next i
    End With
    
    If isProtected Then Call wks.Protect(protectionPassword)
    
End Sub

Public Function getCurrentSelectionAddress(Optional xls As Excel.Application) As String
    If xls Is Nothing Then Set xls = Excel.Application
    On Error Resume Next
    getCurrentSelectionAddress = xls.selection.address
    On Error GoTo 0
End Function



Public Sub applyColors(dict As Scripting.Dictionary)
    Dim key As Variant
    Dim subcol As VBA.Collection
    Dim rng As Excel.Range
    Dim color As Long
    '------------------------------------------------------------------------------------------------------
    
    For Each key In dict.Keys
        Set subcol = F.Dictionaries.getObject(dict, key)
        If F.Collections.isNonEmptyCollection(subcol) Then
            Set rng = F.ranges.join(subcol)
            If F.ranges.isValid(rng) Then
                color = VBA.CLng(key)
                With rng.Interior
                    If F.colors.isValidRgbColor(color) Then
                        rng.Interior.color = VBA.CLng(key)
                    Else
                        .pattern = xlNone
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End If
                End With
            End If
        End If
    Next key
    
End Sub



Public Function getVisibleSelection(Optional ByVal xls As Excel.Application) As Excel.Range
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    If xls Is Nothing Then Set xls = Excel.Application
    
    On Error Resume Next
    Set rng = xls.selection
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If isValid(rng) Then
        Set getVisibleSelection = rng.SpecialCells(xlCellTypeVisible)
    End If
    
ErrHandler:
    
End Function



Public Function haveCommonParts(rng1 As Excel.Range, rng2 As Excel.Range) As Boolean
    If Not isValid(rng1) Then
        haveCommonParts = False
    ElseIf Not isValid(rng2) Then
        haveCommonParts = False
    ElseIf Not rng1.parent Is rng2.parent Then
        haveCommonParts = False
    ElseIf rng1.row >= rng2.row + rng2.rows.count Then
        haveCommonParts = False
    ElseIf rng2.row >= rng1.row + rng1.rows.count Then
        haveCommonParts = False
    ElseIf rng1.column >= rng2.column + rng2.columns.count Then
        haveCommonParts = False
    ElseIf rng2.column >= rng1.column + rng1.columns.count Then
        haveCommonParts = False
    Else
        haveCommonParts = True
    End If
End Function


Public Function toString(value As Object) As String
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    If value Is Nothing Then
        toString = "Nothing"
    ElseIf Not TypeOf value Is Excel.Range Then
        toString = "Invalid data type"
    Else
        Set rng = value
        If F.ranges.isValid(rng) Then
            toString = "{Excel.Range} " & rng.address & " (" & rng.parent.name & ")"
        Else
            toString = "Invalid range"
        End If
    End If
End Function







''Author: https://stackoverflow.com/questions/21580795/subtracting-ranges-in-vba-excel
'Public Function subtract(rFirst As Range, rSecond As Range) As Range
'    Dim rInter As Range
'    Dim rReturn As Range
'    Dim rArea As Range
'
'    Set rInter = intersect(rFirst, rSecond)
'    Set mrBuild = Nothing
'
'    If rInter Is Nothing Then 'no overlap
'        Set rReturn = rFirst
'    ElseIf rInter.address = rFirst.address Then 'total overlap
'        Set rReturn = Nothing
'    Else 'partial overlap
'        For Each rArea In rFirst.Areas
'            Set mrBuild = BuildRange(rArea, rInter) 'recursive
'        Next rArea
'        Set rReturn = mrBuild
'    End If
'
'    Set subtract = rReturn
'
'End Function
'
'
'Private Function BuildRange(rArea As Range, rInter As Range, Optional mrBuild As Range = Nothing) As Range
'' Recursive function for SubtractRanges()
''
'' Subtracts rInter from rArea and adds the result to mrBuild
''
'    Dim app As Excel.Application:                       Set app = rArea.Application
'    Dim rLeft As Range, rRight As Range
'    Dim rTop As Range, rBottom As Range
'    Dim rInterSub As Range
'    Dim GoByColumns As Boolean
'
'    Set rInterSub = intersect(rArea, rInter)
'    If rInterSub Is Nothing Then 'no overlap
'        If mrBuild Is Nothing Then
'            Set mrBuild = rArea
'        Else
'            Set mrBuild = app.Union(mrBuild, rArea)
'        End If
'    ElseIf Not rInterSub.address = rArea.address Then 'some overlap
'        If Not rArea.Cells.CountLarge = 1 Then 'just in case there is only one cell for some impossible reason
'
'            ' Decide whether to go by columns or by rows
'            ' (helps when subtracting whole rows/columns)
'            If Not rInterSub.columns.count = rArea.columns.count And _
'            ((Not rInterSub.Cells.CountLarge = 1 And _
'            (rInterSub.rows.count > rInterSub.columns.count _
'            And rArea.columns.count > 1) Or (rInterSub.rows.count = 1 _
'            And Not rArea.columns.count = 1)) Or _
'            (rInterSub.Cells.CountLarge = 1 _
'            And rArea.columns.count > rArea.rows.count)) Then
'                    GoByColumns = True
'            Else
'                    GoByColumns = False
'            End If
'
'            If Not GoByColumns Then
'                Set rTop = rArea.Resize(rArea.rows.count \ 2) 'split the range top to bottom
'                Set rBottom = rArea.Resize(rArea.rows.count - rTop.rows.count).Offset(rTop.rows.count)
'                Set mrBuild = BuildRange(rTop, rInterSub, mrBuild) 'rerun it
'                Set mrBuild = BuildRange(rBottom, rInterSub, mrBuild)
'            Else
'                Set rLeft = rArea.Resize(, rArea.columns.count \ 2) 'split the range left to right
'                Set rRight = rArea.Resize(, rArea.columns.count - rLeft.columns.count).Offset(, rLeft.columns.count)
'                Set mrBuild = BuildRange(rLeft, rInterSub, mrBuild) 'rerun it
'                Set mrBuild = BuildRange(rRight, rInterSub, mrBuild)
'            End If
'        End If
'    End If
'
'    Set BuildRange = mrBuild
'End Function
