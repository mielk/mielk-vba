VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FSheets"
    '----------------------------------------------------------------------------------------------------------


'***** WITHOUT UNIT TESTS *****
Public Function IsValid(wks As Excel.Worksheet) As Boolean
    Dim strSheetName As String
    '------------------------------------------------------------------------------------------------------
    On Error Resume Next
    strSheetName = wks.name
    IsValid = VBA.Len(strSheetName)
End Function

Public Function getNamedRange(wks As Excel.Worksheet, rangeName As String) As Excel.Range
    On Error Resume Next
    Set getNamedRange = wks.Range(rangeName)
End Function

Public Sub closeParentWorkbook(wks As Excel.Worksheet, Optional saveChanges As Boolean = False)
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    If IsValid(wks) Then
        Set wkb = wks.Parent
        Call wkb.Close(saveChanges)
    End If
End Sub

Public Function getNonEmptyRegion(wks As Excel.Worksheet) As Excel.Range
    Dim firstRow As Long
    Dim firstColumn As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    '------------------------------------------------------------------------------------------------------

    firstRow = 1
    firstColumn = 1
    lastRow = f.ranges.getLastNonEmptyRow(wks)
    lastColumn = f.ranges.getLastNonEmptyColumn(wks)

    If lastRow > 0 And lastColumn > 0 Then
        With wks
            Set getNonEmptyRegion = .Range(.cells(firstRow, firstColumn), .cells(lastRow, lastColumn))
        End With
    End If

End Function

Public Function getDataArray(wks As Excel.Worksheet, Optional startFromA1 As Boolean = False) As Variant
    Dim firstRow As Long
    Dim firstColumn As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    '------------------------------------------------------------------------------------------------------

    If startFromA1 Then
        firstRow = 1
        firstColumn = 1
    Else
        firstRow = f.ranges.getFirstNonEmptyRow(wks)
        firstColumn = f.ranges.firstNonEmptyColumn(wks)
    End If

    lastRow = f.ranges.getLastNonEmptyRow(wks)
    lastColumn = f.ranges.getLastNonEmptyColumn(wks)

    With wks
        getDataArray = .Range(.cells(firstRow, firstColumn), .cells(lastRow, lastColumn))
    End With

End Function

Public Function getContent(filepath As String, Optional sheetIndex As Long = 1, Optional xls As Excel.Application) As Variant
    Dim wkb As Excel.Workbook
    Dim wks As Excel.Worksheet
    Dim isSeparateExcelInstance As Boolean
    Dim fileWasOpen As Boolean
    '------------------------------------------------------------------------------------------------------
    
    If xls Is Nothing Then
        Set xls = New Excel.Application
        xls.visible = False
        isSeparateExcelInstance = True
    End If
    
    Set wkb = f.Books.open_(filepath, True, xls, wasOpen:=fileWasOpen)
    Set wks = wkb.Worksheets(sheetIndex)
    getContent = wks.cells(1, 1).CurrentRegion
    
    If Not fileWasOpen Then
        Call wkb.Close(False)
        Set wkb = Nothing
    End If
    
    If isSeparateExcelInstance Then
        Call xls.Quit
        Set xls = Nothing
    End If
    
End Function

Public Function isShapeValid(shape As Excel.shape) As Boolean
    Dim baseName As String: baseName = f.Crypto.createUUID
    Dim name As String
    '----------------------------------------------------------------------------------------------------------
    
    name = baseName
    On Error Resume Next
    name = shape.DrawingObject.name
    On Error GoTo 0
    
    isShapeValid = (name <> baseName)
    
End Function

Public Function getSheetsDictionary(wkb As Excel.Workbook, Optional includeHiddenSheets As Boolean = False) As Scripting.Dictionary
    Dim wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    Set getSheetsDictionary = f.dictionaries.Create
    If f.Books.IsValid(wkb) Then
        For Each wks In wkb.Worksheets
            If includeHiddenSheets Or wks.visible = xlSheetVisible Then
                Call getSheetsDictionary.Add(wks.name, wks)
            End If
        Next wks
    End If
End Function

Public Function getSheet(wkb As Excel.Workbook, sheetName As String, Optional useAsRegex As Boolean = False) As Excel.Worksheet
    If f.Books.IsValid(wkb) Then
        If useAsRegex Then
            Set getSheet = getSheetByNameRegex(wkb, sheetName)
        Else
            On Error Resume Next
            Set getSheet = wkb.Worksheets(sheetName)
        End If
    End If
End Function

Public Function getSheetsByName(wkb As Excel.Workbook, regex As String, Optional ignoreHiddenSheets As Boolean = True) As VBA.Collection
    Dim wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    
    Set getSheetsByName = New VBA.Collection
    If f.Books.IsValid(wkb) Then
        For Each wks In wkb.Worksheets
            If ignoreHiddenSheets And wks.visible <> xlSheetVisible Then
                'skip this worksheet
            ElseIf f.regex.checkIfMatch(wks.name, regex) Then
                Call getSheetsByName.Add(wks)
            End If
        Next wks
    End If
    
End Function

Private Function getSheetByNameRegex(wkb As Excel.Workbook, regexPattern As String) As Excel.Worksheet
    Static regex As Object
    Dim wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    
    If regex Is Nothing Then
        Set regex = f.regex.Create(regexPattern)
    Else
        regex.pattern = regexPattern
    End If
    
    For Each wks In wkb.Worksheets
        If regex.test(wks.name) Then
            Set getSheetByNameRegex = wks
            Exit For
        End If
    Next wks
    
End Function

Public Sub delete(wkb As Excel.Workbook, sheetName As String)
    Dim displayAlerts As Boolean
    '------------------------------------------------------------------------------------------------------
    
    displayAlerts = wkb.Application.displayAlerts
    wkb.Application.displayAlerts = False
    
    On Error Resume Next
    Call wkb.Worksheets(sheetName).delete
    On Error GoTo 0
    
    wkb.Application.displayAlerts = displayAlerts
    
End Sub

Public Sub deleteAllExcept(wkb As Excel.Workbook, sheetName As String, Optional useRegex As Boolean = False)
    Dim wks As Excel.Worksheet
    Dim toBeDeleted As VBA.Collection
    Dim name As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set toBeDeleted = New VBA.Collection
    
    If f.Books.IsValid(wkb) Then
        For Each wks In wkb.Worksheets
            
            If Not useRegex Then
                If Not f.strings.compareStrings(wks.name, sheetName) Then
                    Call toBeDeleted.Add(wks.name)
                End If
            Else
                Stop
            End If
            
        Next wks
    End If
    
    If f.Collections.isNonEmptyCollection(toBeDeleted) Then
        For Each name In toBeDeleted
            Call delete(wkb, VBA.CStr(name))
        Next name
    End If
    
End Sub


Public Function getUniqueName(wkb As Excel.Workbook, name As String) As String
    Const MAX_LENGTH As Integer = 31
    '------------------------------------------------------------------------------------------------------
    Dim wks As Excel.Worksheet
    Dim strTempName As String
    Dim intIterator As Integer
    Dim intCharsCounter As Integer
    '------------------------------------------------------------------------------------------------------


    'First check if the given name is legal sheet name. -------------------------------------------------|
    strTempName = getLegalSheetName(name)                                                               '|
    getUniqueName = strTempName                                                                         '|
    '----------------------------------------------------------------------------------------------------|


    'Check if the book to be checked is not closed. -----------------------------------------------------|
    If Not f.Books.IsValid(wkb) Then Call Exceptions.InvalidBook.raise
    '----------------------------------------------------------------------------------------------------|


    Do
        Set wks = getSheet(wkb, getUniqueName)
        If wks Is Nothing Then
            Exit Do
        Else
        
            intIterator = intIterator + 1
            getUniqueName = strTempName & "_" & intIterator
                                                                                                    
            'Check if the name with the number appended is not too long
            '(the maximum length is defined by constant MAX_LENGTH).
            intCharsCounter = VBA.Len(getUniqueName)
            If intCharsCounter > MAX_LENGTH Then
                getUniqueName = VBA.Left$(strTempName, _
                   VBA.Len(strTempName) - intCharsCounter + MAX_LENGTH) & "_" & intIterator
            End If
            
        End If
              
    Loop Until wks Is Nothing
    
End Function


Public Function getLegalSheetName(name As String) As String
    Const ILLEGAL_CHARS As String = ":?/\*[]"
    '------------------------------------------------------------------------------------------------------
    Dim intChar As Integer
    Dim strChar As String
    '------------------------------------------------------------------------------------------------------

    For intChar = 1 To VBA.Len(name)
        strChar = VBA.mid$(name, intChar, 1)
        If VBA.InStr(1, ILLEGAL_CHARS, strChar) = 0 Then
            getLegalSheetName = getLegalSheetName & strChar
        End If
    Next intChar

    Select Case VBA.Len(getLegalSheetName)
        Case Is > 31
            getLegalSheetName = VBA.Left$(getLegalSheetName, 31)
        Case 0
            getLegalSheetName = "_"
    End Select

End Function



Public Function getLastSheet(wkb As Excel.Workbook) As Excel.Worksheet
    Set getLastSheet = wkb.Worksheets(wkb.Worksheets.count)
End Function



Public Function getPivotTableByName(wks As Excel.Worksheet, name As String) As Excel.PivotTable
    Dim pivot As Excel.PivotTable
    '------------------------------------------------------------------------------------------------------
    
    For Each pivot In wks.PivotTables
        If f.strings.compareStrings(pivot.name, name) Then
            Set getPivotTableByName = pivot
            Exit For
        End If
    Next pivot
    
End Function




'[Shapes]
Public Function getShape(wks As Excel.Worksheet, shapeName As String) As Excel.shape
    On Error Resume Next
    Set getShape = wks.shapes(shapeName)
End Function

Public Sub removeShape(wks As Excel.Worksheet, shapeName As String)
    Dim shp As Excel.shape
    '------------------------------------------------------------------------------------------------------
    Set shp = getShape(wks, shapeName)
    If Not shp Is Nothing Then
        Call shp.delete
    End If
End Sub

Public Sub removeShapesByCell(wks As Excel.Worksheet, Optional row As Long, Optional column As Long)
    Dim shp As Excel.shape
    Dim cell As Excel.Range
    '------------------------------------------------------------------------------------------------------
    If f.sheets.IsValid(wks) Then
        For Each shp In wks.shapes
            Set cell = shp.TopLeftCell
            If row > 0 And cell.row <> row Then
            ElseIf column > 0 And cell.column <> column Then
            Else
                Call shp.delete
            End If
        Next shp
    End If
End Sub




Public Sub clearColumnsStartingFrom(wks As Excel.Worksheet, initialColumn As Long, _
                                        Optional ignoreErrors As Boolean = True, _
                                        Optional keepFormatting As Boolean = True)
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    If Not f.sheets.IsValid(wks) Then
        If Not ignoreErrors Then Call Exceptions.IllegalSheet.raise
    ElseIf initialColumn <= 0 Then
        If Not ignoreErrors Then Call Exceptions.IllegalValue.raiseWithCustomMessage("Initial column index must be greater than 0")
    ElseIf initialColumn >= wks.columns.count Then
        If Not ignoreErrors Then Call Exceptions.IllegalValue.raiseWithCustomMessage("Initial column index cannot be greater than the total number of columns in the sheet")
    Else
        With wks
            Set rng = .Range(.cells(1, initialColumn), .cells(1, .columns.count)).EntireColumn
            If keepFormatting Then
                Call rng.ClearContents
            Else
                Call rng.clear
            End If
        End With
    End If
    
End Sub


Public Function copy(originalSheet As Excel.Worksheet, destinationBook As Excel.Workbook, _
                                                        Optional sheetName As String) As Excel.Worksheet
    Dim bookSheets As Scripting.Dictionary
    Dim wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    
    If Not f.Books.IsValid(destinationBook) Then
        Call Exceptions.InvalidBook.raise
    ElseIf Not f.sheets.IsValid(originalSheet) Then
        Call Exceptions.InvalidSheet.raise
    Else
        Set bookSheets = f.Books.getSheetsDictionary(destinationBook)
        Call originalSheet.copy(After:=f.Books.getLastSheet(destinationBook))
        
        'Find a worksheet that name is not in the [bookSheets] dictionary - it is a worksheet
        'that has just been added.
        For Each wks In destinationBook.Worksheets
            If Not bookSheets.exists(wks.name) Then
                Set copy = wks
                If VBA.Len(sheetName) Then
                    copy.name = f.sheets.getUniqueName(destinationBook, sheetName)
                End If
                Exit For
            End If
        Next wks
        
    End If
    
End Function

Public Function getFormulasRange(wks As Excel.Worksheet) As Excel.Range
    On Error Resume Next
    Set getFormulasRange = wks.cells.SpecialCells(xlCellTypeFormulas)
End Function
