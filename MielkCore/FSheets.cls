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
Public Function isValid(wks As Excel.Worksheet) As Boolean
    Dim strSheetName As String
    '------------------------------------------------------------------------------------------------------
    On Error Resume Next
    strSheetName = wks.name
    isValid = VBA.Len(strSheetName)
End Function

Public Function getNamedRange(wks As Excel.Worksheet, RangeName As String) As Excel.Range
    On Error Resume Next
    Set getNamedRange = wks.Range(RangeName)
End Function

Public Sub closeParentWorkbook(wks As Excel.Worksheet, Optional saveChanges As Boolean = False)
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    If isValid(wks) Then
        Set wkb = wks.Parent
        Call wkb.Close(saveChanges)
    End If
End Sub

Public Function getNonEmptyRegion(wks As Excel.Worksheet) As Variant
    Stop
'    Dim firstRow As Long
'    Dim firstColumn As Long
'    Dim lastRow As Long
'    Dim lastColumn As Long
'    '------------------------------------------------------------------------------------------------------
'
'    firstRow = 1
'    firstColumn = 1
'    lastRow = lastNonEmptyRow(wks)
'    lastColumn = lastNonEmptyColumn(wks)
'
'    If lastRow > 0 And lastColumn > 0 Then
'        With wks
'            getNonEmptyRegion = .Range(.Cells(firstRow, firstColumn), .Cells(lastRow, lastColumn))
'        End With
'    End If
'
End Function

Public Function getDataArray(wks As Excel.Worksheet, Optional startFromA1 As Boolean = False) As Variant
    Stop
'    Dim firstRow As Long
'    Dim firstColumn As Long
'    Dim lastRow As Long
'    Dim lastColumn As Long
'    '------------------------------------------------------------------------------------------------------
'
'    If startFromA1 Then
'        firstRow = 1
'        firstColumn = 1
'    Else
'        firstRow = F.ranges.getFirstNonEmptyRow(wks)
'        firstColumn = firstNonEmptyColumn(wks)
'    End If
'
'    lastRow = lastNonEmptyRow(wks)
'    lastColumn = lastNonEmptyColumn(wks)
'
'    With wks
'        getDataArray = .Range(.Cells(firstRow, firstColumn), .Cells(lastRow, lastColumn))
'    End With
'
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
    getContent = wks.Cells(1, 1).CurrentRegion
    
    If Not fileWasOpen Then
        Call wkb.Close(False)
        Set wkb = Nothing
    End If
    
    If isSeparateExcelInstance Then
        Call xls.Quit
        Set xls = Nothing
    End If
    
End Function



Public Function getSheetsDictionary(wkb As Excel.Workbook, Optional includeHiddenSheets As Boolean = False) As Scripting.Dictionary
    Dim wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    Set getSheetsDictionary = f.dictionaries.Create
    If f.Books.isValid(wkb) Then
        For Each wks In wkb.Worksheets
            If includeHiddenSheets Or wks.visible = xlSheetVisible Then
                Call getSheetsDictionary.add(wks.name, wks)
            End If
        Next wks
    End If
End Function

Public Function getSheet(wkb As Excel.Workbook, SheetName As String, Optional useAsRegex As Boolean = False) As Excel.Worksheet
    If f.Books.isValid(wkb) Then
        If useAsRegex Then
            Set getSheet = getSheetByNameRegex(wkb, SheetName)
        Else
            On Error Resume Next
            Set getSheet = wkb.Worksheets(SheetName)
        End If
    End If
End Function

Public Function getSheetsByName(wkb As Excel.Workbook, regex As String, Optional ignoreHiddenSheets As Boolean = True) As VBA.Collection
    Dim wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    
    Set getSheetsByName = New VBA.Collection
    If f.Books.isValid(wkb) Then
        For Each wks In wkb.Worksheets
            If ignoreHiddenSheets And wks.visible <> xlSheetVisible Then
                'skip this worksheet
            ElseIf f.regex.checkIfMatch(wks.name, regex) Then
                Call getSheetsByName.add(wks)
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

Public Sub delete(wkb As Excel.Workbook, SheetName As String)
    Dim displayAlerts As Boolean
    '------------------------------------------------------------------------------------------------------
    
    displayAlerts = wkb.Application.displayAlerts
    wkb.Application.displayAlerts = False
    
    On Error Resume Next
    Call wkb.Worksheets(SheetName).delete
    On Error GoTo 0
    
    wkb.Application.displayAlerts = displayAlerts

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
    If Not f.Books.isValid(wkb) Then Call Exceptions.InvalidBook.raise
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
                getUniqueName = VBA.left$(strTempName, _
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
            getLegalSheetName = VBA.left$(getLegalSheetName, 31)
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
Public Function getShape(wks As Excel.Worksheet, shapeName As String) As Excel.Shape
    On Error Resume Next
    Set getShape = wks.Shapes(shapeName)
End Function

Public Sub removeShape(wks As Excel.Worksheet, shapeName As String)
    Dim shp As Excel.Shape
    '------------------------------------------------------------------------------------------------------
    Set shp = getShape(wks, shapeName)
    If Not shp Is Nothing Then
        Call shp.delete
    End If
End Sub

Public Sub removeShapesByCell(wks As Excel.Worksheet, Optional row As Long, Optional column As Long)
    Dim shp As Excel.Shape
    Dim cell As Excel.Range
    '------------------------------------------------------------------------------------------------------
    If f.Sheets.isValid(wks) Then
        For Each shp In wks.Shapes
            Set cell = shp.TopLeftCell
            If row > 0 And cell.row <> row Then
            ElseIf column > 0 And cell.column <> column Then
            Else
                Call shp.delete
            End If
        Next shp
    End If
End Sub


