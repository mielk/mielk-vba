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

Public Function getNamedRange(wks As Excel.Worksheet, rangeName As String) As Excel.Range
    On Error Resume Next
    Set getNamedRange = wks.Range(rangeName)
End Function

Public Sub closeParentWorkbook(wks As Excel.Worksheet, Optional saveChanges As Boolean = False)
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    If isValid(wks) Then
        Set wkb = wks.parent
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

Public Function getContent(filePath As String, Optional sheetIndex As Long = 1, Optional xls As Excel.Application) As Variant
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
    
    Set wkb = F.Books.open_(filePath, True, xls, wasOpen:=fileWasOpen)
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
    Set getSheetsDictionary = F.Dictionaries.Create
    If F.Books.isValid(wkb) Then
        For Each wks In wkb.Worksheets
            If includeHiddenSheets Or wks.visible = xlSheetVisible Then
                Call getSheetsDictionary.Add(wks.name, wks)
            End If
        Next wks
    End If
End Function

Public Function getSheet(wkb As Excel.Workbook, sheetName As String, Optional useAsRegex As Boolean = False) As Excel.Worksheet
    If F.Books.isValid(wkb) Then
        If useAsRegex Then
            Set getSheet = getSheetByNameRegex(wkb, sheetName)
        Else
            On Error Resume Next
            Set getSheet = wkb.Worksheets(sheetName)
        End If
    End If
End Function

Private Function getSheetByNameRegex(wkb As Excel.Workbook, regexPattern As String) As Excel.Worksheet
    Static regex As Object
    Dim wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    
    If regex Is Nothing Then
        Set regex = F.regex.Create(regexPattern)
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
    If Not F.Books.isValid(wkb) Then Call Exceptions.InvalidBook.raise
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
