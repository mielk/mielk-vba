VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FBooks"
    '----------------------------------------------------------------------------------------------------------



'***** WITHOUT UNIT TESTS *****
Public Function IsValid(wkb As Excel.Workbook) As Boolean
    Dim strBookName As String
    '------------------------------------------------------------------------------------------------------
    On Error Resume Next
    strBookName = wkb.name
    IsValid = VBA.Len(strBookName)
End Function

Public Function open_(filepath As String, Optional readOnly As Boolean = False, _
                         Optional excelInstance As Excel.Application, _
                         Optional createIfNotExists As Boolean = False, _
                         Optional forceReopen As Boolean = False, _
                         Optional ByRef wasOpen As Boolean, _
                         Optional ByVal password As String, _
                         Optional ByVal passwordWrite As String, _
                         Optional ignoreReadOnlyRecommended As Boolean = True) As Excel.Workbook
    Dim xls As Excel.Application
    Dim bAlerts As Boolean
    '------------------------------------------------------------------------------------------------------
    
    
    If excelInstance Is Nothing Then
        Set xls = Excel.ThisWorkbook.Application
    Else
        Set xls = excelInstance
    End If


    With xls
        bAlerts = .displayAlerts
        .displayAlerts = False
    End With


    If isOpen(filepath, xls) Then
        wasOpen = True
        Set open_ = xls.Workbooks(f.files.getFileName(filepath))
        If forceReopen Then
            Call open_.Close(False)
            Set open_ = Nothing
        End If
    Else
        wasOpen = False
    End If
    
    
    '[Passwords]
    If VBA.Len(passwordWrite) = 0 Then passwordWrite = password
    
    If open_ Is Nothing Then
        If f.files.fileExists(filepath) Then
            On Error Resume Next
            Set open_ = xls.Workbooks.Open(filepath, _
                                                readOnly:=readOnly, _
                                                updateLinks:=False, _
                                                password:=password, _
                                                WriteResPassword:=passwordWrite, _
                                                ignoreReadOnlyRecommended:=ignoreReadOnlyRecommended)
            
            If open_ Is Nothing And readOnly = False Then
                Set open_ = xls.Workbooks.Open(filepath, _
                                                readOnly:=readOnly, _
                                                updateLinks:=False, _
                                                password:=password, _
                                                WriteResPassword:=passwordWrite, _
                                                ignoreReadOnlyRecommended:=ignoreReadOnlyRecommended, _
                                                CorruptLoad:=xlRepairFile)
            End If
            On Error GoTo 0
        End If
    End If
  
  
    On Error Resume Next
    xls.displayAlerts = bAlerts
    On Error GoTo 0
  
  
End Function

Public Function isOpen(filepath As String, Optional excelInstance As Excel.Application) As Boolean
    Dim xls As Excel.Application
    Dim fileName As String:             fileName = f.files.getFileName(filepath)
    Dim wkb As Excel.Workbook
    Dim addIn As Excel.addIn
    '------------------------------------------------------------------------------------------------------
    
    If excelInstance Is Nothing Then
        Set xls = Excel.ThisWorkbook.Application
    Else
        Set xls = excelInstance
    End If
    
    
    On Error Resume Next
    Set wkb = xls.Workbooks(fileName)
    On Error GoTo 0
    
    '[Search add-ins]
    For Each addIn In xls.AddIns
        If f.strings.compareStrings(addIn.fullName, filepath) Then
            Set wkb = addIn
            Exit For
        End If
    Next addIn

    
    If wkb Is Nothing Then
        isOpen = False
    Else
        isOpen = f.strings.compareStrings(wkb.fullName, filepath)
    End If
    
    
End Function

Public Function addNew(sheetsNumber As Integer, excelInstance As Excel.Application, _
                                                        ParamArray names() As Variant) As Excel.Workbook
    Dim xls As Excel.Application
    Dim defaultSheetsNumber As Integer
    Dim wks As Excel.Worksheet
    Dim varName As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
        
    If excelInstance Is Nothing Then Set xls = Excel.Application Else Set xls = excelInstance
    With xls
        defaultSheetsNumber = .SheetsInNewWorkbook
        .SheetsInNewWorkbook = sheetsNumber
    End With
    
    Set addNew = xls.Workbooks.Add
    
    For Each varName In names
        i = i + 1
        If i <= addNew.Worksheets.count Then
            Set wks = addNew.Worksheets(i)
            wks.name = f.sheets.getLegalSheetName(VBA.CStr(varName))
        End If
    Next varName
    
    xls.SheetsInNewWorkbook = defaultSheetsNumber
    
End Function


Public Function getLastSheet(wkb As Excel.Workbook) As Excel.Worksheet
    Set getLastSheet = wkb.Worksheets(wkb.Worksheets.count)
End Function

Public Sub openInNewExcel(filepath As String, Optional readOnly As Boolean = False)
    Dim xls As Excel.Application
    Dim cursor As XlMousePointer:                       cursor = xlDefault
    '------------------------------------------------------------------------------------------------------
    
    If f.files.fileExists(filepath) Then
        With Excel.Application
            cursor = .cursor
            .cursor = xlWait
        End With
        
        Set xls = VBA.CreateObject("Excel.Application")
        With xls
            .visible = True
            Call .Workbooks.Open(filepath, readOnly:=readOnly)
            Call BringWindowToTop(.hWnd)
        End With
    Else
        Call Exceptions.FileNotFound.raise
    End If
    
ExitPoint:
    Excel.Application.cursor = cursor
    
End Sub




Public Function getSheetsDictionary(wkb As Excel.Workbook) As Scripting.Dictionary
    Dim wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    Set getSheetsDictionary = f.dictionaries.Create(False)
    For Each wks In wkb.Worksheets
        Call f.dictionaries.addItem(getSheetsDictionary, wks.name, wks)
    Next wks
End Function

Public Function getSheets(wkb As Excel.Workbook, Optional sheetNames As Variant) As VBA.Collection
    Dim wks As Excel.Worksheet
    Dim dictSheetNames As Scripting.Dictionary
    Dim varSheetName As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set getSheets = New VBA.Collection
    Set dictSheetNames = f.dictionaries.createFrom(sheetNames, duplicateBehaviour_Skip)
    
    If f.dictionaries.isNonEmptyDictionary(dictSheetNames) Then
        For Each varSheetName In dictSheetNames.Keys
            Set wks = f.sheets.getSheet(wkb, VBA.CStr(varSheetName))
            If Not wks Is Nothing Then
                Call getSheets.Add(wks)
            End If
        Next varSheetName
    Else
        For Each wks In wkb.Worksheets
            Call getSheets.Add(wks)
        Next wks
    End If
    
End Function

Public Sub closeBook(book As Excel.Workbook, Optional saveChanges As Boolean = True)
    If IsValid(book) Then
        Call book.Close(saveChanges)
    End If
End Sub

Public Sub updateLinksFromDictionary(wkb As Excel.Workbook, dictLinks As Scripting.Dictionary, _
                                                    Optional sheetNames As Variant)
    Dim relinker As ExcelRelinker
    '------------------------------------------------------------------------------------------------------
    Set relinker = New ExcelRelinker
    Call relinker.updateLinks(wkb, dictLinks, sheetNames)
End Sub

Public Sub updateLinks(wkb As Excel.Workbook, ByVal oldPath As String, ByVal newPath As String, _
                                                    Optional sheetNames As Variant)
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Set dict = f.dictionaries.createSingleEntryDict(oldPath, newPath)
    Call updateLinksFromDictionary(wkb, dict, sheetNames)
End Sub
