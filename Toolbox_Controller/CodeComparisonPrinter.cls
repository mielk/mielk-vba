VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeComparisonPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "CodeComparisonPrinter"
'----------------------------------------------------------------------------------------------------------
Private Const TEMPLATE_FILE_NAME As String = "comparison_template.xlsm"
Private Const SUMMARY_SHEET_NAME As String = "Summary"
Private Const MODULE_TEMPLATE_SHEET_NAME As String = "Module"
Private Const EXCEL_NAME_COMPARE_PROJECT_PATH As String = "COMPARE_PROJECT_PATH"
Private Const EXCEL_NAME_BASE_PROJECT_PATH As String = "BASE_PROJECT_PATH"
Private Const EXCEL_NAME_MODULES_ARRAY_HEADER As String = "MODULES_ARRAY_HEADER"
Private Const EXCEL_NAME_COUNT_EQUAL_MODULES As String = "COUNTER_EQUAL_MODULES"
Private Const EXCEL_NAME_COUNT_DIFFERENT_MODULES As String = "COUNTER_DIFFERENT_MODULES"
Private Const EXCEL_NAME_COUNT_BASE_ONLY_MODULES As String = "COUNTER_BASE_ONLY_MODULES"
Private Const EXCEL_NAME_COUNT_COMPARE_ONLY_MODULES As String = "COUNTER_COMPARE_ONLY_MODULES"
Private Const STATUS_OK As String = "ComparingCode.Status.Ok"
Private Const STATUS_DIFFERENCES As String = "ComparingCode.Status.Differences"
Private Const STATUS_BASE_ONLY As String = "ComparingCode.Status.BaseOnly"
Private Const STATUS_COMPARE_ONLY As String = "ComparingCode.Status.CompareOnly"
'[Module sheets]
Private Const EXCEL_NAME_MODULE_NAME As String = "MODULE_NAME"
Private Const EXCEL_NAME_METHODS_COUNTER_EQUAL As String = "METHODS_COUNTER_EQUAL"
Private Const EXCEL_NAME_METHODS_COUNTER_DIFFERENT As String = "METHODS_COUNTER_DIFFERENT"
Private Const EXCEL_NAME_METHODS_COUNTER_BASE_ONLY As String = "METHODS_COUNTER_BASE_ONLY"
Private Const EXCEL_NAME_METHODS_COUNTER_COMPARE_ONLY As String = "METHODS_COUNTER_COMPARE_ONLY"
Private Const EXCEL_NAME_METHODS_ARRAY_HEADER As String = "METHODS_ARRAY_HEADER"
Private Const METHOD_STATUS_OK As String = "ComparingCode.Status.Methods.Ok"
Private Const METHOD_STATUS_DIFFERENCES As String = "ComparingCode.Status.Methods.Differences"
Private Const METHOD_STATUS_BASE_ONLY As String = "ComparingCode.Status.Methods.BaseOnly"
Private Const METHOD_STATUS_COMPARE_ONLY As String = "ComparingCode.Status.Methods.CompareOnly"
'----------------------------------------------------------------------------------------------------------
Private Const DETAIL_LINE_ROW_HEIGHT As Long = 18
Private Const EQUAL_TAB_COLOR As Long = 5296274
Private Const DIFFERENT_TAB_COLOR As Long = 255
'----------------------------------------------------------------------------------------------------------
Private Const DECLARATIONS_TAG As String = "DECLARATIONS"
Private Const COLUMNS_DELIMITER As String = "@#$%^"
Private Const ROWS_DELIMITER As String = "<<<ROWEND>>>"
Private Const STATUS_MISSING As String = "missing"
Private Const STATUS_MISSING_IN_BASE As String = "missing"
Private Const IGNORE_ERROR_HANDLING As String = True
Private Const IGNORE_DESCRIPTIONS As String = True
'----------------------------------------------------------------------------------------------------------
Private pComparison As EProjectsComparison
Private pBaseProject As EVbaProject
Private pCompareProject As EVbaProject
'----------------------------------------------------------------------------------------------------------
Private pBook As Excel.Workbook
Private pSummary As Excel.Worksheet
Private pSheets As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------



'[GETTERS]
Private Function getBaseProject() As EVbaProject
    If Not pComparison Is Nothing Then
        Set getBaseProject = pComparison.getBaseProject
    End If
End Function

Private Function getCompareProject() As EVbaProject
    If Not pComparison Is Nothing Then
        Set getCompareProject = pComparison.getCompareProject
    End If
End Function



Public Sub run(comparison As EProjectsComparison)
    Set pComparison = comparison
    Set pBaseProject = comparison.getBaseProject
    Set pCompareProject = comparison.getCompareProject
    
    Call createResultBook
Call ProgressBar.setProgress(0.15)
    Call insertMetaData
Call ProgressBar.setProgress(0.2)
    Call insertModulesArray
Call ProgressBar.setProgress(0.25)

Call ProgressBar.openSubProcess(0.25, 1)
Application.ScreenUpdating = False
    Call insertModuleSheets
Application.ScreenUpdating = True
Call ProgressBar.closeCurrentSubProcess
End Sub


Private Sub createResultBook()
    Dim templatePath As String
    '------------------------------------------------------------------------------------------------------
    templatePath = Paths.getTemplateFilePath(TEMPLATE_FILE_NAME)
    Set pBook = Excel.Workbooks.Add(templatePath)
    Set pSummary = pBook.Worksheets(1)
End Sub


Private Sub insertMetaData()
    If f.sheets.isValid(pSummary) Then
        With pSummary
            .name = SUMMARY_SHEET_NAME
            .Range(EXCEL_NAME_BASE_PROJECT_PATH).Value2 = getBaseProject.getPath
            .Range(EXCEL_NAME_COMPARE_PROJECT_PATH).Value2 = getCompareProject.getPath
            
            '[Counters]
            .Range(EXCEL_NAME_COUNT_EQUAL_MODULES).Value2 = pComparison.countModulesByStatus(True)
            .Range(EXCEL_NAME_COUNT_DIFFERENT_MODULES).Value2 = pComparison.countModulesByStatus(False)
            .Range(EXCEL_NAME_COUNT_BASE_ONLY_MODULES).Value2 = pComparison.getBaseOnlyModules.Count
            .Range(EXCEL_NAME_COUNT_COMPARE_ONLY_MODULES).Value2 = pComparison.getCompareOnlyModules.Count
            
        End With
    End If
End Sub


Private Sub insertModulesArray()
    Dim arr As Variant
    Dim rngHeaders As Excel.Range
    Dim rngData As Excel.Range
    Dim rngFormat As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    arr = createSummaryArray
    
    Set rngHeaders = f.sheets.getNamedRange(pSummary, EXCEL_NAME_MODULES_ARRAY_HEADER)
    
    'Apply formatting
    Set rngFormat = rngHeaders.Offset(1)
    Set rngData = rngFormat.Resize(f.arrays.getSize(arr, 2))
    Call rngFormat.Copy(rngData)
    
    'Print data array
    Set rngData = f.Ranges.pasteData(arr, rngData.Cells(1, 1))
    With rngData
        .EntireRow.RowHeight = DETAIL_LINE_ROW_HEIGHT
        .Locked = True
    End With
    
End Sub


Private Function createSummaryArray() As Variant
    Dim totalItems As Long
    Dim arr As Variant
    Dim key As Variant
    Dim comparison As EModulesComparison
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    With pComparison
        totalItems = .getModuleComparisons.Count + .getBaseOnlyModules.Count + .getCompareOnlyModules.Count
    End With
    
    ReDim arr(1 To 3, 1 To totalItems)
    
    '[Found in both projects]
    For Each key In pComparison.getModuleComparisons.Keys
        i = i + 1
        Set comparison = pComparison.getModuleComparisons.item(key)
        arr(1, i) = key
        arr(2, i) = VBA.IIf(comparison.getStatus, 1, 0)
        arr(3, i) = Msg.getText(VBA.IIf(comparison.getStatus, STATUS_OK, STATUS_DIFFERENCES))
    Next key
    
    '[Found in base project only]
    For Each key In pComparison.getBaseOnlyModules.Keys
        i = i + 1
        arr(1, i) = key
        arr(2, i) = 0
        arr(3, i) = Msg.getText(STATUS_BASE_ONLY)
    Next key
    
    '[Found in compare project only]
    For Each key In pComparison.getCompareOnlyModules.Keys
        i = i + 1
        arr(1, i) = key
        arr(2, i) = 0
        arr(3, i) = Msg.getText(STATUS_COMPARE_ONLY)
    Next key
    
    createSummaryArray = arr
    
End Function


Private Sub insertModuleSheets()
    Dim comparisons As Scripting.Dictionary
    Dim key As Variant
    Dim comparison As EModulesComparison
    Dim wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    
    Set comparisons = pComparison.getModuleComparisons
    For Each key In comparisons.Keys
        Set comparison = comparisons.item(key)
        Set wks = insertModuleSheet(comparison)
    Next key
    
    getModuleTemplateSheet.Visible = xlSheetHidden
    
End Sub

Private Function insertModuleSheet(comparison As EModulesComparison) As Excel.Worksheet
    Set insertModuleSheet = f.sheets.Copy(getModuleTemplateSheet, pBook)
    With insertModuleSheet
        .name = comparison.getBaseModule.getName
        .Range(EXCEL_NAME_MODULE_NAME).Value2 = comparison.getBaseModule.getName
        .Range(EXCEL_NAME_METHODS_COUNTER_EQUAL).Value2 = comparison.getEqualMethods.Count
        .Range(EXCEL_NAME_METHODS_COUNTER_DIFFERENT).Value2 = comparison.getDifferentMethods.Count
        .Range(EXCEL_NAME_METHODS_COUNTER_BASE_ONLY).Value2 = comparison.getBaseOnlyMethods.Count
        .Range(EXCEL_NAME_METHODS_COUNTER_COMPARE_ONLY).Value2 = comparison.getCompareOnlyMethods.Count
        Call insertMethodsArray(.Range(EXCEL_NAME_METHODS_ARRAY_HEADER), comparison)
        .Tab.Color = VBA.IIf(comparison.getStatus, EQUAL_TAB_COLOR, DIFFERENT_TAB_COLOR)
    End With
End Function


Private Function getModuleTemplateSheet() As Excel.Worksheet
    Static wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    If wks Is Nothing Then
        Set wks = f.sheets.getSheet(pBook, MODULE_TEMPLATE_SHEET_NAME)
    End If
    Set getModuleTemplateSheet = wks
End Function

Private Sub insertMethodsArray(rng As Excel.Range, comparison As EModulesComparison)
    Dim wks As Excel.Worksheet
    Dim arr As Variant
    Dim rngHeaders As Excel.Range
    Dim rngFormats As Excel.Range
    Dim rngData As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    Set wks = rng.Parent
    Set rngHeaders = wks.Range(EXCEL_NAME_METHODS_ARRAY_HEADER)
    Set rngFormats = rngHeaders.Offset(1)
    
    arr = createMethodsArray(comparison)
    If f.arrays.isDefined(arr) Then
        
        'Apply formatting
        Set rngData = rngFormats.Resize(f.arrays.getSize(arr, 2))
        Call rngFormats.Copy(rngData)
        
        'Print data array
        Set rngData = f.Ranges.pasteData(arr, rngData.Cells(1, 1))
        With rngData
            .EntireRow.RowHeight = DETAIL_LINE_ROW_HEIGHT
            .Locked = True
        End With
        
    Else
        
        On Error Resume Next
        If f.Ranges.isValid(rngFormats) Then
            Call rngFormats.EntireRow.Delete
        End If
        
    End If
    
End Sub


Private Function createMethodsArray(comparison As EModulesComparison) As Variant
    Dim arr As Variant
    Dim totalItems As Long
    Dim row As Long
    '------------------------------------------------------------------------------------------------------
    
    totalItems = comparison.countMethods + 1
    
    If totalItems Then
        ReDim arr(1 To 6, 1 To totalItems)
        Call appendDeclarations(arr, row, comparison)
        Call appendMethods(arr, row, comparison, CodeComparisonStatus_Equal)
        Call appendMethods(arr, row, comparison, CodeComparisonStatus_Different)
        Call appendMethods(arr, row, comparison, CodeComparisonStatus_BaseOnly)
        Call appendMethods(arr, row, comparison, CodeComparisonStatus_CompareOnly)
        createMethodsArray = arr
    End If
End Function


Private Sub appendDeclarations(arr As Variant, ByRef row As Long, comparison As EModulesComparison)
    row = row + 1
    arr(1, row) = DECLARATIONS_TAG
    arr(2, row) = DECLARATIONS_TAG
    arr(3, row) = VBA.IIf(comparison.isDeclarationEqual, 1, 0)
    arr(4, row) = getDescriptionForComparisonStatus(VBA.IIf(comparison.isDeclarationEqual, _
                                                        CodeComparisonStatus_Equal, CodeComparisonStatus_Different))
    If Not comparison.isDeclarationEqual Then
        arr(5, row) = comparison.getBaseModule.getDeclarations
        arr(6, row) = comparison.getCompareModule.getDeclarations
    End If
End Sub




Private Sub appendMethods(arr As Variant, ByRef row As Long, comparison As EModulesComparison, status As CodeComparisonStatusEnum)
    Dim methods As Scripting.Dictionary
    Dim method As EVbaMethod
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set methods = comparison.getMethodsByStatus(status)
    
    For Each key In methods.Keys
        Set method = methods.item(key)
        row = row + 1
        arr(1, row) = method.getName
        arr(2, row) = f.Developer.procTypeToString(method.getProcType)
        arr(3, row) = VBA.IIf(status = CodeComparisonStatus_Equal, 1, 0)
        arr(4, row) = getDescriptionForComparisonStatus(status)
        If status = CodeComparisonStatus_Different Then
            arr(5, row) = method.getBody
            arr(6, row) = getCompareModuleMethod(comparison, VBA.CStr(key))
        End If
    Next key

End Sub

Private Function getCompareModuleMethod(comparison As EModulesComparison, methodKey As String) As String
    Dim module As EVbaModule
    Dim method As EVbaMethod
    '------------------------------------------------------------------------------------------------------
    
    Set module = comparison.getCompareModule
    Set method = module.getMethod(methodKey)
    If Not method Is Nothing Then
        getCompareModuleMethod = method.getBody
    End If
    
End Function

Private Function getDescriptionForComparisonStatus(status As CodeComparisonStatusEnum) As String
    Select Case status
        Case CodeComparisonStatus_Equal:            getDescriptionForComparisonStatus = Msg.getText(METHOD_STATUS_OK)
        Case CodeComparisonStatus_Different:        getDescriptionForComparisonStatus = Msg.getText(METHOD_STATUS_DIFFERENCES)
        Case CodeComparisonStatus_BaseOnly:         getDescriptionForComparisonStatus = Msg.getText(METHOD_STATUS_BASE_ONLY)
        Case CodeComparisonStatus_CompareOnly:      getDescriptionForComparisonStatus = Msg.getText(METHOD_STATUS_COMPARE_ONLY)
    End Select
End Function
