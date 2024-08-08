VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandlingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ErrorHandlingManager"
'----------------------------------------------------------------------------------------------------------
Private Const SELECT_CONFIG_FILE_TITLE As String = "ErrorHandling.WindowSelectConfigFile.Title"
'----------------------------------------------------------------------------------------------------------
Private WithEvents pProcessor As ItemsProcessor
Attribute pProcessor.VB_VarHelpID = -1
Private pItems As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------
Private pConfig As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------
Private pExcel As Excel.Application
'----------------------------------------------------------------------------------------------------------


Public Sub run()
    '[Select file]
    Call createProcessUnits
    Call loadConfig
    
    If f.dictionaries.isNonEmptyDictionary(pItems) Then
        Call setup
        
        Set pProcessor = setupItemsProcessor
        Call pProcessor.run
        
        Call printPostProcessSummary
        
    End If
    
End Sub


Private Sub createProcessUnits()
    Dim files As Variant
    Dim unit As ErrorHandlingProcessUnit
    Dim file As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set pItems = f.dictionaries.create(False)
    Set files = f.files.selectFiles(FileType:="Excel files, *.xlsm; *.xla; *.xlam; *.xlsb")
    'Set files = f.Collections.create("D:\Dropbox\tm\mielk\MLMH\Mediaplan\mp-sources.xlam", _
                                     "D:\Dropbox\tm\mielk\MLMH\Mediaplan\mp-controller.xlam")
    
    If f.Collections.isNonEmptyCollection(files) Then
        For Each file In files
            Set unit = New ErrorHandlingProcessUnit
            With unit
                Call .setParent(Me)
                Call .setFilePath(VBA.CStr(file))
            End With
            Call f.dictionaries.AddItem(pItems, VBA.CStr(file), unit)
        Next file
    End If
    
End Sub

Private Sub loadConfig()
    Dim filepath As String
    '------------------------------------------------------------------------------------------------------
    
    filepath = f.files.selectFile(Title:=Msg.getText(SELECT_CONFIG_FILE_TITLE), FileType:=FILES_PATTERN_JSON)
    'filepath = "D:\Dropbox\tm\mielk\MLMH\Mediaplan\dev\error-handling.json"
    
    If f.files.FileExists(filepath) Then
        Set pConfig = f.json.getJsonObjectFromFile(filepath)
    End If
    
End Sub


Private Sub setup()
    Set pExcel = New Excel.Application
    With pExcel
        .Visible = True
        .AutomationSecurity = msoAutomationSecurityForceDisable
    End With
End Sub



'[Settings]
Friend Function getVbaProjects() As VBA.Collection
    Stop
End Function

Friend Function getExcel() As Excel.Application
    Set getExcel = pExcel
End Function




Private Function setupItemsProcessor() As ItemsProcessor
    Dim viewer As IItemsProcessorViewer
    '------------------------------------------------------------------------------------------------------
    Set viewer = getProcessViewer
    Set setupItemsProcessor = MApp.create.ItemsProcessor(pItems, viewer)
End Function


Private Function getProcessViewer() As WDefaultItemsProcessorViewer
    Const USER_FORM_WIDTH As Single = 500
    '----------------------------------------------------------------------------------------------------------
    
    Set getProcessViewer = MApp.create.DefaultItemsProcessorViewer
    With getProcessViewer
        Call .setHasSummaryPanel(False)
        Call .setHasDetailsProgressBar(True)
        Call .setWidth(USER_FORM_WIDTH)
        Call .setButtonsCaptions(okButtonCaption:="OK")
        Call .setHasCancelButton(False)
        Call .setWindowTitle("Compacting files")
        Call .setStatusDictionary(getStatusViewDictionary)
    End With
    
End Function

Private Function getStatusViewDictionary() As Scripting.Dictionary
    Dim categories As Variant
    '------------------------------------------------------------------------------------------------------
    
    categories = VBA.Array( _
                        ItemProcessStatus_Unknown, _
                        ItemProcessStatus_Correct, _
                        ItemProcessStatus_Error)
                        
    Set getStatusViewDictionary = ItemsProcessor.getDefaultCategories(categories)
    
End Function



'[Help functions]
Friend Function getEntryPointsForProject(projectName As String) As Scripting.Dictionary
    Const NODE_NAME_PATTERN As String = "app.entryPoints.{0}"
    '------------------------------------------------------------------------------------------------------
    Dim nodeName As String
    Dim entryPoints As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    nodeName = f.Strings.Format(NODE_NAME_PATTERN, projectName)
    Set entryPoints = f.dictionaries.getObjectForCombinedKey(pConfig, nodeName)
    Set getEntryPointsForProject = f.Collections.toDictionary(entryPoints)
    
End Function



Private Sub printPostProcessSummary()
    Const SHEET_NAME As String = "summary"
    '------------------------------------------------------------------------------------------------------
    Dim wkb As Excel.Workbook
    Dim wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    Dim items As VBA.Collection
    Dim item As ErrorHandlingProcessUnit
    Dim summary As Scripting.Dictionary
    Dim arr As Variant
    Dim lastRow As Long
    Dim rng As Excel.Range
    Dim dict As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set items = f.dictionaries.toCollection(pItems)
    
    Set wkb = f.Books.addNew(1, Excel.Application, SHEET_NAME)
    Set wks = f.sheets.getSheet(wkb, SHEET_NAME)
    lastRow = 0
    
    For Each item In items
        Set summary = item.getSummary
        
        
        With wks.Cells(lastRow + 1, 1)
            .value = item.getFileName
            .Font.Bold = True
            .Font.size = 14
        End With
        
        lastRow = lastRow + 1
        
        For Each key In summary.Keys
            Set dict = summary.item(key)
            If f.dictionaries.isNonEmptyDictionary(dict) Then
                With wks.Cells(lastRow + 1, 2)
                    .value = key
                    .Font.Bold = True
                    .Font.size = 12
                End With
                
                arr = f.dictionaries.toArray(dict)
                Set rng = f.Ranges.pasteData(arr, wks.Cells(lastRow + 2, 3), False)
                Call rng.EntireColumn.AutoFit
                lastRow = f.Ranges.getRangeLastRow(rng)
                
            End If
        Next key
            
    Next item
    
End Sub


Private Sub pProcessor_AfterConfirm()
    Call pExcel.Quit
    Set pExcel = Nothing
End Sub


