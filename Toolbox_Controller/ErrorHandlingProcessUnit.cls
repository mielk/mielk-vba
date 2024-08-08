VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandlingProcessUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProcessItem

Private Const CLASS_NAME As String = "ErrorHandlingProcessUnit"
'[PROGRESS BAR MESSAGES] ----------------------------------------------------------------------------------
Private Const PB_LOADING_ENTRY_POINTS As String = "ErrorHandling.ProgressBar.LoadingEntryPoints"
Private Const PB_LOADING_FILE_STRUCTURE As String = "ErrorHandling.ProgressBar.LoadingFileStructure"
Private Const PB_OPENING_PROJECT_FILE As String = "ErrorHandling.ProgressBar.OpeningProjectFile"
Private Const PB_PROCESSING_MODULE As String = "ErrorHandling.ProgressBar.ProcessingModule"
Private Const PB_CLOSING_FILE As String = "ErrorHandling.ProgressBar.ClosingProjectFile"
'----------------------------------------------------------------------------------------------------------
Private pParent As ErrorHandlingManager
Private pProcessor As ItemsProcessor
Private pStatus As ItemProcessStatusEnum
Private pMessage As String
'----------------------------------------------------------------------------------------------------------
Private pFilepath As String
Private pProject As EVbaProject
Private pEntryPoints As Scripting.Dictionary
Private pSummary As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------



'[Setters]
Public Function setParent(value As ErrorHandlingManager) As ErrorHandlingProcessUnit
    Set setParent = Me
    Set pParent = value
End Function

Public Function setItemsProcessor(value As ItemsProcessor) As ErrorHandlingProcessUnit
    Set setItemsProcessor = Me
    Set pProcessor = value
End Function

Public Function setFilePath(value As String) As ErrorHandlingProcessUnit
    Set setFilePath = Me
    pFilepath = value
End Function


'[Getters]
Public Function getCaption() As String
    getCaption = pFilepath
End Function

Public Function getMessage() As String
    getMessage = pMessage
End Function

Public Function getStatus() As Long
    getStatus = pStatus
End Function

Public Function getFileName() As String
    getFileName = f.files.getFileName(pFilepath)
End Function

Public Function getSummary() As Scripting.Dictionary
    Set getSummary = pSummary
End Function



'[API]
Public Sub process()
    Dim xls As Excel.Application
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    Dim modulesCounter As Long
    Dim module As EVbaModule
    Dim modulesDone As Long
    Dim startProgress As Single
    Dim progress As Single
    Dim endProgress As Single
    '------------------------------------------------------------------------------------------------------
    Dim summary As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    
    '[Create dictionary of entry points to differentiate between normal methods and opening ones.] --|
    Call pProcessor.triggerDetailStatusUpdate(Msg.getText(PB_LOADING_ENTRY_POINTS), 0)              '|
    Set pEntryPoints = pParent.getEntryPointsForProject(f.files.getFileName(pFilepath, False))      '|
    '------------------------------------------------------------------------------------------------|

    '[Open workbook to create EVbaProject object.] --------------------------------------------------|
    Call pProcessor.triggerDetailStatusUpdate(Msg.getText(PB_OPENING_PROJECT_FILE), 0.02)           '|
    Set xls = pParent.getExcel                                                                      '|
    Set wkb = xls.Workbooks.Open(pFilepath, , False)                                                '|
    Set pProject = createVbaProjectFromBook(wkb)                                                    '|
    '------------------------------------------------------------------------------------------------|
    
    
    If Not pProject Is Nothing Then
        Call createFileBackup
        
        modulesCounter = pProject.countComponents
        startProgress = 0.1
        endProgress = 0.97
        
        Set pSummary = f.dictionaries.create(False)
        
        For Each module In pProject.getComponents
        
            progress = (modulesDone / modulesCounter) * (endProgress - startProgress) + startProgress
            Call pProcessor.triggerDetailStatusUpdate( _
                        f.Strings.Format(Msg.getText(PB_PROCESSING_MODULE), module.getName), progress)
                        
            '[Append error handling to the current module and save summary dictionaries.] -----------|
            Set summary = module.appendErrorHandling(pEntryPoints)                                  '|
            Call f.dictionaries.AddItem(pSummary, module.getName, summary)                          '|
            '----------------------------------------------------------------------------------------|
            
            modulesDone = modulesDone + 1
            
        Next module
        
        
        '[Saving changes in project file.] ----------------------------------------------------------|
        Call pProcessor.triggerDetailStatusUpdate(Msg.getText(PB_CLOSING_FILE), 0.97)               '|
        Call closeBook(wkb)                                                                         '|
        '--------------------------------------------------------------------------------------------|
        
    End If
        
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then
        pStatus = ItemProcessStatus_Error
        pMessage = errDescription
    Else
        pStatus = ItemProcessStatus_Correct
        pMessage = "OK"
    End If
    Exit Sub

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    GoTo ExitPoint

End Sub

Private Sub createFileBackup()
    Const REGEX_PATTERN As String = "\\(?=[^\\]*$)"
    '------------------------------------------------------------------------------------------------------
    Dim backupFilePath As String
    '------------------------------------------------------------------------------------------------------
    backupFilePath = f.files.uniqueFilePath(f.regex.Replace(pFilepath, REGEX_PATTERN, "\________"))
    Call fso.CopyFile(pFilepath, backupFilePath)
End Sub

Private Function isEntryPoint(className As String, methodName As String) As Boolean
    Dim name As String
    '------------------------------------------------------------------------------------------------------
    
    name = className & "." & methodName
    
    If pEntryPoints Is Nothing Then
        isEntryPoint = False
    Else
        isEntryPoint = pEntryPoints.Exists(name)
    End If
    
End Function


Private Sub closeBook(wkb As Excel.Workbook)
'    Const TIMEOUT As Long = 10
'    '------------------------------------------------------------------------------------------------------
'    Dim startTime As Double
'    '------------------------------------------------------------------------------------------------------
'
'    startTime = VBA.Timer
'
'    Do
'        Call VBA.Err.Clear
'        On Error Resume Next
'        Call wkb.Close(False)
'
'        If VBA.Timer - startTime > TIMEOUT Then
'            Call VBA.Err.Raise(1, Description:="Error while trying to close workbook " & wkb.FullName)
'        End If
'
'    Loop While Err.Number
    
    Call wkb.Save
    'Call wkb.Close(True)
    
End Sub



'[Implementation of IProcessItem]

Private Function IProcessItem_getCaption() As String
    IProcessItem_getCaption = getCaption
End Function

Private Function IProcessItem_getMessage() As Variant
    IProcessItem_getMessage = getMessage
End Function

Private Function IProcessItem_getName() As String
    IProcessItem_getName = getCaption
End Function

Private Function IProcessItem_getStatus() As Long
    IProcessItem_getStatus = getStatus
End Function

Private Sub IProcessItem_preview()
    Stop
End Sub

Private Sub IProcessItem_process()
    Call process
End Sub

Private Function IProcessItem_setItemsProcessor(value As MielkApp.ItemsProcessor) As MielkApp.IProcessItem
    Set IProcessItem_setItemsProcessor = setItemsProcessor(value)
End Function

Private Function IProcessItem_setParams(params As Scripting.IDictionary) As MielkApp.IProcessItem
    Stop
End Function
