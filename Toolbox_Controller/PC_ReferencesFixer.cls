VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PC_ReferencesFixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProcessItem
Implements IProjectCreatorProcessUnit

Private Const CLASS_NAME As String = "PC_ReferencesFixer"
'[Config constants] --------------------------------------------------------------------------------------
Private Const LIB_FILE_PREFIX As String = "lib"
Private Const PROJECT_FILE_PREFIX As String = "project"
'[Errors] -------------------------------------------------------------------------------------------------
Private Const ERROR_GENERAL As String = "CreatingNewProject.Errors.FixingFilesReferences"
Private Const ERROR_OPENING_VBA_PROJECT As String = "CreatingNewProject.Errors.OpeningVbaProject"
Private Const ERROR_FILE_NOT_FOUND As String = "CreatingNewProject.Errors.ReferenceFileNotFound"
Private Const ERROR_REMOVING_REFERENCE As String = "CreatingNewProject.Errors.RemovingSingleReference"
Private Const ERROR_ADDING_REFERENCE As String = "CreatingNewProject.Errors.AddingSingleReference"
'[Status] -------------------------------------------------------------------------------------------------
Private Const SUCCESS_MESSAGE As String = "CreatingNewProject.Messages.FixingReferencesSuccess"
Private Const SEMISUCCEsSS_MESSAGE As String = "CreatingNewProject.Messages.FixingReferencesSemiSuccess"
Private Const FAILURE_MESSAGE As String = "CreatingNewProject.Messages.FixingReferencesFailure"
'[Progress bar] -------------------------------------------------------------------------------------------
Private Const ITEMS_PROCESSOR_CAPTION As String = "CreatingNewProject.Steps.FixingReferences"
Private Const PBAR_FIXING_REFERENCES As String = "CreatingNewProject.Actions.FixingSingleFileReferences"
'----------------------------------------------------------------------------------------------------------
Private pParent As ProjectCreator
'----------------------------------------------------------------------------------------------------------
Private pWarnings As VBA.Collection
Private pErrors As VBA.Collection
Private pSuccessful As VBA.Collection
'----------------------------------------------------------------------------------------------------------



'[SETTERS]
Public Sub setParent(value As ProjectCreator)
    Set pParent = value
End Sub



'[GETTERS]

Public Function getCaption() As String
    getCaption = Msg.getText(ITEMS_PROCESSOR_CAPTION)
End Function

Public Function getMessage() As String
    If pParent.isAborted Then
        getMessage = Msg.getText(FAILED_BECAUSE_OF_PREDECESSORS)
    Else
        Select Case getStatus
            Case ItemProcessStatus_Correct
                getMessage = f.Strings.Format(Msg.getText(SUCCESS_MESSAGE), f.Collections.toString(pSuccessful, , VBA.vbCrLf))
            Case ItemProcessStatus_Error
                getMessage = f.Collections.toString(pErrors, , VBA.vbCrLf)
        End Select
    End If
End Function

Public Function getStatus() As Long
    If pErrors Is Nothing Then
        getStatus = ItemProcessStatus_Unknown
    ElseIf pErrors.Count Then
        getStatus = ItemProcessStatus_Error
    Else
        getStatus = ItemProcessStatus_Correct
    End If
End Function

Private Sub reset()
    Set pWarnings = New VBA.Collection
    Set pErrors = New VBA.Collection
    Set pSuccessful = New VBA.Collection
End Sub





'[API]
Public Sub process()
    Dim xls As Excel.Application
    Dim project As EProject
    Dim files As VBA.Collection
    Dim varFilePath As Variant
    Dim dictReferences As Scripting.Dictionary
    '[For Items Processor updates] ------------------------------------------------------------------------
    Dim caption As String
    Dim counter As Long
    '------------------------------------------------------------------------------------------------------
    Dim validation As ValidationObject
    '------------------------------------------------------------------------------------------------------

    On Error GoTo errHandler
    
    Call reset
    If pParent.isAborted Then Call VBA.Err.Raise(1, , Msg.getText(FAILED_BECAUSE_OF_PREDECESSORS))
    
    With pParent
        Set project = .getProject
        Set dictReferences = adjustReferencesDictionary(.getReferencesDictionary)
        Set xls = createBackgroundExcelInstance
    End With
    
    Set files = f.files.getFolderFiles(project.getFolderPath, False, MACRO_FILE_REGEX_PATTERN)
    For Each varFilePath In files
        
        '[Updating ItemsProcessor progress bar] ---------------------------------------------|
        caption = f.Strings.Format(Msg.getText(PBAR_FIXING_REFERENCES), varFilePath)        '|
        Call pParent.getProcessor.triggerDetailStatusUpdate(caption, counter / files.Count) '|
        counter = counter + 1                                                               '|
        '------------------------------------------------------------------------------------|
        
        '------------------------------------------------------------------------------------|
        Set validation = fixReferencesInFile(xls, VBA.CStr(varFilePath), dictReferences)    '|
        If validation.getStatus = ItemProcessStatus_Error Then                              '|
            Call pErrors.Add(validation.getMessage)                                         '|
        End If                                                                              '|
        '------------------------------------------------------------------------------------|
        
    Next varFilePath
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    
    If Not xls Is Nothing Then Call xls.Quit
    
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call pErrors.Add(f.Strings.Format(Msg.getText(ERROR_GENERAL), errDescription, errNumber))
    
End Sub

Private Function adjustReferencesDictionary(dictReferences As Scripting.Dictionary) As Scripting.Dictionary
    Dim key As Variant
    Dim fileName As String
    Dim dictFileReferences As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set adjustReferencesDictionary = f.dictionaries.create(False)
    
    For Each key In dictReferences.Keys
        fileName = f.Strings.replaceByDictionary(VBA.CStr(key), getReferenceTypeToNameDictionary)
        Set dictFileReferences = adjustSingleFileReferencesDictionary( _
                                        f.cast.toCollection(dictReferences.item(key)))
        Call f.dictionaries.AddItem(adjustReferencesDictionary, fileName, dictFileReferences)
    Next key

End Function

Private Function getReferenceTypeToNameDictionary() As Scripting.Dictionary
    Static instance As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        Set instance = f.dictionaries.createWithItems(False, _
                                KeyValue(LIB_FILE_PREFIX, LIBRARY_NAME), _
                                KeyValue(PROJECT_FILE_PREFIX, pParent.getProject.getCodeName))
    End If
    
    Set getReferenceTypeToNameDictionary = instance
    
End Function

Private Function getReferenceTypeToFolderPathDictionary() As Scripting.Dictionary
    Static instance As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If instance Is Nothing Then
        With pParent.getProject
            Set instance = f.dictionaries.createWithItems(False, _
                                    KeyValue(LIB_FILE_PREFIX, .getLibFolderPath), _
                                    KeyValue(PROJECT_FILE_PREFIX, .getFolderPath))
        End With
    End If
    
    Set getReferenceTypeToFolderPathDictionary = instance
    
End Function

Private Function adjustSingleFileReferencesDictionary(references As VBA.Collection) As Scripting.Dictionary
    Dim varRef As Variant
    Dim refName As String
    Dim refType As String
    Dim refPath As String
    '------------------------------------------------------------------------------------------------------
    
    Set adjustSingleFileReferencesDictionary = f.dictionaries.create(False)
    
    For Each varRef In references
        refType = f.Strings.substring(VBA.CStr(varRef), vbNullString, "-")
        refName = f.Strings.replaceByDictionary(VBA.CStr(varRef), getReferenceTypeToNameDictionary)
        refPath = f.files.concatenatePath( _
                                f.dictionaries.GetItem(getReferenceTypeToFolderPathDictionary, refType), _
                                refName & EXTENSION_EXCEL_ADDIN)
        Call f.dictionaries.AddItem(adjustSingleFileReferencesDictionary, refName, refPath)
    Next varRef
    
End Function

Private Function fixReferencesInFile(xls As Excel.Application, filepath As String, _
                                            dictReferences As Scripting.Dictionary) As ValidationObject
    Const FAILURE_MESSAGE_PATTERN As String = "CreatingNewProject.Errors.FixingFileReferences"
    '------------------------------------------------------------------------------------------------------
    Dim wkb As Excel.Workbook
    Dim project As VBIDE.VBProject
    Dim ref As VBIDE.Reference
    Dim validation As ValidationObject
    Dim message As String
    '------------------------------------------------------------------------------------------------------
    Dim addInName As String
    Dim configReferences As Scripting.Dictionary
    Dim projectReferences As Scripting.Dictionary
    Dim varKey As Variant
    '------------------------------------------------------------------------------------------------------

    '[Get config for this specific file (by file path). -----------------------------|
    addInName = f.regex.getFirstGroupMatch(filepath, EXCEL_ADDIN_NAME_PATTERN)      '|
    Set configReferences = f.dictionaries.GetObject(dictReferences, addInName)      '|
    '--------------------------------------------------------------------------------|
    
    '[Open project and load references into dictionary ------------------------------|
    '(by file name, no extenstion).                                                 '|
    Set wkb = f.Books.open_(filepath, False, xls)                                   '|
    Set project = f.Developer.getVbProject(wkb)                                     '|
    Set projectReferences = getProjectReferencesDictionary(project)                 '|
    '--------------------------------------------------------------------------------|

    If project Is Nothing Then
        message = f.Strings.Format(Msg.getText(ERROR_OPENING_VBA_PROJECT), filepath)
        Call VBA.Err.Raise(1, , message)
    Else
        If f.dictionaries.isNonEmptyDictionary(configReferences) Then
            For Each varKey In configReferences.Keys
                Set validation = fixSingleReference(project, VBA.CStr(varKey), _
                                        f.dictionaries.GetItem(configReferences, varKey), _
                                        f.dictionaries.GetItem(projectReferences, varKey))
                If validation.getStatus = ItemProcessStatus_Correct Then
                    Call pSuccessful.Add(validation.getMessage)
                ElseIf validation.getStatus = ItemProcessStatus_Error Then
                    Call pErrors.Add(validation.getMessage)
                End If
            Next varKey
        End If
    End If

    Call wkb.Save
    'Call wkb.Close(False)

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Set fixReferencesInFile = f.create.ValidationObject(ItemProcessStatus_Correct)
    Exit Function

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    message = f.Strings.Format(Msg.getText(FAILURE_MESSAGE_PATTERN), filepath, errDescription, errNumber)
    Set fixReferencesInFile = f.create.ValidationObject(ItemProcessStatus_Error, message)

End Function

Private Function getProjectReferencesDictionary(project As VBIDE.VBProject) As Scripting.Dictionary
    Dim ref As VBIDE.Reference
    Dim refName As String
    Dim refPath As String
    '----------------------------------------------------------------------------------------------------------
    
    Set getProjectReferencesDictionary = f.dictionaries.create(False)
    
    For Each ref In project.references
        refName = f.Strings.replaceByDictionary( _
                        f.regex.getFirstGroupMatch(ref.FullPath, EXCEL_ADDIN_NAME_PATTERN), _
                        pParent.getProject.getReplacements)
        refPath = ref.FullPath
        Call f.dictionaries.AddItem(getProjectReferencesDictionary, refName, refPath)
    Next ref
    
End Function

Private Function fixSingleReference(project As VBIDE.VBProject, name As String, _
                                        desiredPath As String, currentPath As String) As ValidationObject
    Const SUCCESS_MESSAGE_PATTERN As String = "{0}::{1} [{2}]"
    Const FAILURE_MESSAGE_PATTERN As String = "CreatingNewProject.Errors.FixingSingleReference"
    Dim errNumber As Long, errDescription As String
    '------------------------------------------------------------------------------------------------------
    Dim ref As VBIDE.Reference
    Dim message As String
    Dim counter As Long
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    Set fixSingleReference = f.create.ValidationObject(ItemProcessStatus_Skipped)
    
    If Not f.files.FileExists(desiredPath) Then
        message = f.Strings.Format(Msg.getText(ERROR_FILE_NOT_FOUND), desiredPath, name)
        Set fixSingleReference = f.create.ValidationObject(ItemProcessStatus_Error, message)
    ElseIf f.Strings.compareStrings(desiredPath, currentPath) Then
        ' Nothing to change in this case
    Else
    
        '[Remove previous reference, it directs to the wrong address. -----------------------------------|
        If VBA.Len(currentPath) Then                                                                    '|
            counter = project.references.Count                                                          '|
            Set ref = getReferenceByPath(project, currentPath)                                          '|
            On Error Resume Next                                                                        '|
            Call project.references.Remove(ref)                                                         '|
            On Error GoTo errHandler                                                                    '|
            If project.references.Count <> counter - 1 Then                                             '|
                errNumber = VBA.Err.Number: errDescription = VBA.Err.Description                        '|
                message = f.Strings.Format(Msg.getText(ERROR_REMOVING_REFERENCE), _
                                            name, desiredPath, project.name, errDescription, errNumber) '|
                Set fixSingleReference = f.create.ValidationObject(ItemProcessStatus_Error, message)    '|
            End If                                                                                      '|
        End If                                                                                          '|
        '------------------------------------------------------------------------------------------------|
        
        '[Adding reference] -----------------------------------------------------------------------------|
        counter = project.references.Count                                                              '|
        On Error Resume Next                                                                            '|
        Call project.references.AddFromFile(desiredPath)                                                '|
        On Error GoTo errHandler                                                                        '|
        If project.references.Count <> counter + 1 Then                                                 '|
            errNumber = VBA.Err.Number: errDescription = VBA.Err.Description                            '|
            message = f.Strings.Format(Msg.getText(ERROR_ADDING_REFERENCE), _
                                            name, desiredPath, project.name, errDescription, errNumber) '|
            Set fixSingleReference = f.create.ValidationObject(ItemProcessStatus_Error, message)        '|
        End If                                                                                          '|
        '------------------------------------------------------------------------------------------------|
        
        message = f.Strings.Format(SUCCESS_MESSAGE_PATTERN, project.name, name, desiredPath)
        Set fixSingleReference = f.create.ValidationObject(ItemProcessStatus_Correct, message)
        
    End If
    

'==========================================================================================================
ExitPoint:
    Exit Function

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    message = f.Strings.Format(Msg.getText(FAILURE_MESSAGE_PATTERN), name, project.fileName, _
                                                                                errDescription, errNumber)
    Set fixSingleReference = f.create.ValidationObject(ItemProcessStatus_Error, message)
    
End Function

Private Function getReferenceByPath(project As VBIDE.VBProject, path As String) As VBIDE.Reference
    Dim ref As VBIDE.Reference
    '------------------------------------------------------------------------------------------------------
    For Each ref In project.references
        If ref.FullPath = path Then
            Set getReferenceByPath = ref
            Exit For
        End If
    Next ref
End Function



Private Function createBackgroundExcelInstance() As Excel.Application
    Const ERR_MESSAGE_PATTERN As String = "CreatingNewProject.Errors.CreatingBackgroundExcelInstance"
    '------------------------------------------------------------------------------------------------------

    On Error GoTo errHandler
    Set createBackgroundExcelInstance = VBA.CreateObject(EXCEL_APPLICATION)
    With createBackgroundExcelInstance
        .DisplayAlerts = False
        .AskToUpdateLinks = False
        .Visible = False
    End With

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Function

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    Call VBA.Err.Raise(1, , f.Strings.Format(Msg.getText(ERR_MESSAGE_PATTERN), errDescription, errNumber))

End Function





'[Implementation of IProjectCreatorProcessUnit interface]

Private Sub IProjectCreatorProcessUnit_setParent(value As ProjectCreator)
    Call setParent(value)
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
    'Set IProcessItem_setItemsProcessor = setItemsProcessor(value)
End Function

Private Function IProcessItem_setParams(params As Scripting.IDictionary) As MielkApp.IProcessItem
    Stop
End Function











''[API]
'Public Sub process()
'    Dim xls As Excel.Application
'    Dim processor As ItemsProcessor:        Set processor = pParent.getProcessor
'    Dim project As EProject:                Set project = pParent.getProject
'    Dim files As VBA.Collection
'    Dim varPath As Variant
'    '------------------------------------------------------------------------------------------------------
'
'    On Error GoTo ErrHandler
'
'    Call reset
'
'Call processor.triggerDetailStatusUpdate(Msg.getText(PBAR_CREATING_BACKGROUND_EXCEL_INSTANCE), 0)
'    Set xls = createBackgroundExcelInstance
'
'    If Not xls Is Nothing Then
'Call processor.triggerDetailStatusUpdate(Msg.getText(PBAR_OPENING_PROJECT_FILES), PBARSHARE_CREATING_NEW_EXCEL)
'        Call openFilesInProperOrder(xls)
'    End If
'
'
'
'Call pParent.getProcessor.triggerDetailStatusUpdate(vbNullString, 1)
'
''==========================================================================================================
'ExitPoint:
'    Dim errNumber As Long, errDescription As String
'    Exit Sub
'
''----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
'    Call pErrors.Add(errDescription)
'
'End Sub
'

'
'
'Private Sub openFilesInProperOrder(xls As Excel.Application, _
'                                            Optional startProgress As Double = PBARSHARE_CREATING_NEW_EXCEL)
'    Const ERR_MESSAGE_PATTERN As String = "CreatingNewProject.Errors.OpeningProjectFile"
'    '------------------------------------------------------------------------------------------------------
'    Dim files As VBA.Collection
'    Dim varFile As Variant
'    Dim counter As Long
'    Dim progress As Double:                 progress = startProgress
'    Dim message As String
'    '------------------------------------------------------------------------------------------------------
'
'    On Error GoTo ErrHandler
'
'    files = f.Collections.create(1, 2, 3)
'    For Each varFile In files
'        counter = counter + 1
'        progress = startProgress + (counter / files.Count) * PBARSHARE_OPENING_FILES
'    Next varFile
'
''==========================================================================================================
'ExitPoint:
'    Dim errNumber As Long, errDescription As String
'    Exit Sub
'
''----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
'    Call VBA.Err.Raise(1, , f.Strings.Format(Msg.getText(ERR_MESSAGE_PATTERN), errDescription, errNumber))
'
'End Sub

