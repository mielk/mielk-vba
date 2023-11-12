VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PC_CodeReplacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProcessItem
Implements IProjectCreatorProcessUnit

Private Const CLASS_NAME As String = "PC_CodeReplacer"
'[Errors] -------------------------------------------------------------------------------------------------
Private Const ERROR_GENERAL As String = "CreatingNewProject.Errors.ApplyingChangesToCode"
Private Const ERROR_OPENING_VBA_PROJECT As String = "CreatingNewProject.Errors.OpeningVbaProject"
Private Const ERROR_RENAMING_PROJECT As String = "CreatingNewProject.Errors.RenamingProject"
Private Const ERROR_CONVERTING_PROJECT As String = "CreatingNewProject.Errors.ConvertingVbProject"
Private Const ERROR_CONVERTING_COMPONENT As String = "CreatingNewProject.Errors.ConvertingVbComponent"
'[Status] -------------------------------------------------------------------------------------------------
Private Const SUCCESS_MESSAGE As String = "CreatingNewProject.Messages.ApplyingChangesToCodeSuccess"
Private Const SEMI_FAILURE_MESSAGE As String = "CreatingNewProject.Messages.ApplyingChangesToCodeSemiFailure"
Private Const FAILURE_MESSAGE As String = "CreatingNewProject.Messages.ApplyingChangesToCodeFailure"
'[Progress bar] -------------------------------------------------------------------------------------------
Private Const ITEMS_PROCESSOR_CAPTION As String = "CreatingNewProject.Steps.ApplyingChangesToCode"
Private Const PBAR_ADJUSTING_CODE As String = "CreatingNewProject.Actions.ApplyingChangesToCode"
'[Other texts] --------------------------------------------------------------------------------------------
Private Const VB_COMPONENT_FULL_NAME_PATTERN As String = "[{0}].{1}"
'----------------------------------------------------------------------------------------------------------
Private pParent As ProjectCreator 'ProcessUnit
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
        Case ItemProcessStatus_Error:
            getMessage = f.Strings.Format(Msg.getText(FAILURE_MESSAGE), f.Collections.toString(pErrors, , VBA.vbCrLf))
        Case ItemProcessStatus_Correct:
            getMessage = f.Strings.Format(Msg.getText(SUCCESS_MESSAGE), _
                                    f.Collections.toString(pSuccessful, , VBA.vbCrLf))
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
    Dim project As EProject
    Dim files As VBA.Collection
    Dim varfilepath As Variant
    '[For Items Processor updates] ------------------------------------------------------------------------
    Dim caption As String
    Dim counter As Long
    '------------------------------------------------------------------------------------------------------
    Dim validation As ValidationObject
    '------------------------------------------------------------------------------------------------------

    On Error GoTo errHandler
    
    Call reset
    
    If pParent.isAborted Then Call VBA.Err.Raise(1, , Msg.getText(FAILED_BECAUSE_OF_PREDECESSORS))
    
    Set project = pParent.getProject

    Set files = f.files.getFolderFiles(project.getFolderPath, False, MACRO_FILE_REGEX_PATTERN)
    For Each varfilepath In files
        
        '[Updating ItemsProcessor progress bar] ---------------------------------------------|
        caption = f.Strings.Format(Msg.getText(PBAR_ADJUSTING_CODE), varfilepath)           '|
        Call pParent.getProcessor.triggerDetailStatusUpdate(caption, counter / files.Count) '|
        counter = counter + 1                                                               '|
        '------------------------------------------------------------------------------------|
        
        '------------------------------------------------------------------------------------|
        Set validation = replaceSingleFile(VBA.CStr(varfilepath))                           '|
        If validation.getStatus = ItemProcessStatus_Error Then                              '|
            Call pErrors.Add(validation.getMessage)                                         '|
        End If                                                                              '|
        '------------------------------------------------------------------------------------|
        
    Next varfilepath
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call pErrors.Add(f.Strings.Format(Msg.getText(ERROR_GENERAL), errDescription, errNumber))
    
End Sub


Private Function replaceSingleFile(filepath As String) As ValidationObject
    Const SUCCESS_MESSAGE_PATTERN As String = VB_COMPONENT_FULL_NAME_PATTERN
    Const FAILURE_MESSAGE_PATTERN As String = ERROR_CONVERTING_COMPONENT
    Const GENERAL_FAILURE_MESSAGE_PATTERN As String = ERROR_CONVERTING_PROJECT
    '----------------------------------------------------------------------------------------------------------
    Dim wkb As Excel.Workbook
    Dim project As vbide.VBProject
    Dim component As vbide.VBComponent
    Dim validation As ValidationObject
    Dim message As String
    '----------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    Set wkb = f.Books.open_(filepath, False, pParent.getNoMacrosExcel)
    Set project = f.Developer.getVbProject(wkb) ' ByName(filepath, pExcel)
    If project Is Nothing Then
        message = f.Strings.Format(Msg.getText(ERROR_OPENING_VBA_PROJECT), filepath)
        Call VBA.Err.Raise(1, , message)
    Else
        
        Set validation = renameProject(project)
        If validation.getStatus = ItemProcessStatus_Error Then
            Call VBA.Err.Raise(1, , validation.getMessage)
        End If
        
        For Each component In project.VBComponents
    
            '------------------------------------------------------------------------------------|
            Set validation = convertComponent(component)                                        '|
            If validation.getStatus = ItemProcessStatus_Skipped Then                            '|
            ElseIf validation.getStatus = ItemProcessStatus_Correct Then                        '|
                message = f.Strings.Format(SUCCESS_MESSAGE_PATTERN, _
                                                            project.fileName, component.name)   '|
                Call pSuccessful.Add(message)                                                   '|
            ElseIf validation.getStatus = ItemProcessStatus_Error Then                          '|
                message = f.Strings.Format(Msg.getText(FAILURE_MESSAGE_PATTERN), _
                                    project.fileName, component.name, validation.getMessage)    '|
                Call pErrors.Add(message)                                                       '|
            End If                                                                              '|
            '------------------------------------------------------------------------------------|
            
        Next component
    End If
        
    Call wkb.Close(True)
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    
    Set replaceSingleFile = f.create.ValidationObject(ItemProcessStatus_Correct)
    Exit Function

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    message = f.Strings.Format(Msg.getText(GENERAL_FAILURE_MESSAGE_PATTERN), _
                                                        project.fileName, errDescription, errNumber)
    Set replaceSingleFile = f.create.ValidationObject(ItemProcessStatus_Error, message)

End Function


Private Function renameProject(project As vbide.VBProject) As ValidationObject
    Dim message As String
    Dim newName As String
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    newName = VBA.LCase(f.Strings.replaceByDictionary(project.name, pParent.getProject.getReplacements))
    If newName <> project.name Then
        project.name = newName
    End If
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Set renameProject = f.create.ValidationObject(ItemProcessStatus_Correct)
    Exit Function

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    message = f.Strings.Format(Msg.getText(ERROR_RENAMING_PROJECT), project.name, newName, errDescription, errNumber)
    Set renameProject = f.create.ValidationObject(ItemProcessStatus_Error, message)

End Function

Private Function convertComponent(component As vbide.VBComponent) As ValidationObject
    Const ERROR_READING_MODULE_CONTENT As String = "CreatingNewProject.Errors.ReadingModuleContent"
    Dim errNumber As Long, errDescription As String
    '----------------------------------------------------------------------------------------------------------
    Dim codeModule As vbide.codeModule
    '----------------------------------------------------------------------------------------------------------
    Dim placeholders As Scripting.Dictionary
    Dim placeholder As String
    Dim content As String
    Dim arrContent As Variant
    Dim i As Long
    Dim lineContent As String
    Dim modified As String
    '[For user information] -----------------------------------------------------------------------------------
    Dim message As String
    Dim occurrences As Long
    Dim totalOccurrences As Long
    '----------------------------------------------------------------------------------------------------------

    On Error GoTo errHandler
    
    Set codeModule = component.codeModule
    

    '[Reading module content] -------------------------------------------------------|
    On Error Resume Next                                                            '|
    content = f.Developer.getCodeModuleContent(codeModule)                          '|
    arrContent = VBA.Split(content, VBA.vbCrLf)                                     '|
    If Not f.arrays.isDefined(arrContent) Then                                      '|
        errNumber = VBA.Err.Number: errDescription = VBA.Err.Description            '|
        message = f.Strings.Format(Msg.getText(ERROR_READING_MODULE_CONTENT), _
                                                                errDescription)     '|
        On Error GoTo errHandler                                                    '|
        Call VBA.Err.Raise(errNumber, , message)                                    '|
    End If                                                                          '|
    '--------------------------------------------------------------------------------|
    
    
    '[Modifying module if it contains any searched words] ---------------------------|
    On Error GoTo errHandler                                                        '|
    Set placeholders = pParent.getProject.getReplacements                           '|
    For i = LBound(arrContent) To UBound(arrContent)                                '|
        lineContent = arrContent(i)                                                 '|
        modified = f.Strings.replaceByDictionary(lineContent, placeholders)         '|
        If lineContent = modified Then                                              '|
            'Nothing changed                                                        '|
        Else                                                                        '|
            Call codeModule.ReplaceLine(i + 1 - LBound(arrContent), modified)       '|
            totalOccurrences = totalOccurrences + 1                                 '|
        End If                                                                      '|
    Next i                                                                          '|
    '--------------------------------------------------------------------------------|
    
'==========================================================================================================
ExitPoint:
    If totalOccurrences Then
        Set convertComponent = f.create.ValidationObject(ItemProcessStatus_Correct)
    Else
        Set convertComponent = f.create.ValidationObject(ItemProcessStatus_Skipped)
    End If
    
    Exit Function

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    message = f.Strings.Format(Exceptions.DESCRIPTION_PATTERN, errDescription, errNumber)
    Set convertComponent = f.create.ValidationObject(ItemProcessStatus_Error, message)
    
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


