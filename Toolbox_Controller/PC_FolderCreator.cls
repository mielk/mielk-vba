VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PC_FolderCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProcessItem
Implements IProjectCreatorProcessUnit

Private Const CLASS_NAME As String = "PC_FolderCreator"
'[Errors] -------------------------------------------------------------------------------------------------
Private Const ERR_NAME_ALREADY_USED As String = "CreatingNewProject.Errors.NameAlreadyUsed"
Private Const ERR_CREATING_PARENT_FOLDER As String = "CreatingNewProject.Errors.CreatingParentFolder"
Private Const ERR_COPYING_TEMPLATE_FILES As String = "CreatingNewProject.Errors.CopyingTemplateFiles"
Private Const ERR_DELETING_PREVIOUS_VERSION As String = "CreatingNewProject.Errors.DeletingPreviousVersion"
Private Const ERR_RENAMING_MACRO_FILES As String = "CreatingNewProject.Errors.RenamingMacroFiles"
Private Const ERR_RENAMING_MACRO_FILE As String = "CreatingNewProject.Errors.RenamingMacroFile"
'[Status] -------------------------------------------------------------------------------------------------
Private Const SUCCESS_MESSAGE As String = "CreatingNewProject.Messages.CreatingProjectFolderSuccess"
Private Const FAILURE_MESSAGE As String = "CreatingNewProject.Messages.CreatingProjectFolderFailure"
'[Progress bar] -------------------------------------------------------------------------------------------
Private Const ITEMS_PROCESSOR_CAPTION As String = "CreatingNewProject.Steps.CreatingProjectFolder"
Private Const PBAR_CREATING_DESTINATION_FOLDER As String = "CreatingNewProject.Actions.CreatingDestinationFolder"
Private Const PBAR_COPYING_TEMPLATE_FILES As String = "CreatingNewProject.Actions.CopyingTemplateFiles"
Private Const PBAR_CHECKING_PREVIOUS_VERSION As String = "CreatingNewProject.Actions.CheckingPreviousVersion"
Private Const PBAR_DELETING_PREVIOUS_VERSION As String = "CreatingNewProject.Actions.DeletingPreviousVersion"
Private Const PBAR_RENAMING_MACRO_FILES As String = "CreatingNewProject.Actions.RenamingMacroFiles"
'----------------------------------------------------------------------------------------------------------
Private pParent As ProjectCreator
'----------------------------------------------------------------------------------------------------------
Private pWarnings As VBA.Collection
Private pErrors As VBA.Collection
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
    Select Case getStatus
        Case ItemProcessStatus_Correct
            With pParent.getProject
                getMessage = f.Strings.Format(Msg.getText(SUCCESS_MESSAGE), .getName, .getFolderPath)
            End With
        Case ItemProcessStatus_Error
            getMessage = f.Collections.toString(pErrors, , VBA.vbCrLf)
    End Select
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
End Sub


'[API]
Public Sub process()
    Dim message As String
    Dim processor As ItemsProcessor:                Set processor = pParent.getProcessor
    Dim destinationFolder As Scripting.Folder
    '------------------------------------------------------------------------------------------------------

    On Error GoTo errHandler
    
    Call reset
    
    Call checkProjectNameUniqueness
    
Call processor.triggerDetailStatusUpdate(Msg.getText(PBAR_CREATING_DESTINATION_FOLDER), 0.5, 0.05)
    Set destinationFolder = createDestinationFolder
    
message = f.Strings.Format(Msg.getText(PBAR_COPYING_TEMPLATE_FILES), pParent.getProject.getFolderPath)
Call processor.triggerDetailStatusUpdate(message, 0.65, 0.03)
    If Not destinationFolder Is Nothing Then
        Call copyTemplateFiles

Call processor.triggerDetailStatusUpdate(Msg.getText(PBAR_COPYING_TEMPLATE_FILES), 0.85, 0.05)
        Call renameMacroFiles
        Call removeRedundantFiles
        
    End If

Call pParent.getProcessor.triggerDetailStatusUpdate(vbNullString, 1)

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If f.Collections.isNonEmptyCollection(pErrors) Then Call pParent.abort
    Exit Sub

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    Call pErrors.add(errDescription)
    GoTo ExitPoint

End Sub



Private Sub checkProjectNameUniqueness()
    Dim project As EProject:                Set project = pParent.getProject
    Dim message As String
    '------------------------------------------------------------------------------------------------------

    On Error GoTo errHandler

    Call pParent.getProcessor.triggerDetailStatusUpdate(Msg.getText(PBAR_CHECKING_PREVIOUS_VERSION), 0.1, 0.03)

    If project.isUniqueName Then
        ' There is no such folder yet.
    ElseIf project.overrideIfExists Then
        Call pParent.getProcessor.triggerDetailStatusUpdate(Msg.getText(PBAR_DELETING_PREVIOUS_VERSION), 0.3, 0.03)
        Call f.files.DeleteFolder(project.getFolderPath)                                    '|
    Else
        On Error GoTo 0
        message = f.Strings.Format(Msg.getText(ERR_NAME_ALREADY_USED), project.getName)
        Call VBA.Err.Raise(1, , message)
    End If
    

'==========================================================================================================
ExitPoint:
    Exit Sub

'----------------------------------------------------------------------------------------------------------
errHandler:
    message = f.Strings.Format(Msg.getText(ERR_DELETING_PREVIOUS_VERSION), _
                                    project.getName, VBA.Err.Description, VBA.Err.Number)
    Call VBA.Err.Raise(1, , message)

End Sub


Private Function createDestinationFolder() As Scripting.Folder
    Dim project As EProject:                Set project = pParent.getProject
    Dim destinationFolderPath As String
    Dim message As String
    '------------------------------------------------------------------------------------------------------

    destinationFolderPath = project.getParentFolderPath
    Set createDestinationFolder = f.files.CreateFolder(destinationFolderPath)
                                                                                            
    If Not f.files.FolderExists(destinationFolderPath) Then
        message = f.Strings.Format(Msg.getText(ERR_CREATING_PARENT_FOLDER), _
                                                                    destinationFolderPath)
        On Error GoTo 0
        Call VBA.Err.Raise(1, , message)
    End If

'==========================================================================================================
ExitPoint:
    'No error handling on purpose. Errors should be caught directly by the invoking function.
    
End Function

Private Sub copyTemplateFiles()
    Dim project As EProject:                Set project = pParent.getProject
    Dim message As String
    '------------------------------------------------------------------------------------------------------
    
    Call VBA.Err.Clear
    
    On Error Resume Next
    Call f.files.CopyFolder(Paths.ProjectTemplateFolderPath, project.getFolderPath)
                                                                                            
    If VBA.Err.Number Then
        message = f.Strings.Format(Msg.getText(ERR_CREATING_PARENT_FOLDER), _
                                        project.getFolderPath, VBA.Err.Description, VBA.Err.Number)
        On Error GoTo 0
        Call VBA.Err.Raise(1, , message)
    End If

'==========================================================================================================
ExitPoint:
    'No error handling on purpose. Errors should be caught directly by the invoking function.
    
End Sub


Private Sub renameMacroFiles()
    Dim errNumber As Long, errDescription As String
    '------------------------------------------------------------------------------------------------------
    Dim replacements As Scripting.Dictionary
    Dim message As String
    Dim files As VBA.Collection
    Dim file As Scripting.file
    Dim newPath As String
    Dim oldPath As String
    '------------------------------------------------------------------------------------------------------
    
    With pParent.getProject
        Set files = f.files.getFolderFiles(.getFolderPath, False, getAllPlaceholdersRegex)
        Set replacements = .getReplacements
    End With
    
    For Each file In files
    
        On Error Resume Next
        oldPath = file.path
        newPath = f.Strings.replaceByDictionary(file.path, replacements)
        
        If newPath = oldPath Then
            ' Nothing changed
        Else
            Name oldPath As newPath
            errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
            If errNumber Then
                message = f.Strings.Format(Msg.getText(ERR_RENAMING_MACRO_FILE), _
                                                                oldPath, newPath, errDescription, errNumber)
                On Error GoTo 0
                Call VBA.Err.Raise(errNumber, , message)
                Call pErrors.add(message)
            End If
        End If
    Next file

'==========================================================================================================
ExitPoint:
    Exit Sub
    'No error handling on purpose. Errors should be caught directly by the invoking function.
    
End Sub



Private Sub removeRedundantFiles()
    Const FILE_NAME_PATTERN As String = "{0}-view.xl{1}m"
    '------------------------------------------------------------------------------------------------------
    Dim project As EProject
    Dim fileName As String
    Dim path As String
    '------------------------------------------------------------------------------------------------------
    
    Set project = pParent.getProject
    fileName = f.Strings.Format(FILE_NAME_PATTERN, _
                                    project.getCodeName, _
                                    VBA.IIf(project.isAddIn, "s", "a"))
    path = f.files.concatenatePath(project.getFolderPath, fileName)
    
    Call f.files.DeleteFile(path)
    
End Sub



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
