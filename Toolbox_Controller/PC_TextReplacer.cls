VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PC_TextReplacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProcessItem
Implements IProjectCreatorProcessUnit

Private Const CLASS_NAME As String = "PC_CodeReplacer"
'[Errors] -------------------------------------------------------------------------------------------------
Private Const ERROR_GENERAL As String = "CreatingNewProject.Errors.ApplyingChangesToTextFiles"
Private Const ERROR_READING_FROM_TEXT_FILE As String = "CreatingNewProject.Errors.ReadingFromTextFile"
Private Const ERROR_WRITING_TO_TEXT_FILE As String = "CreatingNewProject.Errors.WritingToTextFile"
'[Status] -------------------------------------------------------------------------------------------------
Private Const SUCCESS_MESSAGE As String = "CreatingNewProject.Messages.ApplyingChangesToTextFilesSuccess"
Private Const SEMISUCCESS_MESSAGE As String = "CreatingNewProject.Messages.ApplyingChangesToTextFilesSemiSuccess"
Private Const FAILURE_MESSAGE As String = "CreatingNewProject.Messages.ApplyingChangesToTextFilesFailure"
'[Progress bar] -------------------------------------------------------------------------------------------
Private Const ITEMS_PROCESSOR_CAPTION As String = "CreatingNewProject.Steps.ApplyingChangesToTextFiles"
Private Const PBAR_ADJUSTING_TEXT_FILE As String = "CreatingNewProject.Actions.ApplyingChangesToTextFile"
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
                If f.Collections.isNonEmptyCollection(pSuccessful) Then
                    getMessage = f.Strings.Format(Msg.getText(SEMISUCCESS_MESSAGE), f.Collections.toString(pSuccessful, , VBA.vbCrLf))
                Else
                    getMessage = f.Collections.toString(pErrors, , VBA.vbCrLf)
                End If
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
    Dim files As VBA.Collection
    Dim varPath As Variant
    Dim replacements As Scripting.Dictionary
    '[For Items Processor purposes] -----------------------------------------------------------------------
    Dim message As String
    Dim counter As Long
    '[For message purposes] -------------------------------------------------------------------------------
    Dim validation As ValidationObject
    '------------------------------------------------------------------------------------------------------

    On Error GoTo errHandler
    
    Call reset
    
    If pParent.isAborted Then Call VBA.Err.Raise(1, , Msg.getText(FAILED_BECAUSE_OF_PREDECESSORS))
    
    
    With pParent.getProject
        Set replacements = .getReplacements
        
        'If project should be created as an add-in, -----------------------------------------|
        ' all "*-view.xlsm" should be replaced with *-view.xlam.                            '|
        If .isAddIn Then                                                                    '|
            Call f.dictionaries.AddItem(replacements, _
                                EXTENSION_EXCEL_MACRO_FILE, EXTENSION_EXCEL_ADDIN)          '|
        End If                                                                              '|
        '------------------------------------------------------------------------------------|
        
    End With
    
    Set files = getTextFilesToBeAdjusted
    For Each varPath In files
    
        '[Updating progress bar on ItemsProcessor] ------------------------------------------|
        message = f.Strings.Format(Msg.getText(PBAR_ADJUSTING_TEXT_FILE), varPath)          '|
        Call pParent.getProcessor.triggerDetailStatusUpdate(message, counter / files.Count) '|
        counter = counter + 1                                                               '|
        '------------------------------------------------------------------------------------|
                                                                                                    
        '------------------------------------------------------------------------------------|
        Set validation = adjustTextFile(VBA.CStr(varPath), replacements)                    '|
        Select Case validation.getStatus                                                    '|
            Case ItemProcessStatus_Error:       Call pErrors.Add(validation.getMessage)     '|
            Case ItemProcessStatus_Correct:     Call pSuccessful.Add(varPath)               '|
        End Select                                                                          '|
        '------------------------------------------------------------------------------------|

    Next varPath

Call pParent.getProcessor.triggerDetailStatusUpdate(vbNullString, 1)

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Sub

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    Call pErrors.Add(f.Strings.Format(Msg.getText(ERROR_GENERAL), errDescription, errNumber))

End Sub



Private Function getTextFilesToBeAdjusted() As VBA.Collection
    Dim project As EProject
    Dim file As Scripting.file
    '----------------------------------------------------------------------------------------------------------
    
    Set getTextFilesToBeAdjusted = New VBA.Collection
    Set project = pParent.getProject

    With getTextFilesToBeAdjusted
    
        '[VBS files]
        For Each file In f.files.getFolderFiles(project.getFolderPath, False, VBS_FILE_PATTERN)
            Call .Add(file.path)
        Next file
            
        '[Msgs - ribbon.json]
        Call .Add(f.files.concatenatePath(project.getFolderPath, Paths.FolderNameResources, _
                                        Paths.FolderNameMsgs, "project", "ribbon.json"))

        '[Ribbon - ribbon.json]
        Call .Add(f.files.concatenatePath(project.getFolderPath, Paths.FolderNameResources, _
                                        Paths.FolderNameRibbon, "ribbon.json"))

        '[Ribbon - ribbon.xml]
        Call .Add(f.files.concatenatePath(project.getFolderPath, Paths.FolderNameResources, _
                                        Paths.FolderNameRibbon, "ribbon.xml"))

    End With

End Function

Private Function adjustTextFile(filepath As String, replacements As Scripting.Dictionary) As ValidationObject
    Dim errNumber As Long, errDescription As String
    '------------------------------------------------------------------------------------------------------
    Dim stream As Scripting.TextStream
    Dim content As String
    '[Validation] -----------------------------------------------------------------------------------------
    Dim message As String
    '------------------------------------------------------------------------------------------------------


    '[Reading original content of the file] -------------------------------------------------|
    On Error Resume Next                                                                    '|
    Call VBA.Err.Clear                                                                      '|
    Set stream = fso.OpenTextFile(filepath, ForReading)                                     '|
    content = stream.ReadAll                                                                '|
    Call stream.Close                                                                       '|
                                                                                            '|
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description                        '|
    If errNumber Then                                                                       '|
        On Error GoTo errHandler                                                            '|
        message = f.Strings.Format(Msg.getText(ERROR_READING_FROM_TEXT_FILE), filepath, _
                                                                errDescription, errNumber)  '|
        Call VBA.Err.Raise(errNumber, , errDescription)                                     '|
    End If                                                                                  '|
    '----------------------------------------------------------------------------------------|
    
    '[Writing modified content to the original text file] -----------------------------------|
    On Error Resume Next                                                                    '|
    Call VBA.Err.Clear                                                                      '|
    content = f.Strings.replaceByDictionary(content, replacements)                          '|
    Set stream = fso.OpenTextFile(filepath, ForWriting)                                     '|
    Call stream.WriteLine(content)                                                          '|
    Call stream.Close                                                                       '|
                                                                                            '|
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description                        '|
    If errNumber Then                                                                       '|
        On Error GoTo errHandler                                                            '|
        message = f.Strings.Format(Msg.getText(ERROR_WRITING_TO_TEXT_FILE), filepath, _
                                                                errDescription, errNumber)  '|
        Call VBA.Err.Raise(errNumber, , message)                                            '|
    End If                                                                                  '|
    '----------------------------------------------------------------------------------------|


'==========================================================================================================
ExitPoint:
    Set adjustTextFile = f.create.ValidationObject(ItemProcessStatus_Correct)
    Exit Function

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    Set adjustTextFile = f.create.ValidationObject(ItemProcessStatus_Error, errDescription)

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

