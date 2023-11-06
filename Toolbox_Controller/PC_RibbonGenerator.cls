VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PC_RibbonGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProcessItem
Implements IProjectCreatorProcessUnit

Private Const CLASS_NAME As String = "PC_RibbonGenerator"
''[Errors] -------------------------------------------------------------------------------------------------
Private Const ERR_INVALID_INPUT_DATA As String = "CreatingNewProject.Errors.RibbonInvalidInput"
Private Const ERR_CREATING_XML_FILE As String = "CreatingNewProject.Errors.CreatingXmlFile"
Private Const ERR_CREATING_JSON_FILE As String = "CreatingNewProject.Errors.CreatingJsonFile"
Private Const ERR_CREATING_MSG_FILE As String = "CreatingNewProject.Errors.CreatingMsgFile"
'[Status] -------------------------------------------------------------------------------------------------
Private Const SUCCESS_MESSAGE As String = "CreatingNewProject.Messages.GeneratingRibbonSuccess"
Private Const SEMISUCCESS_MESSAGE As String = "CreatingNewProject.Messages.GeneratingRibbonSemiSuccess"
Private Const FAILURE_MESSAGE As String = "CreatingNewProject.Messages.GeneratingRibbonFailure"
''[Progress bar] -------------------------------------------------------------------------------------------
Private Const ITEMS_PROCESSOR_CAPTION As String = "CreatingNewProject.Steps.CreatingRibbonComponents"
Private Const PBAR_CREATING_XML_FILE As String = "CreatingNewProject.Actions.CreatingXmlFile"
Private Const PBAR_CREATING_JSON_FILE As String = "CreatingNewProject.Actions.CreatingJsonFile"
Private Const PBAR_CREATING_MSG_FILE As String = "CreatingNewProject.Actions.CreatingMsgFile"
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
    Stop
    If pParent.isAborted Then
        getMessage = Msg.getText(FAILED_BECAUSE_OF_PREDECESSORS)
    Else
        getMessage = Msg.getText(NOT_IMPLEMENTED_YET)
    '    Select Case getStatus
    '        Case ItemProcessStatus_Correct
    '            With pParent.getProject
    '                getMessage = f.Strings.Format(Msg.getText(SUCCESS_MESSAGE), .getName, .getFolderPath)
    '            End With
    '        Case ItemProcessStatus_Error
    '            getMessage = f.Collections.toString(pErrors, , VBA.vbCrLf)
    '    End Select
    End If
End Function

Public Function getStatus() As Long
    If pErrors Is Nothing Then
        getStatus = ItemProcessStatus_Unknown
    ElseIf pErrors.Count Then
        Stop
        getStatus = ItemProcessStatus_Error
    Else
        Stop
        getStatus = ItemProcessStatus_Correct
    End If
End Function

Private Sub reset()
    Set pWarnings = New VBA.Collection
    Set pErrors = New VBA.Collection
End Sub


'[API]
Public Sub process()
    Dim processor As ItemsProcessor:                Set processor = pParent.getProcessor
    '------------------------------------------------------------------------------------------------------
    Dim ribbonControls As VBA.Collection
    '------------------------------------------------------------------------------------------------------

    On Error GoTo errHandler
    
    Call reset
    If pParent.isAborted Then Call VBA.Err.Raise(1, , Msg.getText(FAILED_BECAUSE_OF_PREDECESSORS))
    
    Set ribbonControls = loadRibbonControlsFromSheet
    
    If f.Collections.isNonEmptyCollection(ribbonControls) Then
    
        Call processor.triggerDetailStatusUpdate(Msg.getText(PBAR_CREATING_XML_FILE), 0)
        Call createXmlFile
        
        Call processor.triggerDetailStatusUpdate(Msg.getText(PBAR_CREATING_JSON_FILE), 0.33, 0.05)
        Call createJsonFile
        
        Call processor.triggerDetailStatusUpdate(Msg.getText(PBAR_CREATING_MSG_FILE), 0.67, 0.05)
        Call createMsgFile
    
    End If
    

Call pParent.getProcessor.triggerDetailStatusUpdate(vbNullString, 1)

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Sub

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    Call pWarnings.add(errDescription)

End Sub


Private Function loadRibbonControlsFromSheet() As VBA.Collection
    Dim wks As Excel.Worksheet
    Dim data As Variant
    Dim dtos As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set wks = pParent.getSheet(DataTypes.Ribbon)
    data = f.arrays.fromRange(wks.Cells(1, 1).CurrentRegion, False)
    Set dtos = f.arrays.toDtos(data)
    Set loadRibbonControlsFromSheet = getRibbonControlsCollectionFromDtos(dtos)
    
End Function

Private Function getRibbonControlsCollectionFromDtos(dtos As VBA.Collection) As VBA.Collection
    Dim dto As Scripting.Dictionary
    Dim control As ERibbonControl
    '------------------------------------------------------------------------------------------------------
    Set getRibbonControlsCollectionFromDtos = New VBA.Collection
    For Each dto In dtos
        Set control = New ERibbonControl
        Call control.loadFromDto(dto)
        Call getRibbonControlsCollectionFromDtos.add(control)
    Next dto
End Function




Private Sub createXmlFile()
    Dim destinationPath As String
    '------------------------------------------------------------------------------------------------------
    
    Stop
    
    destinationPath = ""

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call pErrors.add(f.Strings.Format(Msg.getText(ERR_CREATING_XML_FILE), destinationPath, errDescription, errNumber))
        
End Sub

Private Sub createJsonFile()
    Dim destinationPath As String
    '------------------------------------------------------------------------------------------------------
    
    Stop
    
    destinationPath = ""

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call pErrors.add(f.Strings.Format(Msg.getText(ERR_CREATING_JSON_FILE), destinationPath, errDescription, errNumber))
       
End Sub

Private Sub createMsgFile()
    Dim destinationPath As String
    '------------------------------------------------------------------------------------------------------
    
    Stop
    
    destinationPath = ""

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call pErrors.add(f.Strings.Format(Msg.getText(ERR_CREATING_MSG_FILE), destinationPath, errDescription, errNumber))
       
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
