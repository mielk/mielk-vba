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
'[Errors] -------------------------------------------------------------------------------------------------
Private Const ERR_INVALID_INPUT_DATA As String = "CreatingNewProject.Errors.RibbonInvalidInput"
Private Const ERR_CREATING_XML_FILE As String = "CreatingNewProject.Errors.CreatingXmlFile"
Private Const ERR_CREATING_JSON_OBJECT As String = "CreatingNewProject.Errors.CreatingJsonObject"
Private Const ERR_CREATING_JSON_FILE As String = "CreatingNewProject.Errors.CreatingJsonFile"
Private Const ERR_CREATING_MSG_FILE As String = "CreatingNewProject.Errors.CreatingMsgFile"
Private Const ERR_CREATING_VBA_OBJECTS As String = "CreatingNewProject.Errors.CreatingVbaObjects"
'[Status] -------------------------------------------------------------------------------------------------
Private Const SUCCESS_MESSAGE As String = "CreatingNewProject.Messages.GeneratingRibbonSuccess"
Private Const SEMISUCCESS_MESSAGE As String = "CreatingNewProject.Messages.GeneratingRibbonSemiSuccess"
Private Const FAILURE_MESSAGE As String = "CreatingNewProject.Messages.GeneratingRibbonFailure"
'[Progress bar] -------------------------------------------------------------------------------------------
Private Const ITEMS_PROCESSOR_CAPTION As String = "CreatingNewProject.Steps.CreatingRibbonComponents"
Private Const PBAR_CREATING_XML_FILE As String = "CreatingNewProject.Actions.CreatingXmlFile"
Private Const PBAR_CREATING_JSON_FILE As String = "CreatingNewProject.Actions.CreatingJsonFile"
Private Const PBAR_CREATING_MSG_FILE As String = "CreatingNewProject.Actions.CreatingMsgFile"
Private Const PBAR_INSERTING_XML_INTO_FILE As String = "CreatingNewProject.Actions.InsertingXmlIntoFile"
'----------------------------------------------------------------------------------------------------------
Private pParent As ProjectCreator
Private pRibbonTabControl As ERibbonControl
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
    Dim processor As ItemsProcessor:                Set processor = pParent.getProcessor
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    Call reset
    If pParent.isAborted Then Call VBA.Err.Raise(1, , Msg.getText(FAILED_BECAUSE_OF_PREDECESSORS))
    
    Set pRibbonTabControl = getTabControlFromSheet
        
    If Not pRibbonTabControl Is Nothing Then
    
        Call processor.triggerDetailStatusUpdate(Msg.getText(PBAR_CREATING_XML_FILE), 0)
        Call createXmlFile
        
        Call processor.triggerDetailStatusUpdate(Msg.getText(PBAR_CREATING_JSON_FILE), 0.33, 0.05)
        Call createJsonFile
        
        Call processor.triggerDetailStatusUpdate(Msg.getText(PBAR_CREATING_MSG_FILE), 0.67, 0.05)
        Call createMsgFile
    
        Call processor.triggerDetailStatusUpdate(Msg.getText(PBAR_INSERTING_XML_INTO_FILE), 0.8, 0.05)
        Call insertXmlIntoFile
    
    End If
    

Call pParent.getProcessor.triggerDetailStatusUpdate(vbNullString, 1)

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Sub

'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = VBA.Err.Number: errDescription = VBA.Err.Description
    Call pWarnings.Add(errDescription)

End Sub


Private Function getTabControlFromSheet() As ERibbonControl
    Dim wks As Excel.Worksheet
    Dim data As Variant
    Dim dtos As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set wks = pParent.getSheet(DataTypes.Ribbon)
    data = f.arrays.fromRange(wks.Cells(1, 1).CurrentRegion, False)
    Set dtos = f.arrays.toDtos(data)
    Set getTabControlFromSheet = getTabControlFromDtos(dtos)
    
End Function

Private Function getTabControlFromDtos(dtos As VBA.Collection) As ERibbonControl
    Dim dto As Scripting.Dictionary
    Dim ctrlTab As ERibbonControl
    Dim ctrl As ERibbonControl
    Dim prevCtrl As ERibbonControl
    '------------------------------------------------------------------------------------------------------
    Dim prevLevel As Long
    Dim level As Long
    '------------------------------------------------------------------------------------------------------
    Dim message As String
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    For Each dto In dtos
        Set ctrl = createRibbonControl(dto)
        level = ctrl.getLevel
        
        If level = 1 Then
            Set ctrlTab = ctrl
        ElseIf level < prevLevel Then
            Set prevCtrl = prevCtrl.getParent.getParent
        ElseIf level = prevLevel Then
            Set prevCtrl = prevCtrl.getParent
        End If
        
        If Not prevCtrl Is Nothing Then
            Call prevCtrl.addControl(ctrl)
            Call ctrl.setParent(prevCtrl)
        End If
        
        Set prevCtrl = ctrl
        prevLevel = prevCtrl.getLevel
        
    Next dto
    
    Set getTabControlFromDtos = ctrlTab

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Function
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    message = f.Strings.Format(Msg.getText(ERR_CREATING_VBA_OBJECTS), errDescription, errNumber)
    Call pErrors.Add(message)
    
End Function

Private Function createRibbonControl(dto As Scripting.Dictionary) As ERibbonControl
    Set createRibbonControl = New ERibbonControl
    With createRibbonControl
        Call .setProject(pParent.getProject)
        Call .loadFromDto(dto)
    End With
End Function






'[XML]
Private Sub createXmlFile()
    Dim filePath As String
    Dim template As String
    Dim content As String
    '------------------------------------------------------------------------------------------------------
    Dim stream As Scripting.TextStream
    '------------------------------------------------------------------------------------------------------
    
    filePath = f.files.concatenatePath( _
                                    pParent.getProject.getFolderPath, _
                                    Paths.FolderNameResources, _
                                    Paths.FolderNameRibbon, _
                                    RIBBON_XML_FILE_NAME)
    template = f.TextFiles.readTextFile(filePath)
    content = f.Strings.Format(template, pRibbonTabControl.getXml)
                                    
    '[Writing modified content to the original text file] -----------------------------------|
    On Error Resume Next                                                                    '|
    Call VBA.Err.Clear                                                                      '|
    Set stream = fso.OpenTextFile(filePath, ForWriting)                                     '|
    Call stream.WriteLine(content)                                                          '|
    Call stream.Close                                                                       '|
    '----------------------------------------------------------------------------------------|
    
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Call pSuccessful.Add(filePath)
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call pErrors.Add(f.Strings.Format(Msg.getText(ERR_CREATING_XML_FILE), filePath, errDescription, errNumber))
        
End Sub




'[JSON]
Private Sub createJsonFile()
    Dim filePath As String
    Dim json As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    filePath = f.files.concatenatePath(pParent.getProject.getFolderPath, _
                                            Paths.FolderNameResources, _
                                            Paths.FolderNameRibbon, _
                                            "ribbon.json")
                                            
    Set json = f.json.getJsonObjectFromFile(filePath)
    Call f.dictionaries.AddItem(json, Props_Project.controls.getName, getRibbonDefinitionsDictionary)
    Call f.TextFiles.printToTextFile(f.json.ConvertToJson(json, VBA.vbTab), filePath, True)

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Call pSuccessful.Add(filePath)
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call VBA.Err.Raise(errNumber, Description:=f.Strings.Format(Msg.getText(ERR_CREATING_JSON_FILE), _
                                                                                errNumber, errDescription))
       
End Sub


Private Function getRibbonDefinitionsDictionary() As Scripting.Dictionary
    Dim controls As VBA.Collection
    Dim control As ERibbonControl
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    Set getRibbonDefinitionsDictionary = f.dictionaries.create(False)
    
    Set controls = pRibbonTabControl.getDescendants(True)
    For Each control In controls
        With control
            Call f.dictionaries.AddItem( _
                        getRibbonDefinitionsDictionary, _
                        .getIdForXml, _
                        .getPropertiesDictionaryForJsonFile)
        End With
    Next control
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Exit Function
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call VBA.Err.Raise(errNumber, Description:=f.Strings.Format(Msg.getText(ERR_CREATING_JSON_OBJECT), _
                                                                                errNumber, errDescription))
       
End Function







'[MSG JSON FILES]
Private Sub createMsgFile()
    Dim filePath As String
    Dim json As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    filePath = f.files.concatenatePath(pParent.getProject.getFolderPath, _
                                            Paths.FolderNameResources, _
                                            Paths.FolderNameMsgs, _
                                            Paths.FolderNameProject, _
                                            "ribbon.json")
    Set json = getRibbonMessageDefinitions

    Call f.TextFiles.printToTextFile(f.json.ConvertToJson(json, VBA.vbTab), filePath, True)

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Call pSuccessful.Add(filePath)
    Exit Sub
    
'----------------------------------------------------------------------------------------------------------
errHandler:
    errNumber = Err.Number: errDescription = Err.Description
    Call VBA.Err.Raise(errNumber, Description:=f.Strings.Format(Msg.getText(ERR_CREATING_MSG_FILE), _
                                                                                errNumber, errDescription))
       
End Sub

Private Function getRibbonMessageDefinitions() As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Dim controls As VBA.Collection
    Dim control As ERibbonControl
    Dim controlsCaptions As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim tags As Variant
    Dim tag As String
    Dim i As Long
    Dim subdict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set dict = f.dictionaries.create(False)
    
    Set controls = pRibbonTabControl.getDescendants(True)
    For Each control In controls
        With control
            Set controlsCaptions = .getCaptions
            tags = VBA.Split(.getIdForXml, ".")
        End With
        
        If Not f.dictionaries.isNonEmptyDictionary(controlsCaptions) Then
            ' no caption found
        ElseIf Not f.arrays.isDefined(tags) Then
            ' undefined control id
        Else
            Set subdict = dict
            For i = LBound(tags) To UBound(tags)
                tag = f.Strings.toSentenceCase(VBA.CStr(tags(i)))
                If subdict.Exists(tag) Then
                    Set subdict = subdict.item(tag)
                ElseIf i = UBound(tags) Then
                    Call subdict.Add(tag, controlsCaptions)
                Else
                    Call subdict.Add(tag, f.dictionaries.create(False))
                    Set subdict = subdict.item(tag)
                End If
            Next i
            
        End If
        
    Next control
    
    
    Set getRibbonMessageDefinitions = f.dictionaries.createSingleEntryDict( _
                                            props.Messages.getName, _
                                            f.dictionaries.createSingleEntryDict(props.Ribbon.getName, dict))

End Function




Private Sub insertXmlIntoFile()
    Const FILE_PATH As String = "D:\vba-tests\apps\Testowa aplikacja\test-view.xlsm"
    Const XML_PATH As String = "D:\vba-tests\apps\Testowa aplikacja\resources\ribbon\ribbon.xml"
    '-----------------------------------------------------------------------------------------------
    Dim filePath As String
    Dim xmlPath As String
    Dim appender As RibbonXmlAppender
    '-----------------------------------------------------------------------------------------------
    
    With pParent.getProject
        filePath = .getViewFilePath
        xmlPath = f.files.concatenatePath(.getFolderPath, _
                                            Paths.FolderNameResources, _
                                            Paths.FolderNameRibbon, _
                                            "ribbon.xml")
    End With
    
    Set appender = New RibbonXmlAppender
    Call appender.append(filePath, xmlPath)
    
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
