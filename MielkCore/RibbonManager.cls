VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "RibbonManager"
'----------------------------------------------------------------------------------------------------------
Private Const RIBBON_POINTER_EXCEL_NAME As String = "RibbonPointer"
Private Const INVALIDATION_CONTROLS_TAG As String = "invalidationControls"
Private Const GENERAL_INVALIDATION_GROUP As String = "general"
Private Const CONTROLS_TAG As String = "controls"
Private Const TRANSLATABLE_TAG As String = "translatable"
Private Const ICONS_FOLDER_PATH_TAG As String = "iconsFolder"
'----------------------------------------------------------------------------------------------------------
Private Const TAG_VALUE As String = "value"
Private Const TAG_FUNCTIONS As String = "functions"
'----------------------------------------------------------------------------------------------------------
Private pRibbon As IRibbonUI
Private pWorkbook As Excel.Workbook
Private pSettings As Scripting.Dictionary
Private pControlsProps As Scripting.Dictionary
Private pInvalidationGroups As Scripting.Dictionary
Private pIconsFolderPath As String
Private pUpdateDisabled As Boolean
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    Call initializeContainers
End Sub

Private Sub initializeContainers()
    Set pSettings = f.dictionaries.Create(False)
    Set pControlsProps = f.dictionaries.Create(False)
    Set pInvalidationGroups = f.dictionaries.Create(False)
End Sub




'[SETTERS]
Public Function setRibbon(Ribbon As IRibbonUI) As RibbonManager
    Set setRibbon = Me
    Set pRibbon = Ribbon
    Call saveRibbonPointer
End Function

Public Function setWorkbook(value As Excel.Workbook) As RibbonManager
    Set setWorkbook = Me
    Set pWorkbook = value
    Call saveRibbonPointer
End Function

Public Function setJsonFilePath(value As String) As RibbonManager
    Set setJsonFilePath = Me
    Call loadSettings(value)
End Function

Public Function setUpdateDisabled(value As Boolean) As RibbonManager
    Set setUpdateDisabled = Me
    pUpdateDisabled = value
End Function



'[RIBBON OBJECT]
Private Sub saveRibbonPointer()
    If Not f.Books.isValid(pWorkbook) Then
    ElseIf pRibbon Is Nothing Then
    Else
        On Error Resume Next
        Call f.ExcelNames.addValue(pWorkbook, RIBBON_POINTER_EXCEL_NAME, VBA.ObjPtr(pRibbon))
    End If
End Sub

#If VBA7 Then
Public Function readRibbonPointer() As LongPtr
#Else
Public Function readRibbonPointer() As Long
#End If
    Dim varValue As Variant
    '------------------------------------------------------------------------------------------------------
    varValue = f.ExcelNames.getValue(pWorkbook, RIBBON_POINTER_EXCEL_NAME)
    On Error Resume Next
    If VBA.isNumeric(varValue) Then
        #If VBA7 Then
            readRibbonPointer = VBA.CLngPtr(varValue)
        #Else
            readRibbonPointer = VBA.CLng(varValue)
        #End If
    End If
End Function


Private Function getRibbon() As IRibbonUI
    #If VBA7 Then
        Dim ribbonPointer As LongPtr
    #Else
        Dim ribbonPointer As Long
    #End If
    '------------------------------------------------------------------------------------------------------
    If pRibbon Is Nothing Then
        ribbonPointer = readRibbonPointer
        If ribbonPointer Then
            On Error Resume Next
            Call CopyMemory(getRibbon, ribbonPointer, 4)
            Set pRibbon = getRibbon
            On Error GoTo 0
        End If
    End If
    Set getRibbon = pRibbon
End Function




'[SETTINGS]
Private Sub loadSettings(filePath As String)
    Dim json As String
    '------------------------------------------------------------------------------------------------------

    On Error Resume Next
    json = f.TextFiles.readTextFile(filePath)
    If VBA.Len(json) Then Set pSettings = f.json.ParseJson(json)
    On Error GoTo 0

    'Icons folder path.
    pIconsFolderPath = getImageFolderPath(filePath)

    'Ribbon controls properties
    If f.dictionaries.isNonEmptyDictionary(pSettings) Then
        With pSettings
            If .exists(CONTROLS_TAG) Then Set pControlsProps = .Item(CONTROLS_TAG)
        End With
    End If
    
    'Invalidation groups
    Call loadInvalidationGroups

End Sub


Private Function getImageFolderPath(configFilePath As String)
    Const INDIRECT_PATH_SYMBOL As String = "@\"
    Const PARENT_FOLDER_SYMBOL As String = "."
    '------------------------------------------------------------------------------------------------------
    Dim rawPath As String
    Dim parts As Variant
    Dim i As Long
    Dim folder As Scripting.folder
    '------------------------------------------------------------------------------------------------------
    
    rawPath = f.dictionaries.getItem(pSettings, ICONS_FOLDER_PATH_TAG)
    If f.strings.startsWith(rawPath, INDIRECT_PATH_SYMBOL) Then
        parts = VBA.Split(VBA.replace(rawPath, INDIRECT_PATH_SYMBOL, vbNullString), "\")
        Set folder = f.files.createFolder(f.files.getFileFolderPath(configFilePath))
        For i = LBound(parts) To UBound(parts)
            If parts(i) = PARENT_FOLDER_SYMBOL Then
                Set folder = folder.parentFolder
            Else
                Set folder = f.files.getSubfolder(folder, VBA.CStr(parts(i)))
                If folder Is Nothing Then Exit For
            End If
        Next i
    Else
        getImageFolderPath = rawPath
    End If
    
    If Not folder Is Nothing Then
        getImageFolderPath = folder.path
    End If
    
End Function


Private Sub loadInvalidationGroups()
    Dim generalControls As VBA.Collection
    Dim generalDict As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set pInvalidationGroups = f.Objects.coalesce( _
                                    f.dictionaries.getObject(pSettings, INVALIDATION_CONTROLS_TAG), _
                                    f.dictionaries.Create(False))
    
    'Update [general] group.
    Set generalControls = f.dictionaries.getObject(pInvalidationGroups, GENERAL_INVALIDATION_GROUP)
    Set generalDict = f.Objects.coalesce( _
                            f.dictionaries.createFrom(generalControls), f.dictionaries.Create(False))
    For Each key In pControlsProps.Keys
        Call f.dictionaries.addItem(generalDict, key, key)
    Next key
    Set generalControls = f.dictionaries.toCollection(generalDict, DictPart_KeyOnly)
    
    Call f.dictionaries.addItem(pInvalidationGroups, GENERAL_INVALIDATION_GROUP, generalControls, duplicateBehaviour_Override)
    
End Sub


Public Sub assignProperty(controlId As String, propertyName As String, ByRef value As Variant)
    Dim ctrlProps As Scripting.Dictionary
    Dim propDict As Scripting.Dictionary
    Dim functionGenerated As Boolean
    '------------------------------------------------------------------------------------------------------

    'If pControlsProps Is Nothing Then Call loadSettings(pSettingsFilePath)

    On Error Resume Next
    Set ctrlProps = pControlsProps.Item(controlId)
    On Error GoTo 0

    If Not ctrlProps Is Nothing Then
        With ctrlProps
            If .exists(propertyName) Then
                Set propDict = .Item(propertyName)
                value = getValueFromPropDict(controlId, propDict, functionGenerated)
            End If
        End With

        If isPropertyTranslatable(propertyName) Then
            If Not functionGenerated Then
                If VBA.Len(value) Then
                    value = MsgService.getText(VBA.CStr(value))
                End If
            End If
        End If
    End If

End Sub


Private Function getValueFromPropDict(controlId As String, propDict As Scripting.Dictionary, Optional ByRef functionGenerated As Boolean) As Variant
    Dim value As Variant
    Dim Functions As VBA.Collection
    Dim functionsSet As VBA.Collection
    Dim varFunction As Variant
    Dim result As Variant
    '------------------------------------------------------------------------------------------------------

    Call f.dictionaries.assignValueFromDictToVariable(propDict, value, TAG_VALUE)
    Call f.dictionaries.assignValueFromDictToVariable(propDict, Functions, TAG_FUNCTIONS)

    If f.Collections.isNonEmptyCollection(Functions) Then
        functionGenerated = True
        For Each functionsSet In Functions
            For Each varFunction In functionsSet
            
                'Get value from function ------------------------------------------------------------|
                On Error Resume Next                                                                '|
                Call VBA.err.clear                                                                  '|
                result = Application.Run(f.Reflection.getFunctionFullName( _
                                                    VBA.CStr(varFunction), pWorkbook))              '|
                If VBA.err.number = Exceptions.ArgumentNotOptional.getNumber Then                   '|
                    result = Application.Run(f.Reflection.getFunctionFullName( _
                                                    VBA.CStr(varFunction), pWorkbook), controlId)   '|
                End If                                                                              '|
                If Not DEV_MODE Then On Error GoTo errHandler Else On Error GoTo 0                  '|
                '------------------------------------------------------------------------------------
                
                If result = False Then Exit For
            Next varFunction
            
            If Not f.Variables.isUndefined(result, False, False) Then Exit For
            
        Next functionsSet
        getValueFromPropDict = result
    Else
        getValueFromPropDict = value
    End If
    
    
errHandler:
    
End Function

Private Function isPropertyTranslatable(propName As String) As Boolean
    Static translatableProps As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    If translatableProps Is Nothing Then
        Set translatableProps = f.dictionaries.createFrom(Array("label", "tooltip"))
    End If
    isPropertyTranslatable = translatableProps.exists(propName)
End Function





'[Icons]
Public Sub assignControlImage(controlId As String, ByRef image)
    Dim imageFileName As String
    Dim imagePath As String
    '------------------------------------------------------------------------------------------------------
    
    Call assignProperty(controlId, "image", imageFileName)
    If VBA.Len(imageFileName) = 0 Then
        ' Image not defined
    ElseIf Not f.regex.checkIfMatch(imageFileName, IMAGE_FILE_REGEX_PATTERN) Then
        ' msoImage
        image = imageFileName
    Else
        ' File image
        imagePath = f.files.concatenatePath(pIconsFolderPath, imageFileName)
        Set image = LoadPictureGDI(imagePath)
    End If
End Sub







'[UPDATING]
Public Sub update(Optional eventName As String = GENERAL_INVALIDATION_GROUP)
    Dim Ribbon As Variant
    Dim invalidationControls As VBA.Collection
    Dim varControlId As Variant
    '------------------------------------------------------------------------------------------------------

    If Not pUpdateDisabled Then
        Set Ribbon = getRibbon
        If Not Ribbon Is Nothing Then
            Set invalidationControls = f.dictionaries.getObject(pInvalidationGroups, eventName)
            If invalidationControls Is Nothing Then
                Set invalidationControls = f.dictionaries.getObject(pInvalidationGroups, GENERAL_INVALIDATION_GROUP)
            End If
                
            If f.Collections.isNonEmptyCollection(invalidationControls) Then
                With Ribbon
                    For Each varControlId In invalidationControls
                        Call .InvalidateControl(varControlId)
                    Next varControlId
                End With
            End If
        End If
    End If

End Sub
