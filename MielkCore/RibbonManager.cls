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
    Set pSettings = F.Dictionaries.Create(False)
    Set pControlsProps = F.Dictionaries.Create(False)
    Set pInvalidationGroups = F.Dictionaries.Create(False)
End Sub




'[SETTERS]
Public Function setRibbon(ribbon As IRibbonUI) As RibbonManager
    Set setRibbon = Me
    Set pRibbon = ribbon
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
    If Not F.Books.isValid(pWorkbook) Then
    ElseIf pRibbon Is Nothing Then
    Else
        On Error Resume Next
        Call F.ExcelNames.addValue(pWorkbook, RIBBON_POINTER_EXCEL_NAME, VBA.ObjPtr(pRibbon))
    End If
End Sub

#If VBA7 Then
Public Function readRibbonPointer() As LongPtr
#Else
Public Function readRibbonPointer() As Long
#End If
    Dim varValue As Variant
    '------------------------------------------------------------------------------------------------------
    varValue = F.ExcelNames.getValue(pWorkbook, RIBBON_POINTER_EXCEL_NAME)
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
    json = F.TextFiles.readTextFile(filePath)
    If VBA.Len(json) Then Set pSettings = F.json.ParseJson(json)
    On Error GoTo 0

    'Icons folder path.
    pIconsFolderPath = F.Dictionaries.getItem(pSettings, ICONS_FOLDER_PATH_TAG)

    'Ribbon controls properties
    If F.Dictionaries.isNonEmptyDictionary(pSettings) Then
        With pSettings
            If .exists(CONTROLS_TAG) Then Set pControlsProps = .item(CONTROLS_TAG)
        End With
    End If
    
    'Invalidation groups
    Call loadInvalidationGroups

End Sub


Private Sub loadInvalidationGroups()
    Dim generalControls As VBA.Collection
    Dim generalDict As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set pInvalidationGroups = F.Objects.coalesce( _
                                    F.Dictionaries.getObject(pSettings, INVALIDATION_CONTROLS_TAG), _
                                    F.Dictionaries.Create(False))
    
    'Update [general] group.
    Set generalControls = F.Dictionaries.getObject(pInvalidationGroups, GENERAL_INVALIDATION_GROUP)
    Set generalDict = F.Objects.coalesce( _
                            F.Dictionaries.createFrom(generalControls), F.Dictionaries.Create(False))
    For Each key In pControlsProps.Keys
        Call F.Dictionaries.addItem(generalDict, key, key)
    Next key
    Set generalControls = F.Dictionaries.toCollection(generalDict, DictPart_KeyOnly)
    
    Call F.Dictionaries.addItem(pInvalidationGroups, GENERAL_INVALIDATION_GROUP, generalControls, duplicateBehaviour_Override)
    
End Sub


Public Sub assignProperty(controlId As String, propertyName As String, ByRef value As Variant)
    Dim ctrlProps As Scripting.Dictionary
    Dim propDict As Scripting.Dictionary
    Dim functionGenerated As Boolean
    '------------------------------------------------------------------------------------------------------

    'If pControlsProps Is Nothing Then Call loadSettings(pSettingsFilePath)

    On Error Resume Next
    Set ctrlProps = pControlsProps.item(controlId)
    On Error GoTo 0

    If Not ctrlProps Is Nothing Then
        With ctrlProps
            If .exists(propertyName) Then
                Set propDict = .item(propertyName)
                value = getValueFromPropDict(propDict, functionGenerated)
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


Private Function getValueFromPropDict(propDict As Scripting.Dictionary, Optional ByRef functionGenerated As Boolean) As Variant
    Dim value As Variant
    Dim Functions As VBA.Collection
    Dim functionsSet As VBA.Collection
    Dim varFunction As Variant
    Dim result As Variant
    '------------------------------------------------------------------------------------------------------

    Call F.Dictionaries.assignValueFromDictToVariable(propDict, value, TAG_VALUE)
    Call F.Dictionaries.assignValueFromDictToVariable(propDict, Functions, TAG_FUNCTIONS)

    If F.Collections.isNonEmptyCollection(Functions) Then
        functionGenerated = True
        For Each functionsSet In Functions
            For Each varFunction In functionsSet
                On Error Resume Next
                result = Application.Run(F.Reflection.getFunctionFullName(VBA.CStr(varFunction), pWorkbook))
                
                Stop
                'Dodaæ wywo³ywanie z parametrem (potrzebne do checkUserPermission
                
                If result = False Then Exit For
            Next varFunction
            If result Then Exit For
        Next functionsSet
        getValueFromPropDict = result
    Else
        getValueFromPropDict = value
    End If

End Function

Private Function isPropertyTranslatable(propName As String) As Boolean
    Static translatableProps As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    If translatableProps Is Nothing Then
        Set translatableProps = F.Dictionaries.createFrom(Array("label", "tooltip"))
    End If
    isPropertyTranslatable = translatableProps.exists(propName)
End Function





'[Icons]
Public Sub assignControlImage(controlId As String, ByRef image)
    Dim imageFileName As String
    Dim imagePath As String
    '------------------------------------------------------------------------------------------------------
    Call assignProperty(controlId, "image", imageFileName)
    If VBA.Len(imageFileName) Then
        imagePath = F.files.concatenatePath(pIconsFolderPath, imageFileName)
        Set image = LoadPictureGDI(imagePath)
    End If
End Sub





'[UPDATING]
Public Sub update(Optional eventName As String = GENERAL_INVALIDATION_GROUP)
    Dim ribbon As Variant
    Dim invalidationControls As VBA.Collection
    Dim varControlId As Variant
    '------------------------------------------------------------------------------------------------------

    If Not pUpdateDisabled Then
        Set ribbon = getRibbon
        If Not ribbon Is Nothing Then
            Set invalidationControls = F.Dictionaries.getObject(pInvalidationGroups, eventName)
            If invalidationControls Is Nothing Then
                Set invalidationControls = F.Dictionaries.getObject(pInvalidationGroups, GENERAL_INVALIDATION_GROUP)
            End If
                
            If F.Collections.isNonEmptyCollection(invalidationControls) Then
                With ribbon
                    For Each varControlId In invalidationControls
                        Call .InvalidateControl(varControlId)
                    Next varControlId
                End With
            End If
        End If
    End If

End Sub
